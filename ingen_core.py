import os
import re
import sys
import json
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime
from zipfile import ZipFile
from typing import Any, Dict, List, Optional, Tuple

from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer


# =========================
# CONSTANTS
# =========================

APP_NAME = "InGen"
TEMPLATE_NAME = "template.docx"
SETTINGS_FILE = "settings.json"
LOG_FILE = "invitation_log.txt"
DATE_FORMAT = "%d.%m.%Y"

PRIMARY_SURNAME_SUFFIXES = ("ewa", "owa", "ew", "ow")
FALLBACK_SURNAME_SUFFIXES = ("wa", "w")

REQUIRED_PLACEHOLDERS = ("receiver", "date", "time", "class_name")

RESERVED_PLACEHOLDERS = set(REQUIRED_PLACEHOLDERS)


# =========================
# CUSTOM EXCEPTIONS
# =========================

class InvitationError(Exception):
    """Base exception for invitation generator."""


class InputValidationError(InvitationError):
    """Raised when user input is invalid."""


class TemplateReadError(InvitationError):
    """Raised when template cannot be opened or read."""


class TemplateValidationError(InvitationError):
    """Raised when template placeholders are invalid or missing."""


class OutputFileExistsError(InvitationError):
    """Raised when output file exists and overwrite is not allowed."""


class MergeError(InvitationError):
    """Raised when generated documents cannot be merged."""


# =========================
# RESULT OBJECTS
# =========================

@dataclass
class PreviewResult:
    receivers: List[str]
    not_detected: List[str]
    detected_count: int
    total_count: int
    name_order: str


@dataclass
class GenerationResult:
    output_path: str
    formatted_time: str
    receivers: List[str]
    receiver_count: int
    not_detected: List[str]
    detected_count: int
    total_count: int
    class_name: str
    date_text: str
    name_order: str
    custom_placeholders: Dict[str, str]


# =========================
# SETTINGS
# =========================

def get_desktop_path() -> str:
    return os.path.join(os.path.expanduser("~"), "Desktop")


def resource_path(filename: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)


def load_settings() -> dict:
    settings_path = resource_path(SETTINGS_FILE)
    default_settings = {
        "last_template_dir": get_desktop_path(),
        "last_save_dir": get_desktop_path(),
        "last_mode": "single",
        "last_name_order": "name_first",
    }

    if not os.path.exists(settings_path):
        return default_settings

    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        default_settings.update(data)
    except Exception:
        pass

    return default_settings


def save_settings(settings: dict) -> None:
    settings_path = resource_path(SETTINGS_FILE)
    try:
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =========================
# HELPERS
# =========================

def safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def normalize_name(text: str) -> str:
    return " ".join(text.strip().split())


def open_file_windows(file_path: str) -> None:
    try:
        os.startfile(file_path)
    except OSError:
        pass


# =========================
# CUSTOM PLACEHOLDER SUPPORT
# =========================

def normalize_placeholder_name(name: str) -> str:
    return name.strip()


def validate_custom_placeholders(
    custom_placeholders: Optional[Dict[str, Any]] = None,
) -> Dict[str, str]:
    """
    Validate user-defined placeholder names and values.

    Rules:
    - names must contain only letters, numbers, and underscores
    - names cannot start with a number
    - names cannot override reserved placeholders
    - values are converted to strings
    """
    if not custom_placeholders:
        return {}

    validated: Dict[str, str] = {}

    for raw_name, raw_value in custom_placeholders.items():
        name = normalize_placeholder_name(str(raw_name))

        if not name:
            raise InputValidationError("Custom placeholder name cannot be empty.")

        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
            raise InputValidationError(
                f"Invalid custom placeholder name: {name}. Use only letters, numbers, and underscores, and do not start with a number."
            )

        if name in RESERVED_PLACEHOLDERS:
            raise InputValidationError(
                f"Custom placeholder '{name}' is reserved and cannot be overridden."
            )

        validated[name] = "" if raw_value is None else str(raw_value)

    return validated


# =========================
# VALIDATION
# =========================

def validate_date(date_text: str) -> str:
    try:
        datetime.strptime(date_text, DATE_FORMAT)
    except ValueError:
        raise InputValidationError("Date must be in DD.MM.YYYY format.")
    return date_text


def validate_time(time_text: str) -> Tuple[int, int]:
    if not re.fullmatch(r"\d{2}:\d{2}", time_text):
        raise InputValidationError("Time must be in HH:MM format.")

    hour, minute = map(int, time_text.split(":"))
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        raise InputValidationError("Invalid time.")

    return hour, minute


def time_suffix(time_text: str) -> str:
    _, minute = validate_time(time_text)

    if minute in (20, 50):
        suffix = "de"
    elif str(minute).endswith(("6", "9", "0")):
        suffix = "da"
    else:
        suffix = "de"

    return f"{time_text}-{suffix}"


# =========================
# NAME LOGIC
# =========================

def detect_surname_index(parts: List[str]) -> Optional[int]:
    for i, part in enumerate(parts):
        lower_part = part.lower()
        if lower_part.endswith(PRIMARY_SURNAME_SUFFIXES):
            return i

    for i, part in enumerate(parts):
        lower_part = part.lower()
        if lower_part.endswith(FALLBACK_SURNAME_SUFFIXES):
            return i

    return None


def format_name_with_order(full_name: str, name_order: str = "name_first") -> Tuple[str, bool]:
    """
    Returns:
        (formatted_name, detected)

    detected:
        True if surname was detected or if name_order is name_first.
        False only when surname_first is requested and no surname match was found.
    """
    full_name = normalize_name(full_name).title()
    parts = full_name.split()

    if len(parts) < 2:
        return full_name, False

    if name_order == "name_first":
        return full_name, True

    if name_order != "surname_first":
        raise InputValidationError("name_order must be 'name_first' or 'surname_first'.")

    surname_index = detect_surname_index(parts)

    if surname_index is None:
        return full_name, False

    surname = parts[surname_index]
    others = parts[:surname_index] + parts[surname_index + 1:]
    return f"{surname} {' '.join(others)}".strip(), True


def parse_receivers(text: str, name_order: str = "name_first") -> PreviewResult:
    text = text.strip()

    if not text:
        return PreviewResult(
            receivers=[],
            not_detected=[],
            detected_count=0,
            total_count=0,
            name_order=name_order,
        )

    if "," in text:
        raw_names = [name.strip() for name in text.split(",") if name.strip()]
    else:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if len(lines) > 1:
            raw_names = lines
        else:
            words = text.split()
            if len(words) % 2 != 0:
                raise InputValidationError(
                    "One-line input without commas must contain full name pairs.\n"
                    "Example:\n"
                    "Ahmet Atayew Kemal Hojamgulyyew"
                )
            raw_names = []
            for i in range(0, len(words), 2):
                raw_names.append(words[i] + " " + words[i + 1])

    formatted_receivers: List[str] = []
    not_detected: List[str] = []
    detected_count = 0

    for raw_name in raw_names:

        clean_name = normalize_name(raw_name).title()

        formatted_name, detected = format_name_with_order(clean_name, name_order=name_order)

        formatted_receivers.append(formatted_name)

        if name_order == "surname_first":
            if detected:
                detected_count += 1
            else:
                not_detected.append(clean_name)
        else:
            detected_count += 1

    return PreviewResult(
        receivers=formatted_receivers,
        not_detected=not_detected,
        detected_count=detected_count,
        total_count=len(formatted_receivers),
        name_order=name_order,
    )


# =========================
# TEMPLATE
# =========================

def validate_template(template_path: str) -> None:
    if not template_path:
        raise TemplateReadError("Template path is required.")

    if not os.path.exists(template_path):
        raise TemplateReadError(f"Template not found: {template_path}")

    try:
        with ZipFile(template_path) as docx:
            xml = ""
            for file in docx.namelist():
                if file.endswith(".xml"):
                    xml += docx.read(file).decode(errors="ignore")
    except Exception as e:
        raise TemplateReadError(f"Template cannot be read: {e}")

    missing = []
    for placeholder in REQUIRED_PLACEHOLDERS:
        if f"{{{{ {placeholder} }}}}" not in xml and f"{{{{{placeholder}}}}}" not in xml:
            missing.append(placeholder)

    if missing:
        raise TemplateValidationError(
            "Template missing placeholders: " + ", ".join(missing)
        )


# =========================
# OUTPUT PATH
# =========================

def prepare_output_path(
    class_name: str,
    output_path: Optional[str] = None,
    output_dir: Optional[str] = None,
) -> str:
    filename = safe_filename(class_name) + ".docx"

    if output_path:
        return os.path.abspath(output_path)

    if output_dir:
        return os.path.abspath(os.path.join(output_dir, filename))

    return os.path.abspath(filename)


# =========================
# DOC GENERATION
# =========================

def render_receiver_doc(template_path: str, context: dict, temp_dir: str, index: int) -> str:
    tpl = DocxTemplate(template_path)
    tpl.render(context)

    output_path = os.path.join(temp_dir, f"temp_{index}.docx")
    tpl.save(output_path)

    if not os.path.exists(output_path):
        raise MergeError(f"Temporary page was not created: {output_path}")

    return output_path


def merge_docs(doc_paths: List[str], output_path: str) -> None:
    if not doc_paths:
        raise MergeError("No generated documents to merge.")

    try:
        master = Document(doc_paths[0])
        composer = Composer(master)

        for path in doc_paths[1:]:
            composer.append(Document(path))

        composer.save(output_path)
    except Exception as e:
        raise MergeError(f"Could not merge documents: {e}")

    if not os.path.exists(output_path):
        raise MergeError("Merged output file was not created.")


def build_output(
    template_path: str,
    class_name: str,
    date_text: str,
    time_text: str,
    receivers: List[str],
    custom_placeholders: Optional[Dict[str, str]] = None,
    output_path: Optional[str] = None,
    output_dir: Optional[str] = None,
    overwrite: bool = False,
) -> Tuple[str, str]:
    formatted_time = time_suffix(time_text)
    validated_custom_placeholders = validate_custom_placeholders(custom_placeholders)
    final_output_path = prepare_output_path(class_name, output_path=output_path, output_dir=output_dir)

    if os.path.exists(final_output_path) and not overwrite:
        raise OutputFileExistsError(f"File already exists: {final_output_path}")

    temp_dir = tempfile.mkdtemp(prefix="inv_builder_")

    try:
        doc_paths = []

        for i, receiver in enumerate(receivers, start=1):
            context = {
                "receiver": receiver,
                "date": date_text,
                "time": formatted_time,
                "class_name": class_name,
                **validated_custom_placeholders,
            }
            doc_path = render_receiver_doc(template_path, context, temp_dir, i)
            doc_paths.append(doc_path)

        merge_docs(doc_paths, final_output_path)
        return final_output_path, formatted_time

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# =========================
# LOGGING
# =========================

def write_log(
    class_name: str,
    date_text: str,
    time_text: str,
    receivers: List[str],
    output_path: str,
    not_detected: List[str],
    name_order: str,
    template_path: str,
) -> None:
    log_path = resource_path(LOG_FILE)

    line = (
        f"{datetime.now()} | "
        f"class={class_name} | "
        f"date={date_text} | "
        f"time={time_text} | "
        f"name_order={name_order} | "
        f"receivers={len(receivers)} | "
        f"not_detected={len(not_detected)} | "
        f"template={template_path} | "
        f"output={output_path}\n"
    )

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


# =========================
# PUBLIC CORE API
# =========================


def preview_receivers_text(receivers_text: str, name_order: str = "name_first") -> PreviewResult:
    result = parse_receivers(receivers_text, name_order=name_order)
    if not result.receivers:
        raise InputValidationError("No receivers entered.")
    return result


# =========================
# PREVIEW/REPORT HELPERS
# =========================

def format_preview_report(preview: PreviewResult) -> str:
    """
    Creates a readable preview text.
    Useful for GUI preview windows or terminal display.
    """
    lines = []

    lines.append(f"Detected: {preview.detected_count} / {preview.total_count}")
    lines.append("")

    lines.append("Parsed receivers:")
    for i, name in enumerate(preview.receivers, start=1):
        lines.append(f"{i}. {name}")

    if preview.not_detected:
        lines.append("")
        lines.append("Not detected:")
        for name in preview.not_detected:
            lines.append(f"- {name}")

    return "\n".join(lines)


def find_duplicate_receivers(receivers: List[str]) -> List[str]:
    """
    Returns a list of duplicate receiver names without removing them.
    """
    seen = set()
    duplicates = []

    for r in receivers:
        if r in seen and r not in duplicates:
            duplicates.append(r)
        seen.add(r)

    return duplicates


def generate_invitations(
    template_path: str,
    class_name: str,
    date_text: str,
    time_text: str,
    receivers_text: str,
    name_order: str = "name_first",
    custom_placeholders: Optional[Dict[str, Any]] = None,
    output_path: Optional[str] = None,
    output_dir: Optional[str] = None,
    overwrite: bool = False,
    open_after_save: bool = False,
) -> GenerationResult:
    class_name = class_name.strip().title()
    if not class_name:
        raise InputValidationError("Class name cannot be empty.")

    date_text = validate_date(date_text.strip())
    validate_time(time_text.strip())
    validate_template(template_path)

    preview = preview_receivers_text(receivers_text, name_order=name_order)
    validated_custom_placeholders = validate_custom_placeholders(custom_placeholders)

    final_output_path, formatted_time = build_output(
        template_path=template_path,
        class_name=class_name,
        date_text=date_text,
        time_text=time_text.strip(),
        receivers=preview.receivers,
        custom_placeholders=validated_custom_placeholders,
        output_path=output_path,
        output_dir=output_dir,
        overwrite=overwrite,
    )

    write_log(
        class_name=class_name,
        date_text=date_text,
        time_text=formatted_time,
        receivers=preview.receivers,
        output_path=final_output_path,
        not_detected=preview.not_detected,
        name_order=name_order,
        template_path=template_path,
    )

    if open_after_save:
        open_file_windows(final_output_path)

    return GenerationResult(
        output_path=final_output_path,
        formatted_time=formatted_time,
        receivers=preview.receivers,
        receiver_count=len(preview.receivers),
        not_detected=preview.not_detected,
        detected_count=preview.detected_count,
        total_count=preview.total_count,
        class_name=class_name,
        date_text=date_text,
        name_order=name_order,
        custom_placeholders=validated_custom_placeholders,
    )


# =========================
# SELF TESTS
# =========================

def run_self_tests() -> None:
    assert format_name_with_order("Ahmet Atayew", "surname_first")[0] == "Atayew Ahmet"
    assert format_name_with_order("Ahmet Atayew", "name_first")[0] == "Ahmet Atayew"
    assert time_suffix("15:00") == "15:00-da"
    assert time_suffix("15:20") == "15:20-de"
    assert validate_custom_placeholders({"teacher": "Ayna"}) == {"teacher": "Ayna"}
    try:
        validate_custom_placeholders({"receiver": "Override"})
        raise AssertionError("Reserved placeholder validation failed")
    except InputValidationError:
        pass


# =========================
# OPTIONAL TERMINAL DEMO
# =========================

def main():
    run_self_tests()

    settings = load_settings()
    template_dir = settings.get("last_template_dir", get_desktop_path())
    default_template = os.path.join(template_dir, TEMPLATE_NAME)

    template_path = input(f"Template path [{default_template}]: ").strip().strip('\"') or default_template
    template_path = os.path.abspath(template_path)
    class_name = input("Enter class name: ").strip()
    date_text = input("Enter date (DD.MM.YYYY): ").strip()
    time_text = input("Enter time (HH:MM): ").strip()

    print("\nChoose name order:")
    print("1. Name Surname (default)")
    print("2. Surname Name")
    order_choice = input("Enter 1 or 2: ").strip()
    name_order = "surname_first" if order_choice == "2" else "name_first"

    print("\nEnter receivers (END to finish):")
    lines = []
    while True:
        line = input()
        if line.strip().upper() == "END":
            break
        lines.append(line)

    receivers_text = "\n".join(lines)

    try:
        preview = preview_receivers_text(receivers_text, name_order=name_order)

        print("\nParsed receivers:")
        for i, r in enumerate(preview.receivers, start=1):
            print(f"{i}. {r}")

        print(f"\nDetected: {preview.detected_count}/{preview.total_count}")

        if preview.not_detected:
            print("\nNot detected:")
            for item in preview.not_detected:
                print(f"- {item}")

        result = generate_invitations(
            template_path=template_path,
            class_name=class_name,
            date_text=date_text,
            time_text=time_text,
            receivers_text=receivers_text,
            name_order=name_order,
            overwrite=True,
            open_after_save=False,
        )

        print("\nCreated successfully.")
        print("Saved at:", result.output_path)

        settings["last_template_dir"] = os.path.dirname(template_path)
        settings["last_save_dir"] = os.path.dirname(result.output_path)
        settings["last_name_order"] = name_order
        save_settings(settings)

    except InvitationError as e:
        print("\nError:", e)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()