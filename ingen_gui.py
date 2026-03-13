import os
import sys
import json
import tkinter as tk
from tkinter import messagebox, filedialog
import importlib.util
from pathlib import Path
from ingen_core import (
    generate_invitations,
    preview_receivers_text,
    format_preview_report,
    validate_custom_placeholders,
    InvitationError,
)



APP_NAME = "InGen"
BG_DARK = "#f3f4f6"
BG_PANEL = "#ffffff"
BG_INPUT = "#f8fafc"
BG_BUTTON = "#ffffff"
BG_BUTTON_ACTIVE = "#e5e7eb"
FG_MAIN = "#111827"
FG_MUTED = "#6b7280"
BORDER = "#e5e7eb"
ACCENT_YELLOW = "#f2c94c"
ACCENT_YELLOW_ACTIVE = "#e6bb3f"
PANEL_SOFT = "#f8fafc"


def resource_path(filename: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)


class InGenApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("420x420")
        self.root.minsize(360, 360)
        self.root.configure(bg=BG_DARK)

        self.logo_image = None
        self.set_window_icon()
        self._check_packaged_dependencies()
        self.build_main_menu()

    def _button_style(self) -> dict:
        return {
            "bg": "#ffffff",
            "fg": "#000000",
            "activebackground": "#e5e7eb",
            "activeforeground": "#000000",
            "disabledforeground": FG_MUTED,
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 0,
            "padx": 12,
            "pady": 7,
        }

    def _primary_button_style(self) -> dict:
        return {
            "bg": ACCENT_YELLOW,
            "fg": "#111827",
            "activebackground": ACCENT_YELLOW_ACTIVE,
            "activeforeground": "#111827",
            "disabledforeground": FG_MUTED,
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 0,
            "padx": 12,
            "pady": 7,
        }

    def _entry_style(self) -> dict:
        return {
            "bg": BG_INPUT,
            "fg": FG_MAIN,
            "insertbackground": FG_MAIN,
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 1,
            "highlightbackground": BORDER,
            "highlightcolor": "#cbd5e1",
        }

    def _text_style(self) -> dict:
        return {
            "bg": BG_INPUT,
            "fg": FG_MAIN,
            "insertbackground": FG_MAIN,
            "relief": "flat",
            "bd": 0,
            "highlightthickness": 1,
            "highlightbackground": BORDER,
            "highlightcolor": "#cbd5e1",
        }

    def _convert_docx_to_pdf(self, docx_path: str) -> str:
        """Convert DOCX to PDF with a lazy import and clear error messages."""
        try:
            from docx2pdf import convert as docx_to_pdf
        except ImportError:
            raise RuntimeError(
                "PDF export requires the 'docx2pdf' package. Install it with: pip install docx2pdf"
            )

        if not os.path.exists(docx_path):
            raise RuntimeError(f"DOCX file not found: {docx_path}")

        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

        try:
            docx_to_pdf(docx_path, pdf_path)
        except Exception as e:
            platform_name = sys.platform
            if platform_name == "darwin":
                raise RuntimeError(
                    "PDF conversion failed. On macOS, Microsoft Word must be installed for docx2pdf to work.\n"
                    f"Original error: {e}"
                )
            elif platform_name.startswith("win"):
                raise RuntimeError(
                    "PDF conversion failed. On Windows, Microsoft Word must be installed for docx2pdf to work.\n"
                    f"Original error: {e}"
                )
            else:
                raise RuntimeError(
                    "PDF conversion is only supported through Microsoft Word on Windows or macOS.\n"
                    f"Original error: {e}"
                )

        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF conversion finished, but the PDF file was not created.")

        return pdf_path

    def _load_settings(self):
        settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
        if os.path.exists(settings_path):
            try:
                with open(settings_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_settings(self, settings):
        settings_path = os.path.join(os.path.dirname(__file__), "settings.json")
        try:
            with open(settings_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass

    def _bind_mousewheel(self, canvas: tk.Canvas) -> None:
        """Bind mouse wheel only while the pointer is over a specific canvas."""
        def _on_mousewheel(event):
            if sys.platform == "darwin":
                canvas.yview_scroll(int(-event.delta), "units")
            else:
                canvas.yview_scroll(int(-event.delta / 120), "units")

        def _on_linux_scroll_up(_event):
            canvas.yview_scroll(-1, "units")

        def _on_linux_scroll_down(_event):
            canvas.yview_scroll(1, "units")

        def _bind(_event=None):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_linux_scroll_up)
            canvas.bind_all("<Button-5>", _on_linux_scroll_down)

        def _unbind(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind)
        canvas.bind("<Leave>", _unbind)

    def set_window_icon(self) -> None:
        ico_path = resource_path(os.path.join("icons", "ingen_icon.ico"))
        if os.path.exists(ico_path):
            try:
                self.root.iconbitmap(ico_path)
                return
            except Exception:
                pass

    def _check_packaged_dependencies(self) -> None:
        """Warn early if packaged package-data files required by docxcompose are missing."""
        if not getattr(sys, "frozen", False):
            return

        try:
            spec = importlib.util.find_spec("docxcompose")
            if spec is None or not spec.submodule_search_locations:
                raise FileNotFoundError("docxcompose package not found in bundled app")

            pkg_dir = Path(list(spec.submodule_search_locations)[0])
            required = pkg_dir / "templates" / "custom.xml"
            if not required.exists():
                raise FileNotFoundError(str(required))
        except Exception as e:
            messagebox.showwarning(
                "Packaging issue detected",
                "This packaged build is missing required docxcompose template files.\n\n"
                "Please rebuild the app with package data included, for example:\n"
                "py -m PyInstaller --noconfirm --windowed --name InGen "
                "--icon=icons\\ingen_icon.ico --add-data \"assets;assets\" "
                "--collect-data docxcompose ingen_gui.py\n\n"
                f"Missing file/detail: {e}"
            )

    def build_main_menu(self) -> None:
        outer_frame = tk.Frame(self.root, bg=BG_DARK)
        outer_frame.pack(fill="both", expand=True, padx=24, pady=24)

        main_frame = tk.Frame(
            outer_frame,
            bg=BG_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=8,
            pady=8,
        )
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)


        title_label = tk.Label(
            main_frame,
            text="InGen",
            font=("Segoe UI", 24, "bold"),
            bg=BG_PANEL,
            fg=FG_MAIN,
        )
        title_label.pack(pady=(8, 6))

        subtitle_label = tk.Label(
            main_frame,
            text="Invitation Generator",
            font=("Segoe UI", 11),
            fg=FG_MUTED,
            bg=BG_PANEL,
        )
        subtitle_label.pack(pady=(0, 30))

        button_frame = tk.Frame(main_frame, bg=BG_PANEL)
        button_frame.pack(pady=20, anchor="center")

        start_button = tk.Button(
            button_frame,
            text="Start",
            font=("Segoe UI", 11),
            width=20,
            height=2,
            command=self.open_placeholder_selection_window,
            cursor="hand2",
            **self._button_style(),
        )
        start_button.pack(pady=8)

        instructions_button = tk.Button(
            button_frame,
            text="Instructions",
            font=("Segoe UI", 11),
            width=20,
            height=2,
            command=self.open_instructions_window,
            cursor="hand2",
            **self._button_style(),
        )
        instructions_button.pack(pady=8)

    def open_placeholder_selection_window(self) -> None:
        self.root.withdraw()

        selection_window = tk.Toplevel(self.root)
        selection_window.title("InGen - Select Placeholders")
        selection_window.geometry("420x460")
        selection_window.configure(bg=BG_DARK)

        selected = {
            "class_name": tk.BooleanVar(value=True),
            "date": tk.BooleanVar(value=True),
            "time": tk.BooleanVar(value=True),
            "receiver": tk.BooleanVar(value=True),
            "custom": tk.BooleanVar(value=False),
        }

        outer = tk.Frame(selection_window, bg=BG_DARK)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            outer,
            bg=BG_DARK,
            highlightthickness=0,
            bd=0,
        )
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)

        main = tk.Frame(canvas, bg=BG_DARK)
        window_id = canvas.create_window((0, 0), window=main, anchor="nw")

        def _on_main_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        main.bind("<Configure>", _on_main_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        self._bind_mousewheel(canvas)

        content_card = tk.Frame(
            main,
            bg=BG_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=6,
            pady=6,
        )
        content_card.pack(fill="both", expand=True, padx=18, pady=18)

        top_bar = tk.Frame(content_card, bg=BG_PANEL)
        top_bar.pack(fill="x", pady=(20, 10), padx=20)

        tk.Button(
            top_bar,
            text="← Back",
            font=("Segoe UI", 10),
            command=lambda: self._go_back(selection_window),
            cursor="hand2",
            **self._button_style(),
        ).pack(anchor="w")


        tk.Label(
            content_card,
            text="Select placeholders",
            font=("Segoe UI", 16, "bold"),
            fg=FG_MAIN,
            bg=BG_PANEL,
        ).pack(pady=(4, 10), padx=20)

        tk.Label(
            content_card,
            text="Choose at least one placeholder before continuing.",
            font=("Segoe UI", 10),
            fg=FG_MUTED,
            bg=BG_PANEL,
        ).pack(pady=(0, 14), padx=20)

        bars_container = tk.Frame(content_card, bg=BG_PANEL)
        bars_container.pack(fill="x", pady=6, padx=20)

        def make_toggle_bar(label_text: str, variable: tk.BooleanVar):
            bar = tk.Frame(bars_container, bg=ACCENT_YELLOW if variable.get() else BG_PANEL, bd=1, relief="solid")
            bar.pack(fill="x", pady=6)

            label = tk.Label(
                bar,
                text=label_text,
                font=("Segoe UI", 11),
                fg="#111827" if variable.get() else FG_MAIN,
                bg=ACCENT_YELLOW if variable.get() else BG_PANEL,
                padx=14,
                pady=12,
                anchor="w",
            )
            label.pack(fill="x")

            def refresh():
                active = variable.get()
                bar.configure(bg=ACCENT_YELLOW if active else BG_PANEL)
                label.configure(
                    text=label_text,
                    fg="#111827" if active else FG_MAIN,
                    bg=ACCENT_YELLOW if active else BG_PANEL,
                )

            def toggle(_event=None):
                variable.set(not variable.get())
                refresh()

            bar.bind("<Button-1>", toggle)
            label.bind("<Button-1>", toggle)
            refresh()

        make_toggle_bar("Class name", selected["class_name"])
        make_toggle_bar("Date", selected["date"])
        make_toggle_bar("Time", selected["time"])
        make_toggle_bar("Receivers", selected["receiver"])
        make_toggle_bar("Add custom", selected["custom"])

        def continue_to_generator():
            chosen = [key for key, var in selected.items() if var.get()]
            if not chosen:
                messagebox.showerror("Error", "Please select at least one placeholder.")
                return

            selection_window.destroy()
            self.open_generator_window(chosen)

        tk.Button(
            content_card,
            text="Continue",
            font=("Segoe UI", 11),
            width=18,
            command=continue_to_generator,
            cursor="hand2",
            **self._primary_button_style(),
        ).pack(pady=(18, 20), padx=20)

    def open_generator_window(self, selected_placeholders=None) -> None:
        if selected_placeholders is None:
            selected_placeholders = ["class_name", "date", "time", "receiver"]
        self.root.withdraw()  # hide main menu
        generator = tk.Toplevel(self.root)
        generator.title("InGen - Generator")
        generator.geometry("640x600")
        generator.configure(bg=BG_DARK)

        outer = tk.Frame(generator, bg=BG_DARK)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            outer,
            bg=BG_DARK,
            highlightthickness=0,
            bd=0,
        )
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)

        main = tk.Frame(canvas, bg=BG_DARK)
        window_id = canvas.create_window((0, 0), window=main, anchor="nw")

        def _on_main_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        main.bind("<Configure>", _on_main_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        self._bind_mousewheel(canvas)

        content_card = tk.Frame(
            main,
            bg=BG_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=6,
            pady=6,
        )
        content_card.pack(fill="both", expand=True, padx=18, pady=18)

        # Top bar with Back button
        top_bar = tk.Frame(content_card, bg=BG_PANEL)
        top_bar.pack(fill="x", pady=(20, 10), padx=20)

        back_button = tk.Button(
            top_bar,
            text="← Back",
            font=("Segoe UI", 10),
            command=lambda: self._go_back(generator),
            cursor="hand2",
            **self._button_style(),
        )
        back_button.pack(anchor="w")


        title = tk.Label(
            content_card,
            text="Invitation Generator",
            font=("Segoe UI", 16, "bold"),
            fg=FG_MAIN,
            bg=BG_PANEL,
        )
        title.pack(pady=(0, 20), padx=20)

        selected_placeholder_vars = {
            "class_name": tk.BooleanVar(value="class_name" in selected_placeholders),
            "date": tk.BooleanVar(value="date" in selected_placeholders),
            "time": tk.BooleanVar(value="time" in selected_placeholders),
            "receiver": tk.BooleanVar(value="receiver" in selected_placeholders),
            "custom": tk.BooleanVar(value="custom" in selected_placeholders),
        }

        # --- Container for all form sections ---
        form_frame = tk.Frame(content_card, bg=BG_PANEL)
        form_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Template
        template_frame = tk.Frame(form_frame, bg=BG_PANEL)
        template_frame.pack(fill="x", pady=6)

        tk.Label(
            template_frame,
            text="Template:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        template_entry = tk.Entry(template_frame, font=("Segoe UI", 10), **self._entry_style())
        template_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Button(
            template_frame,
            text="Browse",
            font=("Segoe UI", 10),
            width=10,
            command=lambda: self._browse_template(template_entry),
            **self._button_style(),
        ).pack(side="left")

        # Mode
        mode_frame = tk.Frame(form_frame, bg=BG_PANEL)
        mode_frame.pack(fill="x", pady=10)

        tk.Label(
            mode_frame,
            text="Mode:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        mode_var = tk.StringVar(value="single")

        tk.Radiobutton(
            mode_frame,
            text="Single",
            variable=mode_var,
            value="single",
            fg=FG_MAIN,
            bg=BG_PANEL,
            activebackground=BG_PANEL,
            activeforeground=FG_MAIN,
            selectcolor=BG_PANEL,
        ).pack(side="left", padx=4)

        tk.Radiobutton(
            mode_frame,
            text="Several",
            variable=mode_var,
            value="several",
            fg=FG_MAIN,
            bg=BG_PANEL,
            activebackground=BG_PANEL,
            activeforeground=FG_MAIN,
            selectcolor=BG_PANEL,
        ).pack(side="left", padx=4)

        # Name order
        order_frame = tk.Frame(form_frame, bg=BG_PANEL)
        order_frame.pack(fill="x", pady=10)

        tk.Label(
            order_frame,
            text="Name order:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        order_var = tk.StringVar(value="name_first")

        tk.Radiobutton(
            order_frame,
            text="Name Surname",
            variable=order_var,
            value="name_first",
            fg=FG_MAIN,
            bg=BG_PANEL,
            activebackground=BG_PANEL,
            activeforeground=FG_MAIN,
            selectcolor=BG_PANEL,
        ).pack(side="left", padx=4)

        tk.Radiobutton(
            order_frame,
            text="Surname Name",
            variable=order_var,
            value="surname_first",
            fg=FG_MAIN,
            bg=BG_PANEL,
            activebackground=BG_PANEL,
            activeforeground=FG_MAIN,
            selectcolor=BG_PANEL,
        ).pack(side="left", padx=4)

        # Class name
        class_frame = tk.Frame(form_frame, bg=BG_PANEL)
        class_frame.pack(fill="x", pady=6)

        tk.Label(
            class_frame,
            text="Class:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        class_entry = tk.Entry(class_frame, font=("Segoe UI", 10), **self._entry_style())
        class_entry.pack(side="left", fill="x", expand=True)

        # Date
        date_frame = tk.Frame(form_frame, bg=BG_PANEL)
        date_frame.pack(fill="x", pady=6)

        tk.Label(
            date_frame,
            text="Date:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        date_entry = tk.Entry(date_frame, font=("Segoe UI", 10), **self._entry_style())
        date_entry.pack(side="left", fill="x", expand=True)

        # Time
        time_frame = tk.Frame(form_frame, bg=BG_PANEL)
        time_frame.pack(fill="x", pady=6)

        tk.Label(
            time_frame,
            text="Time:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        time_entry = tk.Entry(time_frame, font=("Segoe UI", 10), **self._entry_style())
        time_entry.pack(side="left", fill="x", expand=True)

        # Custom placeholders
        custom_section = tk.Frame(form_frame, bg=BG_PANEL)
        custom_section.pack(fill="x")

        placeholders_label = tk.Label(
            custom_section,
            text="Custom placeholders:",
            font=("Segoe UI", 11),
            fg=FG_MAIN,
            bg=BG_PANEL,
        )
        placeholders_label.pack(anchor="w", pady=(12, 4))

        add_placeholder_frame = tk.Frame(custom_section, bg=BG_PANEL)
        add_placeholder_frame.pack(fill="x", pady=(0, 8))

        tk.Label(
            add_placeholder_frame,
            text="Name:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=BG_PANEL,
            width=12,
            anchor="w",
        ).pack(side="left")

        placeholder_name_entry = tk.Entry(add_placeholder_frame, font=("Segoe UI", 10), **self._entry_style())
        placeholder_name_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Button(
            add_placeholder_frame,
            text="Add",
            font=("Segoe UI", 10),
            width=10,
            command=lambda: self._add_custom_placeholder_row(
                placeholders_container,
                placeholder_name_entry,
                custom_placeholder_entries,
            ),
            **self._button_style(),
        ).pack(side="left")

        placeholders_container = tk.Frame(custom_section, bg=BG_PANEL)
        placeholders_container.pack(fill="x", pady=(0, 8))

        custom_placeholder_entries = {}

        # PDF export option
        pdf_frame = tk.Frame(form_frame, bg=BG_PANEL)
        pdf_frame.pack(fill="x", pady=6)

        pdf_var = tk.BooleanVar(value=False)

        tk.Checkbutton(
            pdf_frame,
            text="Export as PDF",
            variable=pdf_var,
            fg=FG_MAIN,
            bg=BG_PANEL,
            activebackground=BG_PANEL,
            activeforeground=FG_MAIN,
            selectcolor=BG_PANEL,
        ).pack(anchor="w")

        # Receivers
        receivers_section = tk.Frame(form_frame, bg=BG_PANEL)
        receivers_section.pack(fill="both", expand=True)

        receivers_label = tk.Label(
            receivers_section,
            text="Receivers:",
            font=("Segoe UI", 11),
            fg=FG_MAIN,
            bg=BG_PANEL,
        )
        receivers_label.pack(anchor="w", pady=(16, 4))

        receivers_text = tk.Text(
            receivers_section,
            height=10,
            font=("Segoe UI", 10),
            **self._text_style(),
        )
        receivers_text.pack(fill="both", expand=True)

        # (Dynamic placeholder toggles removed, static layout is now used.)

        # Buttons
        button_frame = tk.Frame(form_frame, bg=BG_PANEL)
        button_frame.pack(pady=14)

        tk.Button(
            button_frame,
            text="Preview",
            font=("Segoe UI", 10),
            width=14,
            command=lambda: self._preview_receivers(receivers_text, order_var),
            **self._button_style(),
        ).pack(side="left", padx=8)

        tk.Button(
            button_frame,
            text="Generate",
            font=("Segoe UI", 10),
            width=14,
            command=lambda: self._generate_doc(
                template_entry,
                class_entry,
                date_entry,
                time_entry,
                receivers_text,
                order_var,
                mode_var,
                pdf_var,
                custom_placeholder_entries,
                selected_placeholder_vars,
            ),
            **self._primary_button_style(),
        ).pack(side="left", padx=8)

    def _add_custom_placeholder_row(self, container, name_entry, custom_placeholder_entries):
        raw_name = name_entry.get().strip()

        if not raw_name:
            messagebox.showerror("Error", "Please enter a custom placeholder name.")
            return

        try:
            validated = validate_custom_placeholders({raw_name: ""})
            placeholder_name = next(iter(validated.keys()))
        except InvitationError as e:
            messagebox.showerror("Error", str(e))
            return
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error:\n{e}")
            return

        if placeholder_name in custom_placeholder_entries:
            messagebox.showerror("Error", f"Placeholder '{placeholder_name}' already exists.")
            return

        row = tk.Frame(
            container,
            bg=PANEL_SOFT,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        row.pack(fill="x", pady=3)

        tk.Label(
            row,
            text=f"{placeholder_name}:",
            font=("Segoe UI", 10),
            fg=FG_MAIN,
            bg=PANEL_SOFT,
            width=12,
            anchor="w",
        ).pack(side="left")

        value_entry = tk.Entry(row, font=("Segoe UI", 10), **self._entry_style())
        value_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        usage_label = tk.Label(
            row,
            text=f"{{{{ {placeholder_name} }}}}",
            font=("Segoe UI", 9),
            fg=FG_MUTED,
            bg=PANEL_SOFT,
            padx=6,
        )
        usage_label.pack(side="left", padx=(0, 8))

        tk.Button(
            row,
            text="Remove",
            font=("Segoe UI", 9),
            width=10,
            command=lambda: self._remove_custom_placeholder_row(
                placeholder_name,
                row,
                custom_placeholder_entries,
            ),
            **self._button_style(),
        ).pack(side="left")

        custom_placeholder_entries[placeholder_name] = value_entry
        name_entry.delete(0, tk.END)

    def _remove_custom_placeholder_row(self, placeholder_name, row_widget, custom_placeholder_entries):
        if placeholder_name in custom_placeholder_entries:
            del custom_placeholder_entries[placeholder_name]
        row_widget.destroy()

    def _preview_receivers(self, receivers_text_widget, order_var):
        receivers_text = receivers_text_widget.get("1.0", tk.END).strip()
        name_order = order_var.get()

        if not receivers_text:
            messagebox.showerror("Error", "Please enter receivers.")
            return

        try:
            preview = preview_receivers_text(receivers_text, name_order=name_order)
            report = format_preview_report(preview)

            preview_window = tk.Toplevel(self.root)
            preview_window.title("InGen - Preview")
            preview_window.geometry("520x420")
            preview_window.configure(bg=BG_DARK)

            outer = tk.Frame(preview_window, bg=BG_DARK)
            outer.pack(fill="both", expand=True)

            canvas = tk.Canvas(
                outer,
                bg=BG_DARK,
                highlightthickness=0,
                bd=0,
            )
            canvas.pack(side="left", fill="both", expand=True)

            scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            scrollbar.pack(side="right", fill="y")

            canvas.configure(yscrollcommand=scrollbar.set)

            container = tk.Frame(canvas, bg=BG_DARK)
            window_id = canvas.create_window((0, 0), window=container, anchor="nw")

            def _on_container_configure(_event=None):
                canvas.configure(scrollregion=canvas.bbox("all"))

            def _on_canvas_configure(event):
                canvas.itemconfigure(window_id, width=event.width)

            container.bind("<Configure>", _on_container_configure)
            canvas.bind("<Configure>", _on_canvas_configure)
            self._bind_mousewheel(canvas)

            content_card = tk.Frame(
                container,
                bg=BG_PANEL,
                bd=0,
                highlightthickness=1,
                highlightbackground=BORDER,
                padx=6,
                pady=6,
            )
            content_card.pack(fill="both", expand=True, padx=16, pady=16)


            title = tk.Label(
                content_card,
                text="Preview",
                font=("Segoe UI", 14, "bold"),
                fg=FG_MAIN,
                bg=BG_PANEL,
            )
            title.pack(anchor="w", pady=(16, 10), padx=16)

            text_widget = tk.Text(
                content_card,
                wrap="word",
                font=("Segoe UI", 10),
                padx=12,
                pady=12,
                **self._text_style(),
            )
            text_widget.pack(fill="both", expand=True, padx=16)
            text_widget.insert("1.0", report)
            text_widget.config(state="disabled")

            close_button = tk.Button(
                content_card,
                text="Close",
                font=("Segoe UI", 10),
                width=12,
                command=preview_window.destroy,
                cursor="hand2",
                **self._button_style(),
            )
            close_button.pack(pady=(12, 16), padx=16)

        except InvitationError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error:\n{e}")

    def _generate_doc(
        self,
        template_entry,
        class_entry,
        date_entry,
        time_entry,
        receivers_text,
        order_var,
        mode_var,
        pdf_var,
        custom_placeholder_entries,
        selected_placeholder_vars,
    ):
        template_path = template_entry.get().strip()
        class_name = class_entry.get().strip()
        date_text = date_entry.get().strip()
        time_text = time_entry.get().strip()
        receivers = receivers_text.get("1.0", tk.END).strip()
        name_order = order_var.get()
        mode = mode_var.get()
        custom_placeholders = {
            name: entry.get().strip()
            for name, entry in custom_placeholder_entries.items()
        }

        selected_any = any(var.get() for var in selected_placeholder_vars.values())
        if not selected_any:
            messagebox.showerror("Error", "Please select at least one placeholder.")
            return

        if not template_path:
            messagebox.showerror("Error", "Please choose a template.")
            return

        if selected_placeholder_vars["class_name"].get() and not class_name:
            messagebox.showerror("Error", "Please enter a class name.")
            return

        if selected_placeholder_vars["date"].get() and not date_text:
            messagebox.showerror("Error", "Please enter a date.")
            return

        if selected_placeholder_vars["time"].get() and not time_text:
            messagebox.showerror("Error", "Please enter a time.")
            return

        if selected_placeholder_vars["receiver"].get() and not receivers:
            messagebox.showerror("Error", "Please enter receivers.")
            return

        # Set disabled default placeholders to safe blank values
        if not selected_placeholder_vars["class_name"].get():
            class_name = "Output"

        if not selected_placeholder_vars["date"].get():
            date_text = ""

        if not selected_placeholder_vars["time"].get():
            time_text = "00:00"

        if not selected_placeholder_vars["receiver"].get():
            receivers = "Placeholder"

        settings = self._load_settings()
        start_dir = settings.get("last_save_dir", os.path.expanduser("~/Desktop"))

        if mode == "single":
            save_path = filedialog.asksaveasfilename(
                title="Save Document",
                initialdir=start_dir,
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                initialfile=f"{class_name}.docx",
            )
        else:
            folder = filedialog.askdirectory(
                title="Choose Folder",
                initialdir=start_dir,
            )
            if not folder:
                return
            save_path = os.path.join(folder, f"{class_name}.docx")

        if not save_path:
            return

        settings["last_save_dir"] = os.path.dirname(save_path)
        self._save_settings(settings)

        try:
            result = generate_invitations(
                template_path=template_path,
                class_name=class_name,
                date_text=date_text,
                time_text=time_text,
                receivers_text=receivers,
                name_order=name_order,
                custom_placeholders=custom_placeholders,
                output_path=save_path,
                overwrite=True,
            )

            docx_path = result.output_path

            if pdf_var.get():
                try:
                    pdf_path = self._convert_docx_to_pdf(docx_path)
                    messagebox.showinfo(
                        "Success",
                        f"Documents created:\nDOCX: {docx_path}\nPDF: {pdf_path}",
                    )
                except Exception as e:
                    messagebox.showwarning(
                        "PDF Conversion Failed",
                        f"DOCX created but PDF conversion failed:\n{e}",
                    )
            else:
                messagebox.showinfo(
                    "Success",
                    f"Document created:\n{docx_path}",
                )

        except InvitationError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error:\n{e}")

    def _browse_template(self, entry_widget):
        settings = self._load_settings()
        start_dir = settings.get("last_template_dir", os.path.expanduser("~/Desktop"))

        file_path = filedialog.askopenfilename(
            title="Select Template",
            initialdir=start_dir,
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
        )

        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)

            settings["last_template_dir"] = os.path.dirname(file_path)
            self._save_settings(settings)

    def _go_back(self, window):
        window.destroy()
        self.root.deiconify()

    def open_instructions_window(self) -> None:
        instructions_window = tk.Toplevel(self.root)
        instructions_window.title("InGen - Instructions")
        instructions_window.geometry("700x560")
        instructions_window.minsize(620, 500)
        instructions_window.configure(bg=BG_DARK)

        outer = tk.Frame(instructions_window, bg=BG_DARK)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(
            outer,
            bg=BG_DARK,
            highlightthickness=0,
            bd=0,
        )
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)

        text_frame = tk.Frame(canvas, bg=BG_DARK)
        window_id = canvas.create_window((0, 0), window=text_frame, anchor="nw")

        def _on_text_frame_configure(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfigure(window_id, width=event.width)

        text_frame.bind("<Configure>", _on_text_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        self._bind_mousewheel(canvas)

        content_card = tk.Frame(
            text_frame,
            bg=BG_PANEL,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=6,
            pady=6,
        )
        content_card.pack(fill="both", expand=True, padx=16, pady=16)


        title = tk.Label(
            content_card,
            text="Instructions",
            font=("Segoe UI", 18, "bold"),
            fg=FG_MAIN,
            bg=BG_PANEL,
        )
        title.pack(anchor="w", pady=(16, 12), padx=16)

        instruction_text = (
            "Welcome to InGen — Invitation Generator.\n\n"
            "This application creates invitation documents automatically from a Microsoft Word template.\n\n"
            "HOW THE PROGRAM WORKS\n"
            "1. You choose a Word template (.docx).\n"
            "2. You enter class name, date, time and receivers.\n"
            "3. The program replaces placeholders in the template with real values.\n"
            "4. A finished Word document is generated.\n"
            "5. Optionally the document can also be converted to PDF.\n\n"
            "--------------------------------------------\n"
            "HOW TO CREATE A TEMPLATE (VERY IMPORTANT)\n"
            "--------------------------------------------\n"
            "Your template must be a normal Microsoft Word (.docx) file.\n"
            "You design the page exactly how you want the invitation to look.\n"
            "This includes:\n"
            "• logos\n"
            "• background images\n"
            "• text formatting\n"
            "• layout and alignment\n\n"
            "Wherever dynamic information should appear, you must place a placeholder.\n\n"
            "A placeholder always uses DOUBLE CURLY BRACES:\n"
            "{{ placeholder_name }}\n\n"
            "Example template sentence:\n"
            "Sahypa bilim merkeziniň diňleýjisi {{ receiver }} ...\n"
            "... {{ class_name }} sapagyna {{ date }} senesinde {{ time }} gatnaşmaga çagyrýarys.\n\n"
            "--------------------------------------------\n"
            "SUPPORTED PLACEHOLDERS\n"
            "--------------------------------------------\n"
            "{{ receiver }}    → student full name\n"
            "{{ class_name }}  → class name\n"
            "{{ date }}        → lesson date\n"
            "{{ time }}        → lesson time\n\n"
            "--------------------------------------------\n"
            "CUSTOM PLACEHOLDERS\n"
            "--------------------------------------------\n"
            "You can create your own placeholders in the generator window.\n"
            "Example:\n"
            "Placeholder name: teacher\n"
            "Template usage: {{ teacher }}\n"
            "Value in generator: Nurmuhammet Mugallym\n\n"
            "The program will replace {{ teacher }} with that value everywhere in the document.\n\n"
            "--------------------------------------------\n"
            "RECEIVER INPUT FORMATS\n"
            "--------------------------------------------\n"
            "You can enter receivers in multiple ways:\n\n"
            "One per line:\n"
            "myradow hemra\n"
            "atyaew yhlas\n"
            "yklymow dawut\n\n"
            "Space separated:\n"
            "myradow hemra atyaew yhlas yklymow dawut\n\n"
            "Comma separated:\n"
            "myradow hemra, atyaew yhlas, yklymow dawut\n\n"
            "--------------------------------------------\n"
            "IMPORTANT NOTES\n"
            "--------------------------------------------\n"
            "• The template must be .docx format.\n"
            "• Placeholder names must match exactly.\n"
            "• Do not remove the curly braces.\n"
            "• The layout of the template is preserved exactly in the output.\n"
        )

        text_widget = tk.Text(
            content_card,
            wrap="word",
            font=("Segoe UI", 10),
            padx=12,
            pady=12,
            **self._text_style(),
        )
        text_widget.pack(fill="both", expand=True, padx=16)
        text_widget.insert("1.0", instruction_text)
        text_widget.config(state="disabled")

        close_button = tk.Button(
            content_card,
            text="Close",
            font=("Segoe UI", 10),
            width=12,
            command=instructions_window.destroy,
            cursor="hand2",
            **self._button_style(),
        )
        close_button.pack(anchor="w", pady=(12, 16), padx=16)


def main() -> None:
    root = tk.Tk()
    app = InGenApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()  