from __future__ import annotations

import json
import os
import queue
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image

from app_logic import generate_report, run_sql_pipeline
from assets import resource_path


APP_TITLE = "OracleHC Report Generator"
PRIMARY_COLOR = "#cb0236"
PRIMARY_HOVER = "#a9022d"
PRIMARY_DARK = "#7f011f"
APP_BG = "#f4f5f7"
PANEL_BG = "#ffffff"
PANEL_ALT = "#fafbfc"
BORDER_COLOR = "#dde1e7"
BORDER_STRONG = "#c9ced8"
TEXT_PRIMARY = "#1d1f27"
TEXT_SECONDARY = "#5f6470"
TEXT_MUTED = "#7a808c"
CONTROL_BG = "#fbfcfd"
CONTROL_HOVER = "#eef1f5"
DISABLED_BG = "#e5e8ee"
DISABLED_TEXT = "#8b929f"
SETTINGS_DIR = Path(os.getenv("APPDATA", Path.home())) / APP_TITLE
SETTINGS_FILE = SETTINGS_DIR / "settings.json"


class ReportGeneratorApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("940x720")
        self.minsize(820, 640)

        self.html_folder_var = ctk.StringVar()
        self.word_file_var = ctk.StringVar()
        self.mode_var = ctk.StringVar(value="OracleHC")
        self.status_var = ctk.StringVar(value="Ready")
        self.last_output_file: str | None = None
        self.last_output_folder: str | None = None
        self.generated_report_files: list[str] = []
        self.last_save_dir: str = ""
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.logo_warning: str | None = None

        self._build_ui()
        self._load_settings()
        self._on_mode_changed()
        if self.logo_warning:
            self._append_log(self.logo_warning)
        self._poll_events()

    def _build_ui(self) -> None:
        self.configure(fg_color=APP_BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew", padx=32, pady=28)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(
            container,
            fg_color=PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18), ipady=2)
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)

        eyebrow = ctk.CTkLabel(
            header,
            text="Enterprise Reporting Workspace",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=PRIMARY_COLOR,
        )
        eyebrow.grid(row=0, column=0, sticky="w", padx=(22, 18), pady=(18, 0))

        title = ctk.CTkLabel(
            header,
            text=APP_TITLE,
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=TEXT_PRIMARY,
        )
        title.grid(row=1, column=0, sticky="w", padx=(22, 18), pady=(2, 0))

        subtitle = ctk.CTkLabel(
            header,
            text="Select the source folder and Word template, then generate a polished healthcheck report.",
            font=ctk.CTkFont(size=14),
            text_color=TEXT_SECONDARY,
        )
        subtitle.grid(row=2, column=0, sticky="w", padx=(22, 18), pady=(4, 0))

        self.mode_selector = ctk.CTkSegmentedButton(
            header,
            values=["OracleHC", "SQLHealcheck Tool"],
            variable=self.mode_var,
            selected_color=PRIMARY_COLOR,
            selected_hover_color=PRIMARY_HOVER,
            unselected_color="#f0f2f5",
            unselected_hover_color="#e5e9ef",
            text_color=TEXT_PRIMARY,
            border_width=1,
            corner_radius=8,
            height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda _value: self._on_mode_changed(),
        )
        self.mode_selector.grid(row=3, column=0, sticky="w", padx=(22, 18), pady=(16, 18))

        logo_image = self._load_logo_image()
        if logo_image:
            logo_holder = ctk.CTkFrame(header, fg_color=PANEL_ALT, corner_radius=8, border_width=1, border_color=BORDER_COLOR)
            logo_holder.grid(row=0, column=1, rowspan=4, sticky="e", padx=(20, 22), pady=18)
            logo = ctk.CTkLabel(logo_holder, text="", image=logo_image)
            logo.image = logo_image
            logo.grid(row=0, column=0, padx=18, pady=14)

        form = ctk.CTkFrame(
            container,
            fg_color=PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        form.grid(row=1, column=0, sticky="ew", pady=(0, 16), ipady=4)
        form.grid_columnconfigure(1, weight=1)

        form_heading = ctk.CTkLabel(
            form,
            text="Report Inputs",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(size=15, weight="bold"),
        )
        form_heading.grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(16, 2))

        form_caption = ctk.CTkLabel(
            form,
            text="Choose the source data and the Word template used for the final report.",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=12),
        )
        form_caption.grid(row=1, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 8))

        self.source_row_widgets = self._add_path_row(
            form,
            row=2,
            label="HTML Root Folder",
            variable=self.html_folder_var,
            browse_command=self._browse_html_folder,
        )
        self.sql_root_note = ctk.CTkLabel(
            form,
            text="Folder should contain CSV files, or DB subfolders with CSV files",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self.sql_root_note.grid(row=3, column=1, columnspan=2, sticky="w", pady=(0, 6))

        self._add_path_row(
            form,
            row=4,
            label="Word File / Template",
            variable=self.word_file_var,
            browse_command=self._browse_word_file,
        )
        self._on_mode_changed()

        actions = ctk.CTkFrame(
            container,
            fg_color=PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 16), ipady=6)
        actions.grid_columnconfigure(0, weight=1)

        self.create_button = ctk.CTkButton(
            actions,
            text="Generate Report",
            height=48,
            corner_radius=8,
            fg_color=PRIMARY_COLOR,
            hover_color=PRIMARY_HOVER,
            text_color="#ffffff",
            text_color_disabled=DISABLED_TEXT,
            border_width=1,
            border_color=PRIMARY_DARK,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._start_generation,
        )
        self.create_button.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))

        status_row = ctk.CTkFrame(actions, fg_color="transparent")
        status_row.grid(row=1, column=0, sticky="ew", padx=16, pady=(6, 10))
        status_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            status_row,
            text="Status:",
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.status_label = ctk.CTkLabel(
            status_row,
            textvariable=self.status_var,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.status_label.grid(row=0, column=1, sticky="w")

        self.progress = ctk.CTkProgressBar(
            status_row,
            mode="indeterminate",
            height=10,
            corner_radius=8,
            fg_color="#e8ebf0",
            progress_color=PRIMARY_COLOR,
        )
        self.progress.grid(row=0, column=2, sticky="e", ipadx=80)
        self.progress.set(0)

        log_frame = ctk.CTkFrame(
            container,
            fg_color=PANEL_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        log_frame.grid(row=3, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 10))
        log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            log_header,
            text="Runtime Log",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        self.clear_log_button = ctk.CTkButton(
            log_header,
            text="Clear Log",
            width=104,
            height=32,
            corner_radius=8,
            fg_color=CONTROL_BG,
            hover_color=CONTROL_HOVER,
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=BORDER_STRONG,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._clear_log,
        )
        self.clear_log_button.grid(row=0, column=1, sticky="e")

        self.log_text = ctk.CTkTextbox(
            log_frame,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
            fg_color="#f8fafc",
            text_color="#242833",
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 18))
        self.log_text.insert("end", "Ready.\n")
        self.log_text.configure(state="disabled")

        footer = ctk.CTkFrame(container, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        footer.grid_columnconfigure(0, weight=1)

        self.open_folder_button = ctk.CTkButton(
            footer,
            text="Open Output Folder",
            width=164,
            height=38,
            corner_radius=8,
            fg_color=CONTROL_BG,
            hover_color=CONTROL_HOVER,
            text_color=TEXT_PRIMARY,
            text_color_disabled=DISABLED_TEXT,
            border_width=1,
            border_color=BORDER_STRONG,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_button.grid(row=0, column=1, padx=(0, 10))

        self.open_file_button = ctk.CTkButton(
            footer,
            text="Open Output File",
            width=148,
            height=38,
            corner_radius=8,
            fg_color=PRIMARY_COLOR,
            hover_color=PRIMARY_HOVER,
            text_color="#ffffff",
            text_color_disabled=DISABLED_TEXT,
            border_width=1,
            border_color=PRIMARY_DARK,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_output_file,
            state="disabled",
        )
        self.open_file_button.grid(row=0, column=2)

    def _add_path_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: ctk.StringVar,
        browse_command,
    ) -> tuple[ctk.CTkLabel, ctk.CTkEntry, ctk.CTkButton]:
        parent.grid_rowconfigure(row, weight=0)

        label_widget = ctk.CTkLabel(
            parent,
            text=label,
            width=160,
            anchor="w",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        label_widget.grid(row=row, column=0, sticky="w", padx=(18, 12), pady=12)

        entry = ctk.CTkEntry(
            parent,
            textvariable=variable,
            height=40,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
            fg_color=CONTROL_BG,
            text_color=TEXT_PRIMARY,
            placeholder_text="No file or folder selected",
            placeholder_text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=13),
        )
        entry.grid(row=row, column=1, sticky="ew", pady=12)

        button = ctk.CTkButton(
            parent,
            text="Browse",
            width=104,
            height=40,
            corner_radius=8,
            fg_color=PRIMARY_COLOR,
            hover_color=PRIMARY_HOVER,
            text_color="#ffffff",
            border_width=1,
            border_color=PRIMARY_DARK,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=browse_command,
        )
        button.grid(row=row, column=2, sticky="e", padx=(12, 18), pady=12)
        return label_widget, entry, button

    def _browse_html_folder(self) -> None:
        title = "Select SQL Root Folder" if self._is_sql_mode() else "Select HTML Root Folder"
        path = filedialog.askdirectory(
            title=title,
            initialdir=self._initial_dir(self.html_folder_var),
        )
        if path:
            self.html_folder_var.set(path)

    def _browse_word_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Word Template",
            initialdir=self._initial_dir(self.word_file_var),
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            self.word_file_var.set(path)

    def _initial_dir(self, variable: ctk.StringVar) -> str:
        value = variable.get().strip()
        if not value:
            return str(Path.cwd())
        path = Path(value)
        if path.is_file():
            return str(path.parent)
        if path.is_dir():
            return str(path)
        return str(Path.cwd())

    def _load_logo_image(self) -> ctk.CTkImage | None:
        logo_path = resource_path("assets/tachnen_hpt.png")
        try:
            image = Image.open(logo_path)
        except Exception as exc:
            self.logo_warning = f"Warning: could not load logo asset: {logo_path} ({exc})"
            return None
        return ctk.CTkImage(light_image=image, dark_image=image, size=(120, 54))

    def _is_sql_mode(self) -> bool:
        return self.mode_var.get() == "SQLHealcheck Tool"

    def _on_mode_changed(self) -> None:
        source_label = self.source_row_widgets[0]
        if self._is_sql_mode():
            source_label.configure(text="SQL Root Folder")
            self.sql_root_note.grid()
            self.status_var.set("Ready")
            return

        source_label.configure(text="HTML Root Folder")
        self.sql_root_note.grid_remove()
        self.status_var.set("Ready")

    def _start_generation(self) -> None:
        html_folder = self.html_folder_var.get().strip()
        word_file = self.word_file_var.get().strip()

        validation_error = (
            self._validate_sql_inputs(html_folder, word_file)
            if self._is_sql_mode()
            else self._validate_oracle_inputs(html_folder, word_file)
        )
        if validation_error:
            self.status_var.set("Failed")
            messagebox.showerror(APP_TITLE, validation_error)
            self._append_log(f"Validation failed: {validation_error}")
            return

        if self._is_sql_mode():
            self._start_sql_generation(html_folder, word_file)
            return

        self._start_oracle_generation(html_folder, word_file)

    def _start_oracle_generation(self, html_folder: str, word_file: str) -> None:
        output_file = self._ask_output_file(word_file)
        if not output_file:
            self.status_var.set("Ready")
            self._append_log("Save As canceled. No report was created.")
            return

        self.last_output_file = None
        self.generated_report_files = []
        self.last_output_folder = str(Path(output_file).parent)
        self.last_save_dir = self.last_output_folder
        self._save_settings()
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.create_button.configure(state="disabled", text="Processing...")
        self.status_var.set("Processing")
        self.progress.start()
        self._append_log("")
        self._append_log("Starting report generation...")
        self._append_log(f"Output file: {output_file}")

        worker = threading.Thread(
            target=self._run_generation,
            args=(html_folder, word_file, output_file),
            daemon=True,
        )
        worker.start()

    def _start_sql_generation(self, input_folder: str, word_file: str) -> None:
        output_folder = self._ask_sql_output_folder(input_folder)
        if not output_folder:
            self.status_var.set("Ready")
            self._append_log("Output folder selection canceled. No report was created.")
            return

        self.last_output_file = None
        self.generated_report_files = []
        self.last_output_folder = str(Path(output_folder).resolve())
        self.last_save_dir = self.last_output_folder
        self._save_settings()
        self.open_file_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.create_button.configure(state="disabled", text="Processing...")
        self.status_var.set("Processing")
        self.progress.start()
        self._append_log("")
        self._append_log("Starting SQLHealcheck generation...")
        self._append_log(f"SQL root folder: {input_folder}")
        self._append_log(f"Selected output folder: {output_folder}")
        self._append_log(f"Merged Excel output: {Path(output_folder).resolve() / 'merged_healthcheck_info.xlsx'}")
        self._append_log(f"Word report output: {Path(output_folder).resolve() / 'final_healthcheck_report.docx'}")

        worker = threading.Thread(
            target=self._run_sql_generation,
            args=(input_folder, word_file, output_folder),
            daemon=True,
        )
        worker.start()

    def _ask_output_file(self, word_file: str) -> str:
        template_stem = Path(word_file).stem or "OracleHC_Report"
        initial_dir = self.last_save_dir or str(Path(word_file).parent)
        return filedialog.asksaveasfilename(
            title="Save Report As",
            initialdir=initial_dir,
            initialfile=f"{template_stem}_report.docx",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )

    def _ask_sql_output_folder(self, input_folder: str) -> str:
        initial_dir = self.last_save_dir or input_folder or str(Path.cwd())
        return filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=initial_dir,
        )

    def _run_generation(self, html_folder: str, word_file: str, output_file: str) -> None:
        try:
            output_file = generate_report(
                html_root_folder=html_folder,
                word_file=word_file,
                output_file_path=output_file,
                log_callback=lambda message: self.events.put(("log", message)),
            )
        except Exception as exc:
            self.events.put(("error", f"{exc}\n\n{traceback.format_exc()}"))
            return

        self.events.put(("success", output_file))

    def _run_sql_generation(self, input_folder: str, word_file: str, output_folder: str) -> None:
        try:
            report_files = run_sql_pipeline(
                input_root=input_folder,
                template_file=word_file,
                output_root=output_folder,
                log_callback=lambda message: self.events.put(("log", message)),
            )
        except Exception as exc:
            self.events.put(("error", f"{exc}\n\n{traceback.format_exc()}"))
            return

        self.events.put(("success", report_files))

    def _poll_events(self) -> None:
        while True:
            try:
                event_type, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event_type == "log":
                self._append_log(str(payload))
            elif event_type == "success":
                self._handle_success(payload)
            elif event_type == "error":
                self._handle_error(str(payload))

        self.after(100, self._poll_events)

    def _handle_success(self, result: object) -> None:
        self.progress.stop()
        self.progress.set(0)
        self.create_button.configure(state="normal", text="Generate Report")
        self.status_var.set("Success")
        if isinstance(result, list):
            output_files = [str(path) for path in result]
            self.generated_report_files = output_files
            docx_files = [path for path in output_files if Path(path).suffix.lower() == ".docx"]
            self.last_output_file = docx_files[0] if docx_files else (output_files[0] if len(output_files) == 1 else None)
            self.last_output_folder = str(Path(output_files[0]).parent) if output_files else self.last_output_folder
            self.open_file_button.configure(state="normal" if self.last_output_file else "disabled")
            self.open_folder_button.configure(state="normal")
            self._append_log("Success. SQLHealcheck files created:")
            for output_file in output_files:
                self._append_log(f"  {output_file}")
            messagebox.showinfo(APP_TITLE, "Created SQLHealcheck Excel and Word report.")
            return

        output_file = str(result)
        self.generated_report_files = [output_file]
        self.last_output_file = output_file
        self.last_output_folder = str(Path(output_file).parent)
        self.open_file_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        self._append_log(f"Success: {output_file}")
        messagebox.showinfo(APP_TITLE, f"Report created successfully:\n{output_file}")

    def _handle_error(self, details: str) -> None:
        self.progress.stop()
        self.progress.set(0)
        self.create_button.configure(state="normal", text="Generate Report")
        self.status_var.set("Failed")
        summary = details.splitlines()[0] if details else "Report generation failed."
        self._append_log("Failed.")
        self._append_log(details)
        messagebox.showerror(APP_TITLE, summary)

    def _validate_oracle_inputs(self, html_folder: str, word_file: str) -> str | None:
        if not html_folder or not word_file:
            return "Please select an HTML root folder and Word file/template."

        html_path = Path(html_folder)
        word_path = Path(word_file)

        if not html_path.is_dir():
            return f"HTML root folder does not exist: {html_path}"
        if not word_path.is_file():
            return f"Word file does not exist: {word_path}"
        if word_path.suffix.lower() != ".docx":
            return f"Word file must be a .docx file: {word_path}"
        return None

    def _validate_sql_inputs(self, input_folder: str, word_file: str) -> str | None:
        if not input_folder or not word_file:
            return "Please select a SQL root folder and Word file/template."

        input_path = Path(input_folder)
        word_path = Path(word_file)

        if not input_path.is_dir():
            return f"SQL root folder does not exist: {input_path}"
        if not self._has_sql_healthcheck_input(input_path):
            return f"Selected SQL root folder must contain CSV files or DB subfolders with CSV files: {input_path}"
        if not word_path.is_file():
            return f"Word file does not exist: {word_path}"
        if word_path.suffix.lower() != ".docx":
            return f"Word file must be a .docx file: {word_path}"
        return None

    def _has_sql_healthcheck_input(self, input_path: Path) -> bool:
        if any(child.is_file() and child.suffix.lower() == ".csv" for child in input_path.iterdir()):
            return True
        return any(
            child.is_dir() and any(csv_file.is_file() for csv_file in child.glob("*.csv"))
            for child in input_path.iterdir()
        )

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("end", "Ready.\n")
        self.log_text.configure(state="disabled")

    def _open_output_folder(self) -> None:
        folder = self.last_output_folder
        if folder and Path(folder).exists():
            os.startfile(folder)

    def _open_output_file(self) -> None:
        if self.last_output_file and Path(self.last_output_file).exists():
            os.startfile(self.last_output_file)

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists():
            return

        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        self.html_folder_var.set(data.get("html_folder", ""))
        self.word_file_var.set(data.get("word_file", ""))
        self.mode_var.set(data.get("mode", "OracleHC"))
        self.last_save_dir = data.get("last_save_dir", data.get("output_folder", ""))

    def _save_settings(self) -> None:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        data = {
            "mode": self.mode_var.get(),
            "html_folder": self.html_folder_var.get().strip(),
            "word_file": self.word_file_var.get().strip(),
            "last_save_dir": self.last_save_dir,
        }
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def main() -> None:
    app = ReportGeneratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
