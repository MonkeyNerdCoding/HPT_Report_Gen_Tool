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

from app_logic import generate_report
from assets import resource_path


APP_TITLE = "OracleHC Report Generator"
PRIMARY_COLOR = "#cb0236"
PRIMARY_HOVER = "#a9022d"
SETTINGS_DIR = Path(os.getenv("APPDATA", Path.home())) / APP_TITLE
SETTINGS_FILE = SETTINGS_DIR / "settings.json"


class ReportGeneratorApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("860x680")
        self.minsize(760, 600)

        self.html_folder_var = ctk.StringVar()
        self.word_file_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Ready")
        self.last_output_file: str | None = None
        self.last_output_folder: str | None = None
        self.last_save_dir: str = ""
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self.logo_warning: str | None = None

        self._build_ui()
        self._load_settings()
        if self.logo_warning:
            self._append_log(self.logo_warning)
        self._poll_events()

    def _build_ui(self) -> None:
        self.configure(fg_color="#f7f7f8")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.grid(row=0, column=0, sticky="nsew", padx=28, pady=24)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 22))
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)

        title = ctk.CTkLabel(
            header,
            text=APP_TITLE,
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color="#1f1f24",
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ctk.CTkLabel(
            header,
            text="Select a source root and template, then choose where to save the Word report",
            font=ctk.CTkFont(size=14),
            text_color="#60616a",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(4, 0))

        logo_image = self._load_logo_image()
        if logo_image:
            logo = ctk.CTkLabel(header, text="", image=logo_image)
            logo.image = logo_image
            logo.grid(row=0, column=1, rowspan=2, sticky="e", padx=(24, 0))

        form = ctk.CTkFrame(container, fg_color="#ffffff", corner_radius=8)
        form.grid(row=1, column=0, sticky="ew", pady=(0, 18))
        form.grid_columnconfigure(1, weight=1)

        self._add_path_row(
            form,
            row=0,
            label="HTML Root Folder",
            variable=self.html_folder_var,
            browse_command=self._browse_html_folder,
        )
        self._add_path_row(
            form,
            row=1,
            label="Word File / Template",
            variable=self.word_file_var,
            browse_command=self._browse_word_file,
        )

        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        actions.grid_columnconfigure(0, weight=1)

        self.create_button = ctk.CTkButton(
            actions,
            text="Create Report",
            height=44,
            corner_radius=8,
            fg_color=PRIMARY_COLOR,
            hover_color=PRIMARY_HOVER,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._start_generation,
        )
        self.create_button.grid(row=0, column=0, sticky="ew")

        status_row = ctk.CTkFrame(actions, fg_color="transparent")
        status_row.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        status_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            status_row,
            text="Status:",
            text_color="#4b4c55",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.status_label = ctk.CTkLabel(
            status_row,
            textvariable=self.status_var,
            text_color="#4b4c55",
            font=ctk.CTkFont(size=13),
        )
        self.status_label.grid(row=0, column=1, sticky="w")

        self.progress = ctk.CTkProgressBar(status_row, mode="indeterminate", progress_color=PRIMARY_COLOR)
        self.progress.grid(row=0, column=2, sticky="e", ipadx=80)
        self.progress.set(0)

        log_frame = ctk.CTkFrame(container, fg_color="#ffffff", corner_radius=8)
        log_frame.grid(row=3, column=0, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 8))
        log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            log_header,
            text="Runtime Log",
            text_color="#1f1f24",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        self.clear_log_button = ctk.CTkButton(
            log_header,
            text="Clear Log",
            width=96,
            height=30,
            corner_radius=8,
            fg_color="#ececf0",
            hover_color="#dedee5",
            text_color="#30313a",
            command=self._clear_log,
        )
        self.clear_log_button.grid(row=0, column=1, sticky="e")

        self.log_text = ctk.CTkTextbox(
            log_frame,
            corner_radius=8,
            border_width=1,
            border_color="#dedee5",
            fg_color="#fbfbfc",
            text_color="#202128",
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="word",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        self.log_text.insert("end", "Ready.\n")
        self.log_text.configure(state="disabled")

        footer = ctk.CTkFrame(container, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        footer.grid_columnconfigure(0, weight=1)

        self.open_folder_button = ctk.CTkButton(
            footer,
            text="Open Output Folder",
            width=150,
            height=34,
            corner_radius=8,
            fg_color="#ececf0",
            hover_color="#dedee5",
            text_color="#30313a",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_button.grid(row=0, column=1, padx=(0, 10))

        self.open_file_button = ctk.CTkButton(
            footer,
            text="Open Output File",
            width=140,
            height=34,
            corner_radius=8,
            fg_color=PRIMARY_COLOR,
            hover_color=PRIMARY_HOVER,
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
    ) -> None:
        parent.grid_rowconfigure(row, weight=0)

        ctk.CTkLabel(
            parent,
            text=label,
            width=150,
            anchor="w",
            text_color="#30313a",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=row, column=0, sticky="w", padx=(16, 10), pady=12)

        entry = ctk.CTkEntry(
            parent,
            textvariable=variable,
            height=36,
            corner_radius=8,
            border_color="#d6d6de",
            fg_color="#fbfbfc",
        )
        entry.grid(row=row, column=1, sticky="ew", pady=12)

        button = ctk.CTkButton(
            parent,
            text="Browse",
            width=96,
            height=36,
            corner_radius=8,
            fg_color=PRIMARY_COLOR,
            hover_color=PRIMARY_HOVER,
            command=browse_command,
        )
        button.grid(row=row, column=2, sticky="e", padx=(10, 16), pady=12)

    def _browse_html_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select HTML Root Folder",
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

    def _start_generation(self) -> None:
        html_folder = self.html_folder_var.get().strip()
        word_file = self.word_file_var.get().strip()

        validation_error = self._validate_inputs(html_folder, word_file)
        if validation_error:
            self.status_var.set("Failed")
            messagebox.showerror(APP_TITLE, validation_error)
            self._append_log(f"Validation failed: {validation_error}")
            return

        output_file = self._ask_output_file(word_file)
        if not output_file:
            self.status_var.set("Ready")
            self._append_log("Save As canceled. No report was created.")
            return

        self.last_output_file = None
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

    def _poll_events(self) -> None:
        while True:
            try:
                event_type, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event_type == "log":
                self._append_log(payload)
            elif event_type == "success":
                self._handle_success(payload)
            elif event_type == "error":
                self._handle_error(payload)

        self.after(100, self._poll_events)

    def _handle_success(self, output_file: str) -> None:
        self.progress.stop()
        self.progress.set(0)
        self.create_button.configure(state="normal", text="Create Report")
        self.status_var.set("Success")
        self.last_output_file = output_file
        self.last_output_folder = str(Path(output_file).parent)
        self.open_file_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        self._append_log(f"Success: {output_file}")
        messagebox.showinfo(APP_TITLE, f"Report created successfully:\n{output_file}")

    def _handle_error(self, details: str) -> None:
        self.progress.stop()
        self.progress.set(0)
        self.create_button.configure(state="normal", text="Create Report")
        self.status_var.set("Failed")
        summary = details.splitlines()[0] if details else "Report generation failed."
        self._append_log("Failed.")
        self._append_log(details)
        messagebox.showerror(APP_TITLE, summary)

    def _validate_inputs(self, html_folder: str, word_file: str) -> str | None:
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
        self.last_save_dir = data.get("last_save_dir", data.get("output_folder", ""))

    def _save_settings(self) -> None:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        data = {
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
