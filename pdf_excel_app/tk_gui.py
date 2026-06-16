from __future__ import annotations

import contextlib
import logging
import queue
import threading
import tkinter as tk
import traceback
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .converter import convert_pdfs, project_root


class QueueTextWriter:
    encoding = "utf-8"

    def __init__(self, callback) -> None:
        self.callback = callback
        self._buffer = ""

    def write(self, text: str) -> int:
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self.callback(line.rstrip())
        return len(text)

    def flush(self) -> None:
        if self._buffer.strip():
            self.callback(self._buffer.rstrip())
        self._buffer = ""

    def isatty(self) -> bool:
        return False

    def writable(self) -> bool:
        return True


class QueueLogHandler(logging.Handler):
    def __init__(self, callback) -> None:
        super().__init__()
        self.callback = callback
        self.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    def emit(self, record: logging.LogRecord) -> None:
        self.callback(self.format(record))


class PdfExcelApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF to Excel")
        self.geometry("1100x680")
        self.minsize(920, 620)
        self.selected_pdfs: list[Path] = []
        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.gradio_demo = None
        self.web_url = ""
        self.web_log_handler: QueueLogHandler | None = None
        self.root_dir = project_root()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(120, self._drain_log_queue)
        self.after(300, self.start_web_ui)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        top = ttk.Frame(self, padding=16)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="PDF files").grid(row=0, column=0, sticky="w")
        self.pdf_label = ttk.Label(top, text="No files selected")
        self.pdf_label.grid(row=0, column=1, sticky="ew", padx=10)
        ttk.Button(top, text="Select PDFs", command=self.select_pdfs).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(top, text="Select Folder", command=self.select_folder).grid(row=0, column=3)

        ttk.Label(top, text="Output folder").grid(row=1, column=0, sticky="w", pady=(12, 0))
        self.output_dir_var = tk.StringVar(value=str(self.root_dir / "outputs" / "pdf_export_excel"))
        ttk.Entry(top, textvariable=self.output_dir_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=10, pady=(12, 0))
        ttk.Button(top, text="Browse", command=self.select_output_dir).grid(row=1, column=3, pady=(12, 0))

        ttk.Label(top, text="Output file").grid(row=2, column=0, sticky="w", pady=(12, 0))
        self.output_name_var = tk.StringVar(value="MBB_PDF_Export_Optimized.xlsx")
        ttk.Entry(top, textvariable=self.output_name_var).grid(row=2, column=1, columnspan=3, sticky="ew", padx=10, pady=(12, 0))

        ttk.Label(top, text="Web UI").grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.web_status_var = tk.StringVar(value="Starting Gradio Web UI...")
        ttk.Label(top, textvariable=self.web_status_var).grid(row=3, column=1, columnspan=2, sticky="ew", padx=10, pady=(12, 0))
        self.open_web_button = ttk.Button(top, text="Open Web UI", command=self.open_web_ui, state="disabled")
        self.open_web_button.grid(row=3, column=3, pady=(12, 0))

        actions = ttk.Frame(self, padding=(16, 0, 16, 8))
        actions.grid(row=1, column=0, sticky="ew")
        actions.columnconfigure(1, weight=1)
        self.convert_button = ttk.Button(actions, text="Convert to Excel", command=self.start_conversion)
        self.convert_button.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(actions, mode="indeterminate")
        self.progress.grid(row=0, column=1, sticky="ew", padx=12)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var, padding=(16, 0, 16, 8)).grid(row=2, column=0, sticky="ew")

        log_shell = ttk.Frame(self, padding=(16, 0, 16, 16))
        log_shell.grid(row=3, column=0, sticky="nsew")
        log_shell.rowconfigure(0, weight=1)
        log_shell.columnconfigure(0, weight=1)
        panes = ttk.Panedwindow(log_shell, orient=tk.HORIZONTAL)
        panes.grid(row=0, column=0, sticky="nsew")

        conversion_frame = ttk.LabelFrame(panes, text="Conversion Log", padding=8)
        conversion_frame.rowconfigure(0, weight=1)
        conversion_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(conversion_frame, wrap="word", height=16)
        self.log.grid(row=0, column=0, sticky="nsew")
        conversion_scrollbar = ttk.Scrollbar(conversion_frame, orient="vertical", command=self.log.yview)
        conversion_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=conversion_scrollbar.set)

        web_frame = ttk.LabelFrame(panes, text="Web UI Backend Log", padding=8)
        web_frame.rowconfigure(0, weight=1)
        web_frame.columnconfigure(0, weight=1)
        self.web_log = tk.Text(web_frame, wrap="word", height=16)
        self.web_log.grid(row=0, column=0, sticky="nsew")
        web_scrollbar = ttk.Scrollbar(web_frame, orient="vertical", command=self.web_log.yview)
        web_scrollbar.grid(row=0, column=1, sticky="ns")
        self.web_log.configure(yscrollcommand=web_scrollbar.set)

        panes.add(conversion_frame, weight=1)
        panes.add(web_frame, weight=1)

    def start_web_ui(self) -> None:
        self._append_web_log("Starting Gradio Web UI in the background...")
        try:
            from .gradio_app import launch_web_ui

            self._attach_web_logger()
            writer = QueueTextWriter(lambda message: self.log_queue.put(("web_log", message)))
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                self.gradio_demo, info = launch_web_ui(quiet=False)
            writer.flush()
        except Exception as exc:
            message = str(exc)
            self.log_queue.put(("web_error", message))
            self.log_queue.put(("web_log", traceback.format_exc()))
            return

        local_url = info["local_url"]
        lan_url = info["lan_url"]
        self.web_url = local_url
        self.log_queue.put(("web_ready", f"IP: {info['lan_ip']} | Port: {info['port']} | Local: {local_url} | LAN: {lan_url}"))
        self.log_queue.put(("web_log", f"Web UI started at {local_url}"))
        if lan_url != local_url:
            self.log_queue.put(("web_log", f"LAN address: {lan_url}"))

    def _attach_web_logger(self) -> None:
        if self.web_log_handler is not None:
            return
        self.web_log_handler = QueueLogHandler(lambda message: self.log_queue.put(("web_log", message)))
        for logger_name in ("gradio", "uvicorn", "uvicorn.error", "uvicorn.access"):
            logger = logging.getLogger(logger_name)
            logger.addHandler(self.web_log_handler)
            logger.setLevel(logging.INFO)

    def open_web_ui(self) -> None:
        if self.web_url:
            webbrowser.open(self.web_url)

    def select_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.selected_pdfs = [Path(path) for path in paths]
            self.pdf_label.configure(text=f"{len(self.selected_pdfs)} file(s) selected")

    def select_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if folder:
            self.selected_pdfs = [Path(folder)]
            self.pdf_label.configure(text=str(folder))

    def select_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_dir_var.set(folder)

    def start_conversion(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        if not self.selected_pdfs:
            messagebox.showwarning("Missing PDFs", "Please select PDF files or a folder.")
            return
        self.convert_button.configure(state="disabled")
        self.progress.start(10)
        self.status_var.set("Converting...")
        self.log.delete("1.0", tk.END)
        self.worker = threading.Thread(target=self._convert_worker, daemon=True)
        self.worker.start()

    def _convert_worker(self) -> None:
        try:
            result = convert_pdfs(
                self.selected_pdfs,
                output_dir=self.output_dir_var.get(),
                output_name=self.output_name_var.get(),
                progress=lambda message: self.log_queue.put(("log", message)),
            )
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))
        else:
            self.log_queue.put(("done", result.summary))

    def _drain_log_queue(self) -> None:
        try:
            while True:
                kind, message = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(message)
                elif kind == "web_log":
                    self._append_web_log(message)
                elif kind == "web_ready":
                    self.web_status_var.set(f"Web UI running - {message}")
                    self.open_web_button.configure(state="normal")
                    self._append_web_log(message)
                elif kind == "web_error":
                    self.web_status_var.set("Web UI failed to start")
                    self._append_web_log(f"Error: {message}")
                elif kind == "done":
                    self._append_log(message)
                    self.status_var.set("Done")
                    self.progress.stop()
                    self.convert_button.configure(state="normal")
                    messagebox.showinfo("Completed", message)
                elif kind == "error":
                    self._append_log(f"Error: {message}")
                    self.status_var.set("Error")
                    self.progress.stop()
                    self.convert_button.configure(state="normal")
                    messagebox.showerror("Conversion failed", message)
        except queue.Empty:
            pass
        self.after(120, self._drain_log_queue)

    def _append_log(self, message: str) -> None:
        self.log.insert(tk.END, f"{message}\n")
        self.log.see(tk.END)

    def _append_web_log(self, message: str) -> None:
        self.web_log.insert(tk.END, f"{message}\n")
        self.web_log.see(tk.END)

    def on_close(self) -> None:
        if self.gradio_demo is not None:
            try:
                self.gradio_demo.close()
            except Exception:
                pass
        self.destroy()


def main() -> None:
    app = PdfExcelApp()
    app.mainloop()


if __name__ == "__main__":
    main()
