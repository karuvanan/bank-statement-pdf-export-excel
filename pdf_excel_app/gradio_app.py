from __future__ import annotations

import socket
from pathlib import Path
from typing import Any

from .converter import convert_pdfs, project_root


def _uploaded_path(file_obj: Any) -> Path:
    if isinstance(file_obj, (str, Path)):
        return Path(file_obj)
    if hasattr(file_obj, "name"):
        return Path(file_obj.name)
    raise ValueError("Unsupported uploaded file object.")


def convert_uploaded(files: list[Any] | None, output_name: str) -> tuple[str | None, str]:
    if not files:
        return None, "Please upload one or more PDF files."
    messages: list[str] = []
    root = project_root()
    result = convert_pdfs(
        [_uploaded_path(file_obj) for file_obj in files],
        output_dir=root / "outputs" / "pdf_export_excel",
        output_name=output_name or "MBB_PDF_Export_Optimized.xlsx",
        progress=messages.append,
    )
    return str(result.output_path), "\n".join(messages + [result.summary])


def get_lan_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("192.0.2.1", 80))
            return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "127.0.0.1"


def find_available_port(start_port: int = 7860, host: str = "") -> int:
    for port in range(start_port, start_port + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((host, port))
            except OSError:
                continue
            return port
    raise RuntimeError("No available Gradio port found from 7860 to 7959.")


def build_app():
    try:
        import gradio as gr
    except ImportError as exc:
        missing = getattr(exc, "name", None) or "a Gradio dependency"
        raise RuntimeError(
            f"Missing Python package: {missing}. Run: py -m pip install -r requirements.txt"
        ) from exc

    with gr.Blocks(title="PDF to Excel") as demo:
        gr.Markdown("# PDF to Excel")
        pdf_files = gr.File(label="Upload PDF files", file_count="multiple", file_types=[".pdf"], type="filepath")
        output_name = gr.Textbox(label="Output filename", value="MBB_PDF_Export_Optimized.xlsx")
        convert = gr.Button("Convert")
        output_file = gr.File(label="Excel output")
        log = gr.Textbox(label="Log", lines=14)
        convert.click(convert_uploaded, inputs=[pdf_files, output_name], outputs=[output_file, log])
    return demo


def launch_web_ui(server_name: str = "0.0.0.0", server_port: int | None = None, quiet: bool = False):
    port = server_port or find_available_port()
    root = project_root()
    allowed_paths = [
        str(root),
        str(root / "outputs"),
        str(root / "tmp"),
    ]
    demo = build_app()
    demo.launch(
        server_name=server_name,
        server_port=port,
        prevent_thread_lock=True,
        inbrowser=False,
        quiet=quiet,
        show_error=True,
        allowed_paths=allowed_paths,
    )
    lan_ip = get_lan_ip()
    return demo, {
        "host": server_name,
        "port": port,
        "local_url": f"http://127.0.0.1:{port}",
        "lan_ip": lan_ip,
        "lan_url": f"http://{lan_ip}:{port}",
    }


def main() -> None:
    launch_web_ui(quiet=False)
    try:
        import time

        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
