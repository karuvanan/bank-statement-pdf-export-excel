from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Any

import pypdfium2 as pdfium

from scripts.parse_mbb_ocr import (
    build_statement_summaries,
    derive_signed_amounts,
    load_json,
    parse_page,
)

from .excel_export import export_workbook


Progress = Callable[[str], None]


@dataclass
class ConversionResult:
    output_path: Path
    rows: int
    statements: int
    notes: int
    debit_credit_match: bool
    work_dir: Path

    @property
    def summary(self) -> str:
        status = "matched" if self.debit_credit_match else "needs review"
        return (
            f"Output: {self.output_path}\n"
            f"Rows: {self.rows}\n"
            f"Statements: {self.statements}\n"
            f"Rows with notes: {self.notes}\n"
            f"Debit/Credit check: {status}"
        )


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def safe_stem(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "pdf"


def unique_output_path(output_dir: Path, filename: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(filename).stem or "MBB_PDF_Export"
    suffix = Path(filename).suffix or ".xlsx"
    candidate = output_dir / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_dir / f"{stem}_{timestamp}{suffix}"


def collect_pdfs(paths: Iterable[str | Path]) -> list[Path]:
    pdfs: list[Path] = []
    for item in paths:
        path = Path(item).expanduser()
        if path.is_dir():
            pdfs.extend(sorted(path.glob("*.pdf")))
            pdfs.extend(sorted(path.glob("*.PDF")))
        elif path.suffix.lower() == ".pdf":
            pdfs.append(path)
    seen: set[Path] = set()
    unique = []
    for pdf in pdfs:
        resolved = pdf.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(resolved)
    if not unique:
        raise ValueError("No PDF files found.")
    return unique


def render_pdfs(pdf_paths: list[Path], work_dir: Path, progress: Progress | None = None, scale: float = 4.0) -> Path:
    rendered_dir = work_dir / "rendered"
    rendered_dir.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, Any]] = []
    total_pages = 0

    for pdf_path in pdf_paths:
        doc = pdfium.PdfDocument(str(pdf_path))
        try:
            total_pages += len(doc)
        finally:
            doc.close()

    page_counter = 0
    for pdf_path in pdf_paths:
        doc = pdfium.PdfDocument(str(pdf_path))
        try:
            for index in range(len(doc)):
                page_counter += 1
                if progress:
                    progress(f"Render {page_counter}/{total_pages}: {pdf_path.name} page {index + 1}")
                page = doc[index]
                image = page.render(scale=scale, rotation=0).to_pil()
                image_path = rendered_dir / f"{safe_stem(pdf_path.stem)}__p{index + 1:03d}.png"
                image.save(image_path)
                manifest.append(
                    {
                        "pdf": pdf_path.name,
                        "page": index + 1,
                        "image": str(image_path),
                        "width": image.width,
                        "height": image.height,
                        "scale": scale,
                    }
                )
        finally:
            doc.close()

    manifest_path = work_dir / "render_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path


def run_windows_ocr(manifest_path: Path, ocr_dir: Path, progress: Progress | None = None) -> None:
    script = project_root() / "scripts" / "ocr_pages.ps1"
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if powershell is None:
        raise RuntimeError("PowerShell is required for Windows OCR.")
    if progress:
        progress("Running Windows OCR...")
    command = [
        powershell,
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-ManifestPath",
        str(manifest_path),
        "-OutputDir",
        str(ocr_dir),
    ]
    result = subprocess.run(
        command,
        cwd=project_root(),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if progress and result.stdout:
        progress(result.stdout.strip())
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "OCR failed.")


def parse_ocr(ocr_dir: Path, json_path: Path, progress: Progress | None = None) -> dict[str, Any]:
    if progress:
        progress("Parsing OCR output...")
    rows: list[dict[str, Any]] = []
    summaries: list[dict[str, Any]] = []
    for path in sorted(ocr_dir.glob("*.json")):
        page_rows, page_summaries = parse_page(load_json(path))
        rows.extend(page_rows)
        summaries.extend(page_summaries)

    rows = derive_signed_amounts(rows)
    statement_summaries = build_statement_summaries(rows, summaries)
    payload = {
        "rows": sorted(rows, key=lambda row: (row["source_pdf"], row["page"], row["y"])),
        "summaries": summaries,
        "statement_summaries": statement_summaries,
    }
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def debit_credit_match(summaries: list[dict[str, Any]]) -> bool:
    for item in summaries:
        debit_ok = round(item.get("calculated_debit") or 0, 2) == round(item.get("statement_total_debit") or 0, 2)
        credit_ok = round(item.get("calculated_credit") or 0, 2) == round(item.get("statement_total_credit") or 0, 2)
        if not (debit_ok and credit_ok):
            return False
    return True


def convert_pdfs(
    inputs: Iterable[str | Path],
    output_dir: str | Path | None = None,
    output_name: str = "MBB_PDF_Export_Optimized.xlsx",
    progress: Progress | None = None,
) -> ConversionResult:
    root = project_root()
    output_dir_path = Path(output_dir) if output_dir else root / "outputs" / "pdf_export_excel"
    output_path = unique_output_path(output_dir_path, output_name)
    pdf_paths = collect_pdfs(inputs)

    job_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = root / "tmp" / "gui_jobs" / job_id
    ocr_dir = work_dir / "ocr"
    work_dir.mkdir(parents=True, exist_ok=True)
    ocr_dir.mkdir(parents=True, exist_ok=True)

    if progress:
        progress(f"Found {len(pdf_paths)} PDF file(s).")
    manifest_path = render_pdfs(pdf_paths, work_dir, progress=progress)
    run_windows_ocr(manifest_path, ocr_dir, progress=progress)
    payload = parse_ocr(ocr_dir, work_dir / "extracted_transactions.json", progress=progress)

    if progress:
        progress("Building Excel workbook...")
    export_workbook(payload, output_path)

    rows = payload["rows"]
    summaries = payload["statement_summaries"]
    result = ConversionResult(
        output_path=output_path,
        rows=len(rows),
        statements=len(summaries),
        notes=sum(1 for row in rows if row.get("notes")),
        debit_credit_match=debit_credit_match(summaries),
        work_dir=work_dir,
    )
    if progress:
        progress(result.summary)
    return result


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Convert Maybank PDF statements to Excel.")
    parser.add_argument("inputs", nargs="+", help="PDF files or folders containing PDFs.")
    parser.add_argument("-o", "--output-dir", default=None, help="Output folder.")
    parser.add_argument("-n", "--output-name", default="MBB_PDF_Export_Optimized.xlsx", help="Output .xlsx filename.")
    args = parser.parse_args(argv)

    def log(message: str) -> None:
        print(message)

    try:
        result = convert_pdfs(args.inputs, args.output_dir, args.output_name, progress=log)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(result.output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
