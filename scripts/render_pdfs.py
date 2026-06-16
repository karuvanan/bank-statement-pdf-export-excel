from __future__ import annotations

import json
import re
from pathlib import Path

import pypdfium2 as pdfium


ROOT = Path.cwd()
RENDER_DIR = ROOT / "tmp" / "pdfs" / "rendered"
MANIFEST_PATH = ROOT / "tmp" / "pdfs" / "render_manifest.json"
SCALE = 4.0


def safe_stem(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_")


def main() -> None:
    RENDER_DIR.mkdir(parents=True, exist_ok=True)
    manifest = []

    for pdf_path in sorted(ROOT.glob("*.pdf")):
        doc = pdfium.PdfDocument(str(pdf_path))
        try:
            for index in range(len(doc)):
                page = doc[index]
                bitmap = page.render(scale=SCALE, rotation=0)
                image = bitmap.to_pil()
                out_path = RENDER_DIR / f"{safe_stem(pdf_path.stem)}__p{index + 1:03d}.png"
                image.save(out_path)
                manifest.append(
                    {
                        "pdf": pdf_path.name,
                        "page": index + 1,
                        "image": str(out_path.relative_to(ROOT)),
                        "width": image.width,
                        "height": image.height,
                        "scale": SCALE,
                    }
                )
        finally:
            doc.close()

    MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Rendered {len(manifest)} pages from {len(list(ROOT.glob('*.pdf')))} PDF files.")
    print(MANIFEST_PATH)


if __name__ == "__main__":
    main()
