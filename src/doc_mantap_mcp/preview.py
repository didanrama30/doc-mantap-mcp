"""Render pages of a generated file to PNG so the AI can check the layout."""

from __future__ import annotations

import io
import shutil
import subprocess
import tempfile
from pathlib import Path

OFFICE_EXTS = {".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp", ".doc", ".xls"}
MAX_PAGES = 6


def _soffice() -> str | None:
    for candidate in ("soffice", "libreoffice"):
        if found := shutil.which(candidate):
            return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


def parse_pages(spec: str, count: int) -> list[int]:
    """'1-3,5' -> [0, 1, 2, 4] (0-based, clipped to the document)."""
    pages: list[int] = []
    for part in (spec or "1").replace(" ", "").split(","):
        if not part:
            continue
        if part.lower() == "all":
            pages.extend(range(count))
            continue
        a, _, b = part.partition("-")
        try:
            start = int(a) if a else 1
            end = int(b) if b else (count if part.endswith("-") else start)
        except ValueError:
            raise ValueError(f"Invalid pages '{spec}'. Use e.g. '1', '1-3' or '2,4'.") from None
        pages.extend(range(max(start, 1) - 1, min(end, count)))
    seen, out = set(), []
    for p in pages:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out[:MAX_PAGES]


def _to_pdf(path: Path, workdir: Path) -> Path:
    soffice = _soffice()
    if soffice is None:
        raise RuntimeError(
            f"Previewing {path.suffix} files needs LibreOffice (macOS: `brew install --cask "
            "libreoffice`, Linux: `apt install libreoffice`). PDF previews work without it.")
    profile = workdir / "profile"
    subprocess.run(
        [soffice, f"-env:UserInstallation={profile.as_uri()}", "--headless",
         "--convert-to", "pdf", "--outdir", str(workdir), str(path)],
        check=True, capture_output=True, timeout=180)
    pdf = workdir / (path.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice could not convert {path.name} to PDF.")
    return pdf


def render_pages(path: str, pages: str = "1-2", width: int = 1000) -> tuple[list[bytes], int, str]:
    """Return (png images, total page count, note)."""
    import pypdfium2 as pdfium

    src = Path(path).expanduser().resolve()
    if not src.is_file():
        raise ValueError(f"File not found: {src}")
    ext = src.suffix.lower()
    if ext != ".pdf" and ext not in OFFICE_EXTS:
        raise ValueError("preview_document supports .pdf, .docx and .xlsx files.")
    width = max(300, min(int(width), 1800))
    with tempfile.TemporaryDirectory(prefix="doc-mantap-preview-") as tmp:
        note = ""
        pdf_path = src
        if ext in OFFICE_EXTS:
            pdf_path = _to_pdf(src, Path(tmp))
            note = ("Rendered with LibreOffice; Word/Excel may differ slightly (fonts, chart "
                    "styling).")
        doc = pdfium.PdfDocument(str(pdf_path))
        try:
            count = len(doc)
            images = []
            for index in parse_pages(pages, count):
                page = doc[index]
                scale = width / page.get_width()
                buf = io.BytesIO()
                page.render(scale=scale).to_pil().save(buf, format="PNG", optimize=True)
                images.append(buf.getvalue())
        finally:
            doc.close()
    return images, count, note
