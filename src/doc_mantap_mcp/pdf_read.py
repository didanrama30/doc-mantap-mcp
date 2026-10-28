"""Native PDF text extraction with optional Tesseract OCR for image-only pages."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


def _column_hint(textpage, page_width: float) -> bool:
    """Heuristic signal for a wide central gutter; text order remains PDFium's native order."""
    try:
        count = textpage.count_chars()
        centers = []
        for index in range(count):
            rect = textpage.get_charbox(index)
            if rect and rect[2] > rect[0]:
                centers.append((rect[0] + rect[2]) / 2)
        if len(centers) < 80:
            return False
        left = sum(x < page_width * 0.43 for x in centers)
        right = sum(x > page_width * 0.57 for x in centers)
        middle = sum(page_width * 0.43 <= x <= page_width * 0.57 for x in centers)
        return left > 20 and right > 20 and middle < min(left, right) * 0.28
    except (AttributeError, RuntimeError, TypeError):
        return False


def _ocr_page(page, language: str, tesseract: str) -> tuple[str, str | None]:
    image = page.render(scale=2.5).to_pil()
    with tempfile.TemporaryDirectory(prefix="doc-mantap-ocr-") as tmp:
        image_path = Path(tmp) / "page.png"
        image.save(image_path, format="PNG")
        result = subprocess.run(
            [tesseract, str(image_path), "stdout", "-l", language, "--psm", "3"],
            capture_output=True, text=True, timeout=90)
        if result.returncode == 0:
            return result.stdout.strip(), None
        if language != "eng":
            fallback = subprocess.run(
                [tesseract, str(image_path), "stdout", "-l", "eng", "--psm", "3"],
                capture_output=True, text=True, timeout=90)
            if fallback.returncode == 0:
                return fallback.stdout.strip(), f"OCR retried with eng; requested language '{language}' unavailable."
        return "", (result.stderr.strip() or "Tesseract OCR failed.")


def read_pdf_text(path: str, ocr: bool = True, language: str = "eng+ind") -> dict:
    """Extract native text from every page; optionally OCR pages with no usable text."""
    import pypdfium2 as pdfium

    src = Path(path).expanduser().resolve()
    if not src.is_file():
        raise ValueError(f"PDF file not found: {src}")
    if src.suffix.lower() != ".pdf":
        raise ValueError("pdf_read_text supports PDF files only")
    tesseract = shutil.which("tesseract") if ocr else None
    warnings: list[str] = []
    result_pages = []
    document = pdfium.PdfDocument(str(src))
    try:
        for index in range(len(document)):
            page = document[index]
            textpage = page.get_textpage()
            try:
                text = textpage.get_text_range().strip()
                columns = _column_hint(textpage, page.get_width())
            finally:
                textpage.close()
            is_ocr = False
            page_warning = None
            if len("".join(text.split())) < 8:
                if ocr and tesseract:
                    try:
                        text, page_warning = _ocr_page(page, language, tesseract)
                        is_ocr = bool(text)
                    except (OSError, subprocess.TimeoutExpired) as exc:
                        page_warning = f"Tesseract OCR failed: {exc}"
                elif ocr:
                    page_warning = (
                        "Page appears scanned or image-only, but Tesseract is unavailable. "
                        "Install: macOS `brew install tesseract tesseract-lang`; "
                        "Ubuntu/Debian `apt install tesseract-ocr tesseract-ocr-ind`."
                    )
                if page_warning:
                    warnings.append(f"Page {index + 1}: {page_warning}")
            result_pages.append({
                "page_number": index + 1,
                "text": text,
                "char_count": len(text),
                "is_ocr": is_ocr,
                "columns_detected": columns,
                "layout_preserved": True,
                "warning": page_warning,
            })
            page.close()
    finally:
        document.close()
    return {
        "path": str(src),
        "page_count": len(result_pages),
        "pages": result_pages,
        "text": "\n\n".join(f"--- Page {p['page_number']} ---\n{p['text']}" for p in result_pages),
        "warnings": warnings,
    }
