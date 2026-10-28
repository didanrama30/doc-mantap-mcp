"""Render and cache PDF/Office page previews as real image artifacts."""

from __future__ import annotations

import hashlib
import io
import json
import os
import platform
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

OFFICE_EXTS = {".docx", ".docm", ".doc", ".rtf", ".xlsx", ".xls", ".pptx", ".ppt",
               ".odt", ".ods", ".odp"}
WORD_EXTS = {".docx", ".docm", ".doc", ".rtf"}
MAX_PAGES = 6


@dataclass
class PreviewImage:
    data: bytes
    page: int
    path: str
    uri: str
    width: int
    height: int
    format: str


@dataclass
class PreviewResult:
    images: list[PreviewImage]
    total_pages: int | None
    note: str
    renderer_used: str

    def __iter__(self):
        """Keep the old internal (image bytes, count, note) unpacking convention working."""
        yield [image.data for image in self.images]
        yield self.total_pages
        yield self.note


def _soffice() -> str | None:
    for candidate in ("soffice", "libreoffice"):
        if found := shutil.which(candidate):
            return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


def parse_pages(spec: str, count: int, max_pages: int = MAX_PAGES) -> list[int]:
    """Parse a page expression like '1-3,5' into 0-based indexes."""
    pages: list[int] = []
    for part in (spec or "1").replace(" ", "").split(","):
        if not part:
            continue
        if part.lower() == "all":
            pages.extend(range(count))
            continue
        a, sep, b = part.partition("-")
        try:
            start = int(a) if a else 1
            end = int(b) if b else (count if sep else start)
        except ValueError:
            raise ValueError(f"Invalid pages '{spec}'. Use e.g. '1', '1-3' or '2,4'.") from None
        if start < 1 or end < start:
            raise ValueError(f"Invalid page range '{part}'. Page numbers start at 1.")
        pages.extend(range(start - 1, min(end, count)))
    seen: set[int] = set()
    ordered = [p for p in pages if p not in seen and not seen.add(p)]
    return ordered[:max(1, min(int(max_pages), 50))]


def _fingerprint(src: Path) -> str:
    stat = src.stat()
    digest = hashlib.sha256()
    with src.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    identity = f"{src}\0{stat.st_mtime_ns}\0{stat.st_size}\0{digest.hexdigest()}"
    return hashlib.sha256(identity.encode()).hexdigest()


def _cache_root() -> Path:
    root = Path(tempfile.gettempdir()) / "doc-mantap-preview-cache"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _to_pdf_libreoffice(src: Path, workdir: Path) -> Path:
    soffice = _soffice()
    if soffice is None:
        raise RuntimeError("LibreOffice/soffice not installed")
    profile = workdir / "lo-profile"
    result = subprocess.run(
        [soffice, f"-env:UserInstallation={profile.as_uri()}", "--headless",
         "--convert-to", "pdf", "--outdir", str(workdir), str(src)],
        capture_output=True, text=True, timeout=180)
    pdf = workdir / (src.stem + ".pdf")
    if result.returncode or not pdf.exists():
        details = (result.stderr or result.stdout).strip()
        raise RuntimeError(details or "LibreOffice conversion failed")
    return pdf


def _to_pdf_word(src: Path, workdir: Path) -> Path:
    """Use installed Microsoft Word via AppleScript or Windows COM."""
    output = workdir / (src.stem + ".word.pdf")
    if platform.system() == "Darwin":
        if not shutil.which("osascript"):
            raise RuntimeError("osascript unavailable")
        src_escaped = str(src).replace("\\", "\\\\").replace('"', '\\"')
        out_escaped = str(output).replace("\\", "\\\\").replace('"', '\\"')
        script = (f'tell application "Microsoft Word"\n'
                  f' set doc to open POSIX file "{src_escaped}"\n'
                  f' save as doc file name "{out_escaped}" file format format PDF\n'
                  f' close doc saving no\nend tell')
        proc = subprocess.run(["osascript", "-e", script], capture_output=True,
                              text=True, timeout=180)
        if proc.returncode or not output.is_file():
            raise RuntimeError((proc.stderr or proc.stdout).strip() or "Word export failed")
        return output
    if platform.system() == "Windows":
        ps = workdir / "word-export.ps1"
        ps.write_text(
            "$ErrorActionPreference='Stop'\n"
            "$word=$null; $doc=$null\n"
            "try { $word=New-Object -ComObject Word.Application; $word.Visible=$false; "
            "$doc=$word.Documents.Open($args[0],$false,$true); "
            "$doc.ExportAsFixedFormat($args[1],17); $doc.Close(0) } "
            "finally { if ($doc) {$doc.Close(0)}; if ($word) {$word.Quit()} }\n",
            encoding="utf-8")
        proc = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                               "-File", str(ps), str(src), str(output)],
                              capture_output=True, text=True, timeout=180)
        if proc.returncode or not output.is_file():
            raise RuntimeError((proc.stderr or proc.stdout).strip() or "Word COM export failed")
        return output
    raise RuntimeError("Microsoft Word automation supported only on macOS/Windows")


def _quicklook(src: Path, workdir: Path, width: int) -> Path:
    if platform.system() != "Darwin" or not shutil.which("qlmanage"):
        raise RuntimeError("QuickLook thumbnail unavailable on this platform")
    result = subprocess.run(["qlmanage", "-t", "-s", str(width), "-o", str(workdir), str(src)],
                            capture_output=True, text=True, timeout=90)
    candidates = sorted(workdir.glob("*.png"))
    if result.returncode or not candidates:
        raise RuntimeError((result.stderr or result.stdout).strip()
                           or "QuickLook produced no thumbnail")
    return candidates[0]


def _conversion(src: Path, cache: Path, width: int) -> tuple[Path | None, Path | None,
                                                              str, list[str]]:
    key = _fingerprint(src)
    converted = cache / "conversions" / key
    converted.mkdir(parents=True, exist_ok=True)
    pdf = converted / "converted.pdf"
    renderer_file = converted / "renderer.txt"
    errors = []
    if pdf.is_file():
        renderer = renderer_file.read_text(encoding="utf-8") if renderer_file.exists() else "cache"
        return pdf, None, renderer, errors
    quicklook_cache = converted / f"quicklook-{width}.png"
    if quicklook_cache.is_file():
        return None, quicklook_cache, "quicklook", errors
    with tempfile.TemporaryDirectory(prefix="doc-mantap-convert-") as tmp:
        workdir = Path(tmp)
        try:
            generated = _to_pdf_libreoffice(src, workdir)
            shutil.copy2(generated, pdf)
            renderer_file.write_text("libreoffice", encoding="utf-8")
            return pdf, None, "libreoffice", errors
        except (OSError, RuntimeError, subprocess.TimeoutExpired) as exc:
            errors.append(f"LibreOffice: {exc}")
        if src.suffix.lower() in WORD_EXTS:
            try:
                generated = _to_pdf_word(src, workdir)
                shutil.copy2(generated, pdf)
                renderer_file.write_text("microsoft_word", encoding="utf-8")
                return pdf, None, "microsoft_word", errors
            except (OSError, RuntimeError, subprocess.TimeoutExpired) as exc:
                errors.append(f"Microsoft Word: {exc}")
        try:
            quicklook = _quicklook(src, workdir, width)
            shutil.copy2(quicklook, quicklook_cache)
            return None, quicklook_cache, "quicklook", errors
        except (OSError, RuntimeError, subprocess.TimeoutExpired) as exc:
            errors.append(f"QuickLook: {exc}")
    return None, None, "", errors


def _actionable_error(ext: str, errors: list[str]) -> str:
    os_name = platform.system() or platform.platform()
    if os_name == "Darwin":
        install = "Install LibreOffice: `brew install --cask libreoffice`; " \
                  "or install Microsoft Word from Microsoft 365. QuickLook is built into macOS."
    elif os_name == "Linux":
        install = "Install LibreOffice: `sudo apt install libreoffice` (Debian/Ubuntu). " \
                  "Microsoft Word/QuickLook fallback unavailable on Linux."
    elif os_name == "Windows":
        install = "Install LibreOffice with `winget install TheDocumentFoundation.LibreOffice` " \
                  "or Microsoft Word desktop (COM automation)."
    else:
        install = "Install LibreOffice or use PDF input."
    detail = "\n".join(f"- {error}" for error in errors)
    return (f"Cannot preview {ext} on {os_name}. Tried LibreOffice, platform Word export, "
            f"and QuickLook where available.\n{detail}\n{install} PDFs can be rendered directly.")


def _image(data: bytes, page_num: int, path: Path, fmt: str) -> PreviewImage:
    from PIL import Image as PILImage
    with PILImage.open(io.BytesIO(data)) as image:
        width, height = image.size
    return PreviewImage(data, page_num, str(path), path.as_uri(), width, height, fmt)


def _load_cache(cache_path: Path) -> PreviewResult | None:
    meta_path = cache_path / "metadata.json"
    if not meta_path.is_file():
        return None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        images = []
        for item in meta["images"]:
            image_path = cache_path / item["file"]
            data = image_path.read_bytes()
            images.append(_image(data, item["page"], image_path, item["format"]))
        return PreviewResult(images, meta["total_pages"], meta["note"], meta["renderer_used"])
    except (OSError, KeyError, ValueError, json.JSONDecodeError):
        return None


def _write_cache(cache_path: Path, images: list[PreviewImage], total: int | None,
                 note: str, renderer: str) -> None:
    cache_path.mkdir(parents=True, exist_ok=True)
    records = []
    for image in images:
        filename = f"page-{image.page:04d}.{image.format}"
        out = cache_path / filename
        out.write_bytes(image.data)
        records.append({"page": image.page, "file": filename, "format": image.format})
    payload = {"images": records, "total_pages": total, "note": note,
               "renderer_used": renderer}
    temp = cache_path / ".metadata.tmp"
    temp.write_text(json.dumps(payload), encoding="utf-8")
    temp.replace(cache_path / "metadata.json")


def render_pages(path: str, pages: str = "1-2", width: int = 1000,
                 dpi: int | None = None, image_format: str = "png",
                 max_pages: int = MAX_PAGES) -> PreviewResult:
    """Render selected PDF/Office pages, persisting reusable image artifacts."""
    import pypdfium2 as pdfium

    src = Path(path).expanduser().resolve()
    if not src.is_file():
        raise ValueError(f"File not found: {src}")
    ext = src.suffix.lower()
    if ext != ".pdf" and ext not in OFFICE_EXTS:
        raise ValueError("preview_document supports PDF and Office document formats")
    if image_format not in ("png", "jpeg"):
        raise ValueError("format must be 'png' or 'jpeg'")
    width = max(300, min(int(width), 1800))
    dpi = max(72, min(int(dpi), 600)) if dpi is not None else None
    max_pages = max(1, min(int(max_pages), 50))
    fingerprint = _fingerprint(src)
    options = json.dumps({"pages": pages, "width": width, "dpi": dpi,
                          "format": image_format, "max_pages": max_pages}, sort_keys=True)
    key = hashlib.sha256(f"{fingerprint}\0{options}".encode()).hexdigest()
    cache_path = _cache_root() / key
    if cached := _load_cache(cache_path):
        cached.note = (cached.note + " Cached preview.").strip()
        return cached

    note = ""
    renderer = "pypdfium2"
    input_pdf = src
    if ext in OFFICE_EXTS:
        conversion_cache = _cache_root()
        converted, quicklook, renderer, errors = _conversion(src, conversion_cache, width)
        if quicklook:
            from PIL import Image as PILImage
            data = quicklook.read_bytes()
            if image_format == "jpeg":
                image = PILImage.open(io.BytesIO(data)).convert("RGB")
                buffer = io.BytesIO(); image.save(buffer, format="JPEG", quality=90)
                data = buffer.getvalue()
            image_path = cache_path / f"page-0001.{image_format}"
            cache_path.mkdir(parents=True, exist_ok=True)
            image_path.write_bytes(data)
            result_image = _image(data, 1, image_path, image_format)
            note = "QuickLook supplied a document thumbnail; page count is unknown."
            _write_cache(cache_path, [result_image], None, note, renderer)
            # Re-read so paths point to the stable cache artifact.
            return _load_cache(cache_path)
        if converted is None:
            raise RuntimeError(_actionable_error(ext, errors))
        input_pdf = converted
        note = (f"Converted with {renderer}; Microsoft Office may paginate or style content "
                "differently.")
        renderer = f"{renderer}+pypdfium2"

    document = pdfium.PdfDocument(str(input_pdf))
    try:
        count = len(document)
        selected = parse_pages(pages, count, max_pages)
        images = []
        for index in selected:
            page = document[index]
            scale = (dpi / 72) if dpi is not None else (width / page.get_width())
            rendered = page.render(scale=scale).to_pil()
            if dpi is None and rendered.width > width:
                ratio = width / rendered.width
                rendered = rendered.resize((width, max(1, round(rendered.height * ratio))))
            buffer = io.BytesIO()
            if image_format == "jpeg":
                rendered.convert("RGB").save(buffer, format="JPEG", quality=90, optimize=True)
            else:
                rendered.save(buffer, format="PNG", optimize=True)
            image_path = cache_path / f"page-{index + 1:04d}.{image_format}"
            images.append(_image(buffer.getvalue(), index + 1, image_path, image_format))
            page.close()
    finally:
        document.close()
    if not images:
        raise RuntimeError(f"No preview images rendered for {src} (0 of {count} pages selected).")
    _write_cache(cache_path, images, count, note, renderer)
    cached = _load_cache(cache_path)
    if cached is None or not cached.images:
        raise RuntimeError(f"Preview rendering failed to produce image artifacts for {src}.")
    return cached
