"""Environment and optional dependency diagnostics for doc-mantap-mcp."""

from __future__ import annotations

import ctypes.util
import importlib.util
import os
import platform
import shutil
import subprocess
from pathlib import Path


def _install_hints(system: str) -> dict[str, str]:
    if system == "Darwin":
        return {
            "python": "Install project dependencies with `uv sync`.",
            "libreoffice": "brew install --cask libreoffice",
            "tesseract": "brew install tesseract tesseract-lang",
            "pango": "brew install pango",
            "fonts": "Install fonts with Font Book, or use bundled fonts for PDF output.",
            "word": "Install Microsoft Word from Microsoft 365; QuickLook is built into macOS.",
        }
    if system == "Windows":
        return {
            "python": "Install project dependencies with `uv sync`.",
            "libreoffice": "winget install TheDocumentFoundation.LibreOffice",
            "tesseract": "winget install UB-Mannheim.TesseractOCR",
            "pango": "See https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#windows",
            "fonts": "Install fonts with Windows Settings > Personalization > Fonts.",
            "word": "Install Microsoft Word desktop for COM export.",
        }
    return {
        "python": "Install project dependencies with `uv sync`.",
        "libreoffice": "sudo apt install libreoffice",
        "tesseract": "sudo apt install tesseract-ocr tesseract-ocr-ind",
        "pango": "sudo apt install libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz-subset0",
        "fonts": "sudo apt install fonts-liberation fonts-dejavu-core",
        "word": "Microsoft Word/QuickLook automation is unavailable on Linux.",
    }


def _check(name: str, available: bool, details: str = "", install: str = "") -> dict:
    return {"name": name, "available": bool(available), "status": "ok" if available else "missing",
            "details": details, "install": "" if available else install}


def _tesseract_languages(binary: str | None) -> tuple[list[str], str]:
    if not binary:
        return [], ""
    try:
        result = subprocess.run([binary, "--list-langs"], capture_output=True, text=True,
                                timeout=10)
        output = result.stdout + "\n" + result.stderr
        languages = [line.strip() for line in output.splitlines()
                     if line.strip() and not line.strip().lower().startswith(
                         ("list of available", "usage:"))]
        return languages, ""
    except (OSError, subprocess.TimeoutExpired) as exc:
        return [], str(exc)


def environment_report() -> dict:
    """Return required/optional dependency availability and actionable install hints."""
    system = platform.system()
    install = _install_hints(system)
    package_modules = {
        "python-docx": "docx", "pypdfium2": "pypdfium2", "Pillow": "PIL",
        "WeasyPrint": "weasyprint", "MCP SDK": "mcp", "Pygments": "pygments",
        "XlsxWriter": "xlsxwriter", "Matplotlib": "matplotlib", "PyYAML": "yaml",
    }
    packages = []
    for label, module in package_modules.items():
        available = importlib.util.find_spec(module) is not None
        packages.append(_check(label, available, f"Python module: {module}", install["python"]))

    binaries = {name: shutil.which(name) for name in ("soffice", "libreoffice", "tesseract",
                                                       "qlmanage", "osascript", "fc-list")}
    office = binaries["soffice"] or binaries["libreoffice"]
    dependencies = [
        *packages,
        _check("LibreOffice", bool(office), office or "soffice/libreoffice not found",
               install["libreoffice"]),
        _check("Tesseract", bool(binaries["tesseract"]), binaries["tesseract"] or "not found",
               install["tesseract"]),
        _check("Pango", bool(ctypes.util.find_library("pango-1.0")
                             or ctypes.util.find_library("libpango-1.0-0")),
               "Native text layout library used by PDF generation", install["pango"]),
    ]
    languages, language_error = _tesseract_languages(binaries["tesseract"])
    for language in ("eng", "ind"):
        dependencies.append(_check(f"Tesseract language: {language}", language in languages,
                                   language_error or (", ".join(languages) if languages else ""),
                                   install["tesseract"]))

    if system == "Darwin":
        word_available = False
        if binaries["osascript"]:
            try:
                result = subprocess.run([binaries["osascript"], "-e",
                                         'exists application "Microsoft Word"'],
                                        capture_output=True, text=True, timeout=10)
                word_available = result.returncode == 0 and "true" in result.stdout.lower()
            except (OSError, subprocess.TimeoutExpired):
                pass
        dependencies.append(_check("Microsoft Word automation", word_available,
                                   "AppleScript availability", install["word"]))
        dependencies.append(_check("QuickLook", bool(binaries["qlmanage"]),
                                   binaries["qlmanage"] or "not found", install["word"]))
    elif system == "Windows":
        office_dirs = [Path(os.environ.get("ProgramFiles", "C:/Program Files")),
                       Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)"))]
        candidates = [directory / "Microsoft Office/root/Office16/WINWORD.EXE"
                      for directory in office_dirs]
        word_available = bool(shutil.which("WINWORD.EXE") or shutil.which("winword")
                              or any(candidate.is_file() for candidate in candidates))
        dependencies.append(_check("Microsoft Word automation", word_available,
                                   "Word desktop COM automation", install["word"]))

    font_dir = Path(__file__).parent / "assets" / "fonts"
    bundled = sorted(p.name for p in font_dir.iterdir() if p.is_dir()) if font_dir.is_dir() else []
    dependencies.append(_check("Bundled fonts", bool(bundled), ", ".join(bundled),
                               install["fonts"]))
    system_fonts = (Path("/System/Library/Fonts").is_dir() if system == "Darwin" else
                    Path("C:/Windows/Fonts").is_dir() if system == "Windows" else
                    bool(binaries["fc-list"]))
    dependencies.append(_check("System fonts", system_fonts, "Platform font directory/listing",
                               install["fonts"]))

    status = {item["name"]: item["available"] for item in dependencies}
    tool_requirements = {
        "markdown_to_pdf": {"requires": ["MCP SDK", "WeasyPrint", "Pango", "Pillow",
                                           "Matplotlib", "PyYAML"]},
        "markdown_to_docx": {"requires": ["MCP SDK", "python-docx", "Pillow", "PyYAML"]},
        "create_excel": {"requires": ["MCP SDK", "XlsxWriter", "Matplotlib", "PyYAML"]},
        "preview_document": {"requires": ["pypdfium2", "Pillow"],
                             "optional": ["LibreOffice", "Microsoft Word automation",
                                          "QuickLook"]},
        "pdf_read_text": {"requires": ["pypdfium2"],
                          "optional": ["Tesseract", "Tesseract language: eng",
                                       "Tesseract language: ind"]},
        "docx_read/docx_patch/docx_insert_image": {"requires": ["python-docx", "Pillow"]},
    }
    for tool in tool_requirements.values():
        requirements = tool["requires"]
        tool["ready"] = all(status.get(name, False) for name in requirements)
        tool["missing"] = [name for name in requirements if not status.get(name, False)]
    return {
        "platform": system or platform.platform(),
        "python": platform.python_version(),
        "dependencies": dependencies,
        "tools": tool_requirements,
        "missing_required": [name for name in ("MCP SDK", "python-docx", "pypdfium2")
                             if not status.get(name, False)],
    }
