"""Bundled OFL fonts (see assets/fonts/*/OFL.txt and scripts/fetch_fonts.py)."""

from __future__ import annotations

import json
import os
import tempfile
import threading
from functools import lru_cache
from pathlib import Path

FONT_DIR = Path(__file__).parent / "assets" / "fonts"
_LOCK = threading.Lock()


@lru_cache(maxsize=1)
def manifest() -> tuple[dict, ...]:
    return tuple(json.loads((FONT_DIR / "fonts.json").read_text(encoding="utf-8")))


def font_face_css(families: set[str]) -> str:
    """@font-face rules for the requested bundled families."""
    rules = []
    for f in manifest():
        if f["family"] in families:
            uri = (FONT_DIR / f["file"]).as_uri()
            rules.append(
                f"@font-face {{ font-family: '{f['family']}'; font-weight: {f['weight']}; "
                f"font-style: {f['style']}; src: url('{uri}') format('woff2'); }}")
    return "\n".join(rules)


def ttf_files(families: set[str]) -> list[Path]:
    """TrueType copies of the bundled fonts (matplotlib cannot read WOFF2)."""
    from fontTools.ttLib import TTFont

    from . import __version__

    cache = Path(tempfile.gettempdir()) / f"doc-mantap-mcp-fonts-{__version__}"
    out = []
    with _LOCK:
        cache.mkdir(parents=True, exist_ok=True)
        for f in manifest():
            if f["family"] not in families or f["style"] != "normal":
                continue
            target = cache / (f["file"].replace("/", "_").replace(".woff2", ".ttf"))
            if not target.exists():
                font = TTFont(FONT_DIR / f["file"])
                font.flavor = None
                tmp = target.with_suffix(f".{os.getpid()}.tmp")
                font.save(tmp)
                os.replace(tmp, target)
            out.append(target)
    return out
