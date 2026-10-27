"""Shared helpers."""

from pathlib import Path

import markdown

MD_EXTENSIONS = ["tables", "fenced_code", "sane_lists", "attr_list", "toc"]


def md_to_html(content: str) -> str:
    return markdown.markdown(content, extensions=MD_EXTENSIONS)


def prepare_output(output_path: str, suffix: str) -> Path:
    """Resolve the path, force the extension and create parent folders."""
    path = Path(output_path).expanduser().resolve()
    if path.suffix.lower() != suffix:
        path = path.with_suffix(suffix)
    path.parent.mkdir(parents=True, exist_ok=True)
    return path
