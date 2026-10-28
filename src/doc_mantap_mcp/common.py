"""Shared helpers: output paths and the Markdown -> semantic HTML pipeline.

Both the PDF and the DOCX renderer consume the BeautifulSoup tree produced by
``parse_markdown``. Special blocks (charts, KPI cards) are replaced by placeholder
``div`` elements whose parsed specs live on the returned ``ParsedDoc``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

import markdown
import yaml
from bs4 import BeautifulSoup, Comment, NavigableString, Tag
from markdown.extensions.toc import slugify_unicode

from .charts import parse_chart_spec

MD_EXTENSIONS = [
    "tables", "sane_lists", "attr_list", "def_list", "abbr", "footnotes", "md_in_html",
    "toc", "smarty", "pymdownx.superfences", "pymdownx.highlight", "pymdownx.tasklist",
    "pymdownx.tilde", "pymdownx.caret", "pymdownx.mark",
]
MD_CONFIG = {
    "pymdownx.highlight": {"use_pygments": False},
    "toc": {"slugify": slugify_unicode},
}

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg"}
PAGE_BREAK_HTML = '<div class="page-break"></div>'

LABELS = {
    "en": {"note": "Note", "info": "Info", "tip": "Tip", "success": "Success",
           "important": "Important", "warning": "Warning", "caution": "Caution",
           "danger": "Danger", "contents": "Contents", "notes": "Notes",
           "prepared_by": "Prepared by", "date": "Date", "organization": "Organization"},
    "id": {"note": "Catatan", "info": "Info", "tip": "Tips", "success": "Berhasil",
           "important": "Penting", "warning": "Peringatan", "caution": "Perhatian",
           "danger": "Bahaya", "contents": "Daftar Isi", "notes": "Catatan",
           "prepared_by": "Disusun oleh", "date": "Tanggal", "organization": "Organisasi"},
}

# type -> (color, icon glyph)
CALLOUTS = {
    "note": ("2563EB", "i"), "info": ("0284C7", "i"), "tip": ("059669", "★"),
    "success": ("16A34A", "✓"), "important": ("7C3AED", "!"), "warning": ("D97706", "!"),
    "caution": ("DC2626", "!"), "danger": ("DC2626", "×"),
}
_CALLOUT_ALIASES = {
    "hint": "tip", "check": "success", "done": "success", "error": "danger", "bug": "danger",
    "attention": "warning", "abstract": "note", "summary": "note", "question": "info",
    "example": "info", "quote": "note",
}

GOOD, BAD, FLAT = "good", "bad", "flat"


def labels(lang: str) -> dict[str, str]:
    return LABELS.get((lang or "en").split("-")[0].lower(), LABELS["en"])


def prepare_output(output_path: str, suffix: str) -> Path:
    """Resolve the path, force the extension and create parent folders."""
    path = Path(output_path).expanduser().resolve()
    if path.suffix.lower() != suffix:
        path = path.with_suffix(suffix)
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def load_block(text: str, kind: str, strings: bool = False):
    """Parse a fenced block body as YAML (JSON is valid YAML too)."""
    try:
        return yaml.load(text, Loader=yaml.BaseLoader if strings else yaml.SafeLoader)
    except yaml.YAMLError as exc:
        raise ValueError(f"Invalid ```{kind} block (expected YAML or JSON): {exc}") from None


# --------------------------------------------------------------------------- #
# Markdown normalisation (AI-written GFM -> what Python-Markdown expects)
# --------------------------------------------------------------------------- #

_FENCE_RE = re.compile(r"^(?P<indent>[ ]*)(?P<fence>`{3,}|~{3,})")
_LIST_RE = re.compile(r"^(?P<indent>[ ]*)(?:[-*+]|\d{1,9}[.)])[ ]+\S")
_TABLE_ROW_RE = re.compile(r"^[ ]{0,3}\|")
_TABLE_SEP_RE = re.compile(r"^[ ]{0,3}\|?[ ]*:?-{2,}:?[ ]*(\|[ ]*:?-{2,}:?[ ]*)*\|?[ ]*$")
_HEADING_RE = re.compile(r"^[ ]{0,3}#{1,6}[ ]")
_PAGEBREAK_RE = re.compile(
    r"^[ ]*(?:\\pagebreak|\\newpage|<!--\s*page-?break\s*-->)[ ]*$", re.I)


def _indent(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def _is_fence_close(line: str, fence: str) -> bool:
    m = _FENCE_RE.match(line)
    return bool(m and m.group("fence")[0] == fence[0]
                and len(m.group("fence")) >= len(fence)
                and line.strip() == m.group("fence"))


def _fix_list_indent(lines: list[str]) -> list[str]:
    """Rescale 2/3-space nested lists to the 4 spaces Python-Markdown needs."""
    out = list(lines)
    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        m = _FENCE_RE.match(line)
        if m and _indent(line) == 0:  # skip top-level fenced code
            fence, i = m.group("fence"), i + 1
            while i < n and not _is_fence_close(lines[i], fence):
                i += 1
            i += 1
            continue
        if not (_LIST_RE.match(line) and _indent(line) == 0):
            i += 1
            continue
        # Find the extent of this list block.
        start, j, prev_blank, fence = i, i + 1, False, None
        while j < n:
            cur = lines[j]
            if fence:
                if _is_fence_close(cur, fence):
                    fence = None
            elif not cur.strip():
                prev_blank = True
            elif _indent(cur) == 0 and not _LIST_RE.match(cur) and (
                    prev_blank or _FENCE_RE.match(cur) or _HEADING_RE.match(cur)
                    or _TABLE_ROW_RE.match(cur) or cur.startswith(">")):
                break
            else:
                fm = _FENCE_RE.match(cur)
                if fm and _indent(cur) > 0:
                    fence = fm.group("fence")
                prev_blank = False
            j += 1
        block = range(start, j)
        # Smallest indent used by a nested list item (outside code fences).
        unit, fence = None, None
        for k in block:
            cur = lines[k]
            if fence:
                if _is_fence_close(cur, fence):
                    fence = None
                continue
            fm = _FENCE_RE.match(cur)
            if fm:
                fence = fm.group("fence")
                continue
            if _LIST_RE.match(cur) and _indent(cur) > 0:
                unit = min(unit or 99, _indent(cur))
        if unit and unit < 4:
            fence, delta = None, 0
            for k in block:
                cur = lines[k]
                ind = _indent(cur)
                if fence:
                    out[k] = " " * (ind + delta) + cur.lstrip(" ") if cur.strip() else cur
                    if _is_fence_close(cur, fence):
                        fence = None
                    continue
                if ind and cur.strip():
                    new = (ind // unit) * 4 + ind % unit
                    out[k] = " " * new + cur.lstrip(" ")
                    fm = _FENCE_RE.match(cur)
                    if fm:
                        fence, delta = fm.group("fence"), new - ind
        i = j
    return out


def normalize_markdown(text: str) -> str:
    """Make GitHub-flavoured Markdown written by AI render as intended.

    - lists, tables and code fences directly after a paragraph get a blank line
    - 2-space nested lists become 4-space
    - ``\\pagebreak``, ``\\newpage`` and ``<!-- pagebreak -->`` become page breaks
    """
    lines = _fix_list_indent(
        text.replace("\r\n", "\n").replace("\r", "\n").expandtabs(4).split("\n"))
    out: list[str] = []
    fence: str | None = None
    block: str | None = None  # None (after a blank line), para, list, table, quote, other
    last_block: str | None = None  # type of the previous non-blank block
    for i, line in enumerate(lines):
        if fence:
            out.append(line)
            if _is_fence_close(line, fence):
                fence = None
            continue
        if not line.strip():
            out.append(line)
            if block is not None:
                last_block = block
            block = None
            continue
        if _PAGEBREAK_RE.match(line):
            out.extend(["", PAGE_BREAK_HTML, ""])
            block = last_block = None
            continue
        indent = _indent(line)
        if block is None and last_block == "quote" and line.lstrip().startswith(">") \
                and indent < 4:
            # GFM keeps blank-line separated quotes apart; Python-Markdown merges them.
            out.extend(["<!-- -->", ""])
        fm = _FENCE_RE.match(line)
        if fm:
            if block == "para" or (block and indent == 0):
                out.append("")
            fence = fm.group("fence")
            out.append(line)
            block = "list" if block == "list" and indent else None
            continue
        if _LIST_RE.match(line):
            if block in ("para", "table", "quote", "other") and indent == 0:
                out.append("")
            out.append(line)
            block = "list"
            continue
        if (_TABLE_ROW_RE.match(line) and i + 1 < len(lines)
                and _TABLE_SEP_RE.match(lines[i + 1]) and block not in (None, "table")):
            if indent == 0:
                out.append("")
            out.append(line)
            block = "table"
            continue
        out.append(line)
        if block is None:
            if _HEADING_RE.match(line):
                block = None
            elif _TABLE_ROW_RE.match(line):
                block = "table"
            elif line.lstrip().startswith(">"):
                block = "quote"
            elif line.lstrip().startswith("<"):
                block = "other"
            else:
                block = "para"
    return "\n".join(out)


# --------------------------------------------------------------------------- #
# Parsed document
# --------------------------------------------------------------------------- #

@dataclass
class Heading:
    level: int
    text: str
    id: str


@dataclass
class ParsedDoc:
    soup: BeautifulSoup
    lang: str = "en"
    headings: list[Heading] = field(default_factory=list)
    charts: list[dict] = field(default_factory=list)
    kpis: list[list[dict]] = field(default_factory=list)
    footnotes: dict[str, Tag] = field(default_factory=dict)

    @property
    def labels(self) -> dict[str, str]:
        return labels(self.lang)


def parse_kpis(text: str) -> list[dict]:
    data = load_block(text, "kpi", strings=True)
    if isinstance(data, dict):
        data = data.get("items") or data.get("kpis") or [
            {"label": k, "value": v} for k, v in data.items()]
    if not isinstance(data, list) or not data:
        raise ValueError("A ```kpi block must be a YAML list of items with "
                         "'label' and 'value' (optional: delta, note, lower_is_better).")
    items = []
    for n, raw in enumerate(data, 1):
        if not isinstance(raw, dict):
            raise ValueError(f"kpi item {n} must be a mapping like "
                             "{label: Revenue, value: '$1.2M', delta: '+12%'}.")
        label = raw.get("label") or raw.get("title") or raw.get("name")
        value = raw.get("value")
        if label is None or value is None:
            raise ValueError(f"kpi item {n} needs both 'label' and 'value'.")
        delta = raw.get("delta") or raw.get("change")
        trend = (raw.get("trend") or "").lower()
        if not trend and delta:
            first = str(delta).strip()[:1]
            trend = ("up" if first in "+▲↑" else "down" if first in "-−▼↓" else "")
        lower_better = str(raw.get("lower_is_better", "")).lower() in ("true", "yes", "1")
        tone = FLAT
        if trend in ("up", "down"):
            tone = GOOD if (trend == "up") != lower_better else BAD
        items.append({
            "label": str(label), "value": str(value),
            "delta": str(delta) if delta else None,
            "note": str(raw["note"]) if raw.get("note") else None,
            "trend": trend, "tone": tone,
        })
    return items


def _special_blocks(doc: ParsedDoc) -> None:
    for pre in doc.soup.find_all("pre"):
        code = pre.find("code")
        if code is None:
            continue
        lang = next((c[9:] for c in code.get("class", []) if c.startswith("language-")), "")
        lang = lang.lower()
        text = code.get_text()
        if lang == "chart":
            doc.charts.append(parse_chart_spec(text))
            pre.replace_with(doc.soup.new_tag(
                "div", attrs={"class": "dm-chart", "data-index": str(len(doc.charts) - 1)}))
        elif lang in ("kpi", "kpis", "stats", "metrics"):
            doc.kpis.append(parse_kpis(text))
            pre.replace_with(doc.soup.new_tag(
                "div", attrs={"class": "dm-kpis", "data-index": str(len(doc.kpis) - 1)}))
        else:
            pre["data-lang"] = lang


_ALERT_RE = re.compile(r"^\s*\[!([A-Za-z]+)\][ \t]*([^\n]*)\n?")


def _callouts(doc: ParsedDoc) -> None:
    soup, names = doc.soup, doc.labels
    for bq in soup.find_all("blockquote"):
        first = next((c for c in bq.children if isinstance(c, Tag)), None)
        if first is None or first.name != "p" or not first.contents:
            continue
        head = first.contents[0]
        if not isinstance(head, NavigableString):
            continue
        m = _ALERT_RE.match(str(head))
        if not m:
            continue
        kind = m.group(1).lower()
        kind = _CALLOUT_ALIASES.get(kind, kind)
        if kind not in CALLOUTS:
            kind = "note"
        rest = str(head)[m.end():]
        if rest.strip():
            head.replace_with(NavigableString(rest))
        else:
            head.extract()
            while first.contents and isinstance(first.contents[0], Tag) \
                    and first.contents[0].name == "br":
                first.contents[0].extract()
        if not first.get_text(strip=True) and not first.find("img"):
            first.decompose()
        box = soup.new_tag("div", attrs={"class": f"callout callout-{kind}",
                                         "data-kind": kind})
        title = soup.new_tag("p", attrs={"class": "callout-title"})
        title.string = m.group(2).strip() or names.get(kind, kind.title())
        box.append(title)
        for child in list(bq.contents):
            box.append(child.extract())
        bq.replace_with(box)


def _task_lists(soup: BeautifulSoup) -> None:
    for box in soup.select("li.task-list-item input"):
        done = box.has_attr("checked")
        box.replace_with(soup.new_tag("span", attrs={
            "class": f"task-box {'done' if done else 'todo'}"}))


def _footnotes(doc: ParsedDoc) -> None:
    holder = doc.soup.find("div", class_="footnote")
    if holder is None:
        return
    for li in holder.find_all("li", id=True):
        for back in li.find_all("a", class_="footnote-backref"):
            back.decompose()
        doc.footnotes[li["id"]] = li
    holder.extract()


def _images(soup: BeautifulSoup, base_dir: Path) -> None:
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        scheme = urlparse(src).scheme.lower()
        if scheme in ("http", "https", "data"):
            continue
        if scheme == "file":
            path = Path(url2pathname(urlparse(src).path))
        else:
            path = Path(unquote(src)).expanduser()
            if not path.is_absolute():
                path = base_dir / path
        path = path.resolve()
        img["data-path"] = str(path)
        img["src"] = path.as_uri()
    for img in soup.find_all("img"):
        parent = img.parent
        if parent is None or parent.name != "p":
            continue
        others = [c for c in parent.contents if c is not img
                  and not (isinstance(c, NavigableString) and not c.strip())]
        if others:
            continue
        figure = soup.new_tag("figure", attrs={"class": "image"})
        figure.append(img.extract())
        if img.get("title"):
            cap = soup.new_tag("figcaption")
            cap.string = img["title"]
            figure.append(cap)
        parent.replace_with(figure)


_NUM_RE = re.compile(
    r"^[\s(+\-−~≈]*(?:rp\.?|idr|usd|eur|us\$|\$|€|£|¥)?\s*\d[\d.,\s]*"
    r"(?:%|x|k|m|b|t|bn|jt|rb|pp|pts?)?\)?\s*$", re.I)


def _tables(soup: BeautifulSoup) -> None:
    """Right-align numeric columns that have no explicit alignment."""
    for table in soup.find_all("table"):
        rows = [r.find_all(["th", "td"]) for r in table.find_all("tr")]
        if not rows:
            continue
        for j in range(max(len(r) for r in rows)):
            cells = [r[j] for r in rows if len(r) > j]
            if any("text-align" in (c.get("style") or "") for c in cells):
                continue
            values = [t for c in cells if c.name == "td"
                      and (t := c.get_text(strip=True)) not in ("", "-", "–", "—", "n/a")]
            if values and sum(bool(_NUM_RE.match(v)) for v in values) >= 0.8 * len(values):
                for c in cells:
                    c["style"] = "text-align: right;"


def parse_markdown(content: str, lang: str = "en", base_dir: Path | None = None) -> ParsedDoc:
    md = markdown.Markdown(extensions=MD_EXTENSIONS, extension_configs=MD_CONFIG)
    soup = BeautifulSoup(md.convert(normalize_markdown(content)), "html.parser")
    for comment in soup.find_all(string=lambda s: isinstance(s, Comment)):
        comment.extract()
    doc = ParsedDoc(soup=soup, lang=lang or "en")
    _special_blocks(doc)
    _callouts(doc)
    _task_lists(soup)
    _footnotes(doc)
    _images(soup, base_dir or Path.cwd())
    _tables(soup)
    doc.headings = [Heading(int(h.name[1]), h.get_text(" ", strip=True), h.get("id", ""))
                    for h in soup.find_all(["h1", "h2", "h3"])]
    return doc
