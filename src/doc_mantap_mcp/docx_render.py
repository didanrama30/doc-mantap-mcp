"""Render Markdown -> DOCX with python-docx."""

from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .common import md_to_html, prepare_output
from .themes import Theme

MONO = "Consolas"


def _rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color.upper())


def _set_font(target, name: str) -> None:
    """Set the font for all scripts so Word renders it consistently."""
    target.font.name = name
    rpr = target.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)


def _shade(element, fill: str) -> None:
    """Set background color on a paragraph or table cell."""
    props = element.get_or_add_pPr() if hasattr(element, "get_or_add_pPr") \
        else element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    props.append(shd)


def _border(paragraph, side: str, color: str, size: int) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    edge = OxmlElement(f"w:{side}")
    for k, v in {"w:val": "single", "w:sz": str(size), "w:space": "6",
                 "w:color": color}.items():
        edge.set(qn(k), v)
    bdr.append(edge)
    ppr.append(bdr)


def _add_page_number(section, theme: Theme) -> None:
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])
    run.font.size = Pt(9)
    run.font.color.rgb = _rgb(theme.muted)


def _setup_styles(doc: Document, theme: Theme) -> None:
    for section in doc.sections:
        section.page_height, section.page_width = Cm(29.7), Cm(21.0)
        section.top_margin = section.bottom_margin = Cm(2.5)
        section.left_margin = section.right_margin = Cm(2.2)

    normal = doc.styles["Normal"]
    _set_font(normal, theme.docx_body)
    normal.font.size = Pt(theme.body_size_pt)
    normal.font.color.rgb = _rgb(theme.text)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.3

    sizes = {1: 24, 2: 17, 3: 13.5, 4: 12, 5: 11, 6: 11}
    for level, size in sizes.items():
        style = doc.styles[f"Heading {level}"]
        _set_font(style, theme.docx_heading)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.italic = False
        style.font.color.rgb = _rgb(theme.primary if level <= 3 else theme.text)
        style.paragraph_format.space_before = Pt(18 if level <= 2 else 12)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.keep_with_next = True


def _add_inline(paragraph, node, theme: Theme, bold=False, italic=False,
                code=False, link=False) -> None:
    """Add inline content (text, bold, italic, code, links) recursively."""
    for child in node.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if not text:
                continue
            run = paragraph.add_run(text)
            run.bold = bold or None
            run.italic = italic or None
            if bold:
                run.font.color.rgb = _rgb(theme.primary)
            if code:
                _set_font(run, MONO)
                run.font.size = Pt(theme.body_size_pt - 1.5)
                run.font.color.rgb = _rgb(theme.accent)
            if link:
                run.font.color.rgb = _rgb(theme.accent)
                run.underline = True
        elif isinstance(child, Tag):
            if child.name == "br":
                paragraph.add_run().add_break()
                continue
            _add_inline(
                paragraph, child, theme,
                bold=bold or child.name in ("strong", "b"),
                italic=italic or child.name in ("em", "i"),
                code=code or child.name == "code",
                link=link or child.name == "a",
            )


def _add_list(doc, list_tag: Tag, theme: Theme, ordered: bool, depth: int = 1) -> None:
    base = "List Number" if ordered else "List Bullet"
    style = base if depth == 1 else f"{base} {min(depth, 3)}"
    for li in list_tag.find_all("li", recursive=False):
        p = doc.add_paragraph(style=style)
        p.paragraph_format.space_after = Pt(3)
        nested = []
        for child in li.children:
            if isinstance(child, Tag) and child.name in ("ul", "ol"):
                nested.append(child)
            elif isinstance(child, Tag):
                _add_inline(p, child if child.name == "p" else _wrap(child), theme)
            elif str(child).strip():
                _add_inline(p, _wrap(str(child).rstrip("\n")), theme)
        for sub in nested:
            _add_list(doc, sub, theme, sub.name == "ol", depth + 1)


def _wrap(node) -> Tag:
    """Wrap a single node in a span so it can go through _add_inline."""
    span = BeautifulSoup("<span></span>", "html.parser").span
    span.append(node.__copy__() if isinstance(node, Tag) else NavigableString(str(node)))
    return span


def _add_table(doc, table_tag: Tag, theme: Theme) -> None:
    rows = table_tag.find_all("tr")
    if not rows:
        return
    ncols = max(len(r.find_all(["th", "td"])) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = doc.styles["Table Grid"]
    for i, row in enumerate(rows):
        cells = row.find_all(["th", "td"])
        is_header = all(c.name == "th" for c in cells)
        for j, cell_tag in enumerate(cells):
            cell = table.cell(i, j)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            _add_inline(p, cell_tag, theme, bold=is_header)
            if is_header:
                _shade(cell._tc, theme.primary)
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            elif i % 2 == 0:
                _shade(cell._tc, theme.surface)
    doc.add_paragraph()


def _add_code_block(doc, pre: Tag) -> None:
    lines = pre.get_text().rstrip("\n").split("\n")
    p = doc.add_paragraph()
    _shade(p._p, "0F172A")
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.line_spacing = 1.1
    for idx, line in enumerate(lines):
        run = p.add_run(line)
        _set_font(run, MONO)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
        if idx < len(lines) - 1:
            run.add_break()


def _add_cover(doc, theme: Theme, title: str, subtitle: str | None,
               author: str | None, date: str | None) -> None:
    for _ in range(8):
        doc.add_paragraph()
    bar = doc.add_paragraph().add_run("━━━━━━")
    bar.font.color.rgb = _rgb(theme.accent)
    bar.font.size = Pt(20)

    r = doc.add_paragraph().add_run(title)
    _set_font(r, theme.docx_heading)
    r.font.size, r.bold = Pt(32), True
    r.font.color.rgb = _rgb(theme.primary)

    if subtitle:
        r = doc.add_paragraph().add_run(subtitle)
        r.font.size = Pt(15)
        r.font.color.rgb = _rgb(theme.muted)

    meta = " · ".join(x for x in (author, date) if x)
    if meta:
        for _ in range(4):
            doc.add_paragraph()
        m = doc.add_paragraph()
        _border(m, "top", theme.accent, 6)
        r = m.add_run(meta)
        r.font.size = Pt(11)
        r.font.color.rgb = _rgb(theme.muted)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def render_docx(content: str, output_path: str, theme: Theme, title: str | None = None,
                subtitle: str | None = None, author: str | None = None,
                date: str | None = None) -> str:
    out = prepare_output(output_path, ".docx")
    doc = Document()
    _setup_styles(doc, theme)
    _add_page_number(doc.sections[0], theme)
    if title:
        doc.core_properties.title = title
        _add_cover(doc, theme, title, subtitle, author, date)
    if author:
        doc.core_properties.author = author

    soup = BeautifulSoup(md_to_html(content), "html.parser")
    for el in soup.children:
        if not isinstance(el, Tag):
            continue
        name = el.name
        if name in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(name[1])
            h = doc.add_heading(level=level)
            _add_inline(h, el, theme)
            if level == 1:
                _border(h, "bottom", theme.accent, 18)
            elif level == 2:
                _border(h, "left", theme.accent, 24)
        elif name == "p":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _add_inline(p, el, theme)
        elif name in ("ul", "ol"):
            _add_list(doc, el, theme, ordered=name == "ol")
        elif name == "table":
            _add_table(doc, el, theme)
        elif name == "pre":
            _add_code_block(doc, el)
        elif name == "blockquote":
            for inner in el.find_all("p") or [el]:
                q = doc.add_paragraph()
                q.paragraph_format.left_indent = Cm(0.4)
                _shade(q._p, theme.surface)
                _border(q, "left", theme.accent, 24)
                _add_inline(q, inner, theme, italic=True)
                for run in q.runs:
                    run.font.color.rgb = _rgb(theme.muted)
        elif name == "hr":
            _border(doc.add_paragraph(), "bottom", theme.surface, 6)
        else:
            _add_inline(doc.add_paragraph(), el, theme)

    doc.save(str(out))
    return str(out)
