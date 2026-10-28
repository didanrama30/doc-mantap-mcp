"""doc-mantap-mcp server."""

import functools
from typing import Literal

from mcp.server.mcpserver import Image, MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import ToolAnnotations

from .docx_render import render_docx
from .docx_inspect import outline_docx, read_docx
from .docx_patch import patch_docx
from .guide import guide
from .pdf import render_pdf
from .pdf_read import read_pdf_text
from .preview import render_pages
from .themes import THEMES, get_theme
from .xlsx_render import render_xlsx
from .xlsx_spec import Sheet

mcp = MCPServer(
    "doc-mantap-mcp",
    instructions=(
        "Turns AI-written content into beautiful files.\n"
        "- Documents: write Markdown, then call markdown_to_pdf (most polished) or "
        "markdown_to_docx (editable Word). Pass `title` (plus subtitle, document_type, author, "
        "date, organization, logo) for a designed cover; don't repeat the title as a heading. "
        "Use toc=true for long documents and lang='id' for Indonesian text.\n"
        "- Rich blocks: callouts (`> [!TIP]`), ```kpi cards, ```chart blocks "
        "(bar/line/area/pie/donut/scatter), footnotes, task lists, images, page breaks "
        "(\\pagebreak).\n"
        "- Spreadsheets: create_excel builds designed workbooks with real Excel tables, "
        "formulas, KPI cards, conditional formatting and native charts.\n"
        "- Existing files: docx_read/docx_outline inspect Word templates, docx_patch makes "
        "anchored edits, and pdf_read_text extracts PDF page text. markdown_to_docx supports "
        "mode='append' with target_path to add Markdown to an existing Word file.\n"
        "Call design_guide once for the full syntax and design rules, list_themes for themes, "
        "and preview_document to look at the result before handing it over."
    ),
)

_WRITE = {"read_only_hint": False, "destructive_hint": True, "idempotent_hint": True}


def _explain(func):
    """Send anticipated errors (bad spec, missing library, file problems) to the model.

    The SDK hides the text of any exception that is not a ToolError, which would leave the
    model unable to fix its input.
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ToolError:
            raise
        except (ValueError, RuntimeError, OSError) as exc:
            raise ToolError(str(exc)) from exc
    return wrapper


@mcp.tool(annotations=ToolAnnotations(title="List themes", read_only_hint=True))
def list_themes() -> list[dict]:
    """List the available themes with their colors, fonts and default cover style."""
    return [t.summary() for t in THEMES.values()]


@mcp.tool(annotations=ToolAnnotations(title="Design guide", read_only_hint=True))
def design_guide(topic: Literal["documents", "excel", "all"] = "all") -> str:
    """Syntax reference (callouts, KPI cards, charts, Excel spec) and design rules.

    Read it before writing a long document or a workbook.
    """
    return guide(topic)


@mcp.tool(annotations=ToolAnnotations(title="Read Word document", read_only_hint=True))
@_explain
def docx_read(path: str, include_runs: bool = True,
              include_headers_footers: bool = True) -> dict:
    """Read DOCX text, paragraph styles, run formatting, tables and document furniture as JSON.

    This reads OOXML directly through python-docx; LibreOffice is not required.
    """
    return read_docx(path, include_runs, include_headers_footers)


@mcp.tool(annotations=ToolAnnotations(title="Word document outline", read_only_hint=True))
@_explain
def docx_outline(path: str) -> dict:
    """Return compact heading hierarchy, table dimensions and detected placeholders."""
    return outline_docx(path)


@mcp.tool(annotations=ToolAnnotations(title="Patch Word document", read_only_hint=False,
                                     destructive_hint=True, idempotent_hint=False))
@_explain
def docx_patch(path: str, operations: list[dict], output_path: str | None = None,
               backup: bool = True, overwrite: bool = False) -> dict:
    """Patch existing DOCX/DOCM in OOXML without regenerating its structure.

    Operations: replace_text(anchor, new_text), insert_after(anchor, paragraphs),
    fill_table_cell(table, row, col, content), append_to_cell(...), or
    insert_code_block(anchor, text). Default output is a new .patched file. To edit
    in place, set output_path to path; an unused .bak is required automatically.
    """
    return patch_docx(path, operations, output_path, backup, overwrite)


@mcp.tool(annotations=ToolAnnotations(title="Read PDF text", read_only_hint=True))
@_explain
def pdf_read_text(path: str, ocr: bool = True, language: str = "eng+ind") -> dict:
    """Extract native PDF text per page; OCR blank/scanned pages with Tesseract if installed."""
    return read_pdf_text(path, ocr, language)


@mcp.tool(annotations=ToolAnnotations(title="Markdown to PDF", **_WRITE))
@_explain
def markdown_to_pdf(content: str, output_path: str, theme: str = "modern",
                    title: str | None = None, subtitle: str | None = None,
                    author: str | None = None, date: str | None = None,
                    organization: str | None = None, document_type: str | None = None,
                    logo: str | None = None, cover_style: str | None = None,
                    toc: bool = False, page_size: str = "A4", orientation: str = "portrait",
                    lang: str = "en", primary_color: str | None = None,
                    accent_color: str | None = None, base_dir: str | None = None) -> str:
    """Convert Markdown into a designed PDF (cover, running header, page numbers).

    Supports callouts (> [!TIP]), ```kpi and ```chart blocks, footnotes, task lists and
    syntax-highlighted code. See design_guide for the syntax.

    Args:
        content: Document body in Markdown.
        output_path: Output file path, e.g. ~/Documents/report.pdf.
        theme: modern, corporate, academic, fresh, elegant, minimal, startup or warm.
        title: Adds a cover page with this title.
        subtitle: Cover subtitle.
        author: Cover author ("Prepared by").
        date: Cover date, written as you want it shown.
        organization: Organization name on the cover and in the footer.
        document_type: Small label above the title, e.g. "Quarterly Report".
        logo: Path to a logo image for the cover.
        cover_style: band, full, classic or minimal (default depends on the theme).
        toc: Add a table of contents after the cover.
        page_size: A4, Letter, Legal or A5.
        orientation: portrait or landscape.
        lang: Language code of the text (e.g. "en", "id"): hyphenation, labels, numbers.
        primary_color: Brand color (hex) replacing the theme's primary color.
        accent_color: Brand accent color (hex).
        base_dir: Folder that relative image paths are resolved against (default: output folder).
    """
    result = render_pdf(content, output_path, get_theme(theme, primary_color, accent_color),
                        title, subtitle, author, date, organization=organization,
                        document_type=document_type, logo=logo, cover_style=cover_style,
                        toc=toc, page_size=page_size, orientation=orientation, lang=lang,
                        base_dir=base_dir)
    return (f"PDF created: {result.path} ({result.pages} pages). "
            "Use preview_document to check the layout.")


@mcp.tool(annotations=ToolAnnotations(title="Markdown to Word", read_only_hint=False,
                                     destructive_hint=True, idempotent_hint=False))
@_explain
def markdown_to_docx(content: str, output_path: str | None = None, theme: str = "modern",
                     title: str | None = None, subtitle: str | None = None,
                     author: str | None = None, date: str | None = None,
                     organization: str | None = None, document_type: str | None = None,
                     logo: str | None = None, cover_style: str | None = None,
                     toc: bool = False, page_size: str = "A4", orientation: str = "portrait",
                     lang: str = "en", primary_color: str | None = None,
                     accent_color: str | None = None, base_dir: str | None = None,
                     mode: Literal["create", "append"] = "create",
                     target_path: str | None = None, start_after: str | None = None) -> str:
    """Convert Markdown into a styled, editable Word document.

    Uses real Word styles (headings, lists, captions, TOC) so the file stays easy to edit.
    Same Markdown features and arguments as markdown_to_pdf.

    Args:
        content: Document body in Markdown.
        output_path: Output file path, e.g. ~/Documents/report.docx. In append mode, defaults
            to a sibling .appended.docx file; pass target_path to update target itself with backup.
        theme: modern, corporate, academic, fresh, elegant, minimal, startup or warm.
        title: Adds a cover page with this title.
        subtitle: Cover subtitle.
        author: Cover author ("Prepared by").
        date: Cover date, written as you want it shown.
        organization: Organization name on the cover and in the footer.
        document_type: Small label above the title, e.g. "Proposal".
        logo: Path to a logo image (PNG/JPG) for the cover.
        cover_style: band, full, classic or minimal (default depends on the theme).
        toc: Add a table of contents (Word fills in page numbers when the file opens).
        page_size: A4, Letter, Legal or A5.
        orientation: portrait or landscape.
        lang: Language code of the text (e.g. "en", "id"): proofing language and labels.
        primary_color: Brand color (hex) replacing the theme's primary color.
        accent_color: Brand accent color (hex).
        base_dir: Folder that relative image paths are resolved against (default: output folder).
        mode: "create" for a designed document, or "append" to add content to target_path.
        target_path: Existing .docx to preserve and append into when mode="append".
        start_after: Unique paragraph text or p:index insertion point in append mode.
    """
    path = render_docx(content, output_path, get_theme(theme, primary_color, accent_color),
                       title, subtitle, author, date, organization=organization,
                       document_type=document_type, logo=logo, cover_style=cover_style,
                       toc=toc, page_size=page_size, orientation=orientation, lang=lang,
                       base_dir=base_dir, mode=mode, target_path=target_path,
                       start_after=start_after)
    note = " Word will offer to update fields on open; accept to fill in the TOC." if toc else ""
    action = "DOCX updated" if mode == "append" else "DOCX created"
    return f"{action}: {path}.{note}"


@mcp.tool(annotations=ToolAnnotations(title="Create Excel workbook", **_WRITE))
@_explain
def create_excel(output_path: str, sheets: list[Sheet], theme: str = "modern",
                 title: str | None = None, author: str | None = None,
                 company: str | None = None, currency: str | None = None, lang: str = "en",
                 primary_color: str | None = None, accent_color: str | None = None) -> str:
    """Create a designed Excel workbook from a declarative description.

    Each sheet can have a banner title, KPI cards, an assumptions block (named input cells),
    Excel tables with formulas, totals, dropdowns, conditional formatting and sparklines,
    native charts and notes. Fonts and colors follow the theme. Formulas are checked
    before writing. Call design_guide(topic="excel") for a full example.

    Args:
        output_path: Output file path, e.g. ~/Documents/sales.xlsx.
        sheets: Sheet definitions (see the schema and design_guide).
        theme: modern, corporate, academic, fresh, elegant, minimal, startup or warm.
        title: Workbook title (document properties).
        author: Workbook author.
        company: Company name (document properties).
        currency: Currency symbol for currency columns, e.g. "Rp", "$", "€".
            Default: "Rp" when lang is "id", otherwise "$".
        lang: Language code, e.g. "en" or "id" (labels and default currency).
        primary_color: Brand color (hex) replacing the theme's primary color.
        accent_color: Brand accent color (hex).
    """
    result = render_xlsx(output_path, sheets, get_theme(theme, primary_color, accent_color),
                         title=title, author=author, company=company, currency=currency,
                         lang=lang)
    lines = [f"Excel workbook created: {result.path}", *result.summary]
    if result.warnings:
        lines += ["Warnings:", *(f"  - {w}" for w in result.warnings)]
    lines.append("Formulas recalculate when the file is opened in Excel.")
    return "\n".join(lines)


@mcp.tool(annotations=ToolAnnotations(title="Preview document", read_only_hint=True))
@_explain
def preview_document(path: str, pages: str = "1-2", width: int = 1000) -> list:
    """Render pages of a PDF (or DOCX/XLSX when LibreOffice is installed) as images.

    Use it after creating a document to check covers, page breaks, tables and charts.

    Args:
        path: File to preview.
        pages: Pages to render, e.g. "1", "1-3", "2,5" or "all" (max 6 per call).
        width: Image width in pixels (300-1800).
    """
    images, count, note = render_pages(path, pages, width)
    text = f"{path}: showing {len(images)} of {count} page(s). {note}".strip()
    return [text, *(Image(data=png, format="png") for png in images)]


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
