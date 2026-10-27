"""doc-mantap-mcp server."""

from mcp.server.mcpserver import MCPServer

from .docx_render import render_docx
from .pdf import render_pdf
from .themes import THEMES, get_theme

mcp = MCPServer(
    "doc-mantap-mcp",
    instructions=(
        "Use these tools to turn AI-written content into beautiful PDF/DOCX documents. "
        "Write content in Markdown (headings, lists, tables, quotes, code blocks). "
        "If you pass `title`, a cover page is created, so don't repeat it as a heading. "
        "Indent nested list items with 4 spaces. "
        "Call list_themes to see available themes."
    ),
)


@mcp.tool()
def list_themes() -> list[dict]:
    """List available themes with their description and main colors."""
    return [
        {"name": t.name, "description": t.description,
         "primary": f"#{t.primary}", "accent": f"#{t.accent}"}
        for t in THEMES.values()
    ]


@mcp.tool()
def markdown_to_pdf(content: str, output_path: str, theme: str = "modern",
                    title: str | None = None, subtitle: str | None = None,
                    author: str | None = None, date: str | None = None) -> str:
    """Convert Markdown into a styled A4 PDF.

    Args:
        content: Document body in Markdown.
        output_path: Output file path, e.g. ~/Documents/report.pdf.
        theme: Theme name (modern, corporate, academic, fresh).
        title: Adds a cover page with this title.
        subtitle: Cover subtitle.
        author: Cover author.
        date: Cover date.
    """
    path = render_pdf(content, output_path, get_theme(theme), title, subtitle, author, date)
    return f"PDF created: {path}"


@mcp.tool()
def markdown_to_docx(content: str, output_path: str, theme: str = "modern",
                     title: str | None = None, subtitle: str | None = None,
                     author: str | None = None, date: str | None = None) -> str:
    """Convert Markdown into a styled, editable Word document.

    Args:
        content: Document body in Markdown.
        output_path: Output file path, e.g. ~/Documents/report.docx.
        theme: Theme name (modern, corporate, academic, fresh).
        title: Adds a cover page with this title.
        subtitle: Cover subtitle.
        author: Cover author.
        date: Cover date.
    """
    path = render_docx(content, output_path, get_theme(theme), title, subtitle, author, date)
    return f"DOCX created: {path}"


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
