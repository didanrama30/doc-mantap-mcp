"""Render Markdown -> PDF via HTML + CSS (WeasyPrint)."""

from html import escape
from string import Template

from .common import md_to_html, prepare_output
from .themes import Theme

CSS_TEMPLATE = Template("""
@page {
  size: A4;
  margin: 2.5cm 2.2cm 2.5cm 2.2cm;
  @bottom-center { content: counter(page) " / " counter(pages);
                   font-family: $css_body; font-size: 9pt; color: #$muted; }
  @top-right { content: string(doctitle); font-family: $css_body;
               font-size: 8.5pt; color: #$muted; }
}
@page :first { @top-right { content: none; } @bottom-center { content: none; } }

html { font-family: $css_body; font-size: ${body_size_pt}pt; color: #$text; }
body { line-height: 1.65; }

h1, h2, h3, h4, h5, h6 { font-family: $css_heading; color: #$primary;
                         line-height: 1.25; page-break-after: avoid; }
h1 { font-size: 2em; margin: 0 0 .6em; padding-bottom: .25em;
     border-bottom: 3px solid #$accent; string-set: doctitle content(); }
h2 { font-size: 1.45em; margin: 1.6em 0 .5em; padding-left: .5em;
     border-left: 4px solid #$accent; }
h3 { font-size: 1.15em; margin: 1.3em 0 .4em; }
h4, h5, h6 { font-size: 1em; margin: 1em 0 .3em; color: #$text; }

p { margin: 0 0 .8em; text-align: justify; hyphens: auto; }
a { color: #$accent; text-decoration: none; border-bottom: 1px dotted #$accent; }
strong { color: #$primary; }
ul, ol { margin: 0 0 .9em; padding-left: 1.4em; }
li { margin-bottom: .25em; }
li::marker { color: #$accent; }
hr { border: none; border-top: 1px solid #$surface; margin: 1.5em 0; }
img { max-width: 100%; }

blockquote { margin: 1em 0; padding: .8em 1.1em; background: #$surface;
             border-left: 4px solid #$accent; color: #$muted; font-style: italic;
             border-radius: 0 6px 6px 0; }
blockquote p:last-child { margin-bottom: 0; }

code { font-family: "JetBrains Mono", Menlo, Consolas, monospace; font-size: .88em;
       background: #$surface; padding: .1em .35em; border-radius: 3px; }
pre { background: #0F172A; color: #E2E8F0; padding: 1em 1.2em; border-radius: 8px;
      font-size: .85em; line-height: 1.5; white-space: pre-wrap; page-break-inside: avoid; }
pre code { background: none; color: inherit; padding: 0; }

table { width: 100%; border-collapse: collapse; margin: 1em 0 1.3em; font-size: .92em;
        page-break-inside: avoid; }
th { background: #$primary; color: white; text-align: left; padding: .55em .7em;
     font-weight: 600; }
td { padding: .5em .7em; border-bottom: 1px solid #$surface; }
tr:nth-child(even) td { background: #$surface; }

.cover { page-break-after: always; height: 24cm; display: flex;
         flex-direction: column; justify-content: center; }
.cover .bar { width: 4cm; height: 8px; background: #$accent; margin-bottom: 1cm; }
.cover h1.title { font-size: 2.8em; border: none; margin: 0 0 .3em; }
.cover .subtitle { font-size: 1.3em; color: #$muted; margin: 0 0 2cm; }
.cover .meta { font-size: .95em; color: #$muted; border-top: 1px solid #$accent;
               padding-top: .5em; }
""")


def _cover_html(title: str, subtitle: str | None, author: str | None,
                date: str | None) -> str:
    meta = " · ".join(escape(x) for x in (author, date) if x)
    return (
        "<section class='cover'><div class='bar'></div>"
        f"<h1 class='title'>{escape(title)}</h1>"
        + (f"<p class='subtitle'>{escape(subtitle)}</p>" if subtitle else "")
        + (f"<p class='meta'>{meta}</p>" if meta else "")
        + "</section>"
    )


def build_css(theme: Theme) -> str:
    return CSS_TEMPLATE.substitute(theme.to_dict())


def render_pdf(content: str, output_path: str, theme: Theme, title: str | None = None,
               subtitle: str | None = None, author: str | None = None,
               date: str | None = None) -> str:
    # Lazy import so DOCX still works without the PDF system libraries.
    try:
        from weasyprint import CSS, HTML
    except OSError as exc:
        raise RuntimeError(
            "WeasyPrint needs the Pango system library. On macOS: `brew install pango`. "
            f"Details: {exc}"
        ) from exc

    out = prepare_output(output_path, ".pdf")
    cover = _cover_html(title, subtitle, author, date) if title else ""
    doc_title = escape(title) if title else "Document"
    html = (f"<!doctype html><html lang='en'><head><meta charset='utf-8'>"
            f"<title>{doc_title}</title></head><body>{cover}{md_to_html(content)}"
            "</body></html>")
    HTML(string=html).write_pdf(str(out), stylesheets=[CSS(string=build_css(theme))])
    return str(out)
