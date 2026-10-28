"""Render Markdown -> PDF via HTML + CSS (WeasyPrint)."""

from __future__ import annotations

import base64
import contextlib
import os
import sys
from dataclasses import dataclass
from html import escape
from pathlib import Path
from string import Template
from urllib.parse import urlparse
from urllib.request import url2pathname

from bs4 import BeautifulSoup, Tag

from . import highlight
from .charts import render_chart
from .common import CALLOUTS, IMAGE_EXTS, ParsedDoc, labels, parse_markdown, prepare_output
from .fonts import FONT_DIR, font_face_css
from .themes import COVER_STYLES, Theme, on_color, shade, tint

PAGE_SIZES_MM = {"a4": (210, 297), "a5": (148, 210), "letter": (215.9, 279.4),
                 "legal": (215.9, 355.6)}

PANGO_HELP = (
    "PDF output needs the Pango library. macOS: `brew install pango`. "
    "Debian/Ubuntu: `apt install libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz-subset0`. "
    "Windows: see https://doc.courtbouillon.org/weasyprint/stable/first_steps.html. "
    "DOCX and Excel output work without it."
)


@dataclass
class PageSetup:
    width_mm: float
    height_mm: float
    margin_side_mm: float
    margin_top_mm: float
    margin_bottom_mm: float

    @property
    def text_width_cm(self) -> float:
        return (self.width_mm - 2 * self.margin_side_mm) / 10


def page_setup(page_size: str = "A4", orientation: str = "portrait") -> PageSetup:
    key = (page_size or "A4").strip().lower()
    if key not in PAGE_SIZES_MM:
        raise ValueError(f"Unknown page_size '{page_size}'. Use A4, A5, Letter or Legal.")
    w, h = PAGE_SIZES_MM[key]
    orient = (orientation or "portrait").strip().lower()
    if orient not in ("portrait", "landscape"):
        raise ValueError("orientation must be 'portrait' or 'landscape'.")
    if orient == "landscape":
        w, h = h, w
    side = 15 if min(w, h) < 200 else 22
    return PageSetup(w, h, side, side + 2, side + 2)


def css_string(text: str) -> str:
    return '"' + str(text).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ") + '"'


# --------------------------------------------------------------------------- #
# CSS
# --------------------------------------------------------------------------- #

BASE_CSS = Template("""
@page {
  size: ${page_w}mm ${page_h}mm;
  margin: ${m_top}mm ${m_side}mm ${m_bottom}mm ${m_side}mm;
  @top-left { content: $header_left; font-family: $body_stack; font-size: 7pt;
              color: #$muted; letter-spacing: .08em; text-transform: uppercase;
              vertical-align: bottom; padding-bottom: 5mm; }
  @top-right { content: string(section); font-family: $body_stack; font-size: 7pt;
               color: #$muted; vertical-align: bottom; padding-bottom: 5mm; }
  @bottom-left { content: $footer_left; font-family: $body_stack; font-size: 7pt;
                 color: #$muted; vertical-align: top; padding-top: 5mm; }
  @bottom-right { content: counter(page) " / " counter(pages); font-family: $body_stack;
                  font-size: 7.5pt; font-weight: 700; color: #$primary;
                  vertical-align: top; padding-top: 5mm; }
  @footnote { border-top: .6pt solid #$border; padding-top: 5pt; margin-top: 10pt; }
}
@page :first { @top-left { content: none; } @top-right { content: none; } }
@page cover { margin: 0;
  @top-left { content: none; } @top-right { content: none; }
  @bottom-left { content: none; } @bottom-right { content: none; } }

html { font-family: $body_stack; font-size: ${base}pt; color: #$text; line-height: 1.62; }
body { margin: 0; }
h1, h2, h3, h4, h5, h6 { font-family: $heading_stack; color: #$primary; line-height: 1.22;
  font-weight: $hw; break-after: avoid; font-variant-numeric: lining-nums; hyphens: manual; }
h1 { font-size: 2.05em; margin: 0 0 .75em; letter-spacing: -.012em; string-set: section content(text); }
* + h1 { margin-top: 1.5em; }
h2 { font-size: 1.42em; margin: 1.75em 0 .6em; letter-spacing: -.005em; }
h3 { font-size: 1.13em; margin: 1.4em 0 .4em; color: #$h3; }
h4 { font-family: $body_stack; font-size: .8em; margin: 1.35em 0 .35em; color: #$muted;
     text-transform: uppercase; letter-spacing: .09em; font-weight: 700; }
h5, h6 { font-size: .95em; margin: 1em 0 .3em; color: #$text; }
h1 + h2, h2 + h3, h3 + h4 { margin-top: .5em; }
.no-h1 h2 { string-set: section content(text); }

p { margin: 0 0 .8em; text-align: $align; hyphens: auto; orphans: 2; widows: 2; }
a { color: #$link; text-decoration: none; }
strong { font-weight: 700; color: #$strong; }
mark { background: #$mark; color: inherit; padding: 0 .12em; border-radius: 2pt; }
del { color: #$muted; }
ins { text-decoration: underline; text-decoration-color: #$accent; }
sup, sub { font-size: .7em; line-height: 0; }
abbr { text-decoration: none; border-bottom: .5pt dotted #$muted; }
hr { border: none; border-top: .8pt solid #$border; margin: 1.8em 0; }
img { max-width: 100%; }

ul, ol { margin: 0 0 .9em; padding-left: 1.35em; }
li { margin-bottom: .28em; }
li > p { margin-bottom: .3em; text-align: left; }
ul li::marker { color: #$accent; }
ol li::marker { color: #$accent; font-weight: 700; }
li ul, li ol { margin: .28em 0 .2em; }
ul.task-list { list-style: none; padding-left: .15em; }
.task-box { display: inline-block; width: .82em; height: .82em; border: 1pt solid #$muted;
            border-radius: 2.5pt; margin-right: .5em; vertical-align: -.1em; }
.task-box.done { background: #$accent; border-color: #$accent; }
.task-box.done::after { content: "✓"; display: block; color: #$accent_on; font-size: .72em;
                        line-height: 1.12em; text-align: center; font-weight: 800; }
dt { font-weight: 700; color: #$primary; margin-top: .6em; }
dd { margin: .15em 0 .45em 1.2em; }

table { width: 100%; border-collapse: separate; border-spacing: 0; margin: 1.1em 0 1.5em;
        font-size: .88em; line-height: 1.45; font-variant-numeric: tabular-nums lining-nums; }
thead { display: table-header-group; }
tr { break-inside: avoid; }
th { background: #$primary; color: #$primary_on; font-weight: 700; text-align: left;
     padding: .62em .8em; font-size: .94em; }
td { padding: .55em .8em; border-bottom: .6pt solid #$border; vertical-align: top; }
tbody tr:nth-child(even) td { background: #$zebra; }
tbody tr:last-child td { border-bottom: 1.2pt solid #$primary; }
td strong { color: inherit; }
td p { margin: 0; text-align: inherit; }

blockquote { margin: 1.3em 0; padding: .1em 0 .1em 1.15em; border-left: 2.5pt solid #$accent;
             color: #$quote; font-style: italic; font-size: 1.04em; }
blockquote p { text-align: left; }
blockquote > :last-child, .callout > :last-child { margin-bottom: 0; }

.callout { margin: 1.15em 0 1.35em; padding: .8em 1.05em .8em 1.05em; border-radius: 0 5pt 5pt 0;
           border-left: 3.5pt solid; break-inside: avoid; }
.callout p, .callout li { text-align: left; }
.callout-title { font-weight: 700; font-size: .9em; margin: 0 0 .35em; }
.callout-title::before { display: inline-block; width: 1.32em; height: 1.32em; line-height: 1.34em;
  border-radius: 50%; text-align: center; font-size: .78em; color: white; margin-right: .55em;
  vertical-align: .06em; font-weight: 800; }
$callout_css

.kpis { display: flex; gap: 8pt; margin: 1.1em 0 1.4em; break-inside: avoid; }
.kpi { flex: 1 1 0; background: #$surface; border-radius: 6pt; padding: .85em .95em .85em; }
.kpi::before { content: ""; display: block; width: 1.7em; height: 2.6pt; border-radius: 2pt;
               background: #$accent; margin-bottom: .65em; }
.kpi-label { font-size: .68em; text-transform: uppercase; letter-spacing: .08em; color: #$muted;
             font-weight: 700; line-height: 1.3; }
.kpi-value { font-family: $heading_stack; font-size: 1.7em; font-weight: 800; color: #$primary;
             line-height: 1.15; margin-top: .3em; font-variant-numeric: lining-nums; }
.kpi-delta { display: inline-block; margin-top: .55em; font-size: .7em; font-weight: 700;
             padding: .16em .6em; border-radius: 8pt; }
.kpi-delta.good { color: #047857; background: #D1FAE5; }
.kpi-delta.bad { color: #B91C1C; background: #FEE2E2; }
.kpi-delta.flat { color: #$muted; background: #FFFFFF; }
.kpi-note { font-size: .68em; color: #$muted; margin-top: .45em; line-height: 1.35; }

.code { margin: 1em 0 1.35em; background: #$code_bg; border-radius: 6pt; break-inside: avoid; }
.code-lang { font-family: $mono_stack; font-size: .58em; letter-spacing: .14em; color: #$code_label;
             padding: 1.1em 1.6em 0; font-weight: 700; }
.code pre { margin: 0; padding: .9em 1.25em 1em; font-family: $mono_stack; font-size: .78em;
            line-height: 1.6; color: #$code_fg; white-space: pre-wrap; overflow-wrap: anywhere;
            font-feature-settings: "liga" 0, "calt" 0; }
.code-lang + pre { padding-top: .35em; }
code { font-family: $mono_stack; font-size: .86em; background: #$surface; color: #$code_inline;
       padding: .06em .32em; border-radius: 3pt; font-feature-settings: "liga" 0, "calt" 0; }
pre code { background: none; color: inherit; padding: 0; font-size: inherit; border-radius: 0; }
$pygments_css

figure { margin: 1.2em 0 1.5em; break-inside: avoid; }
figure img { display: block; margin: 0 auto; max-width: 100%; }
figure.chart img { margin-top: .55em; }
.chart-title { font-family: $heading_stack; font-weight: 700; font-size: 1em; color: #$text; }
.chart-subtitle { font-size: .8em; color: #$muted; margin-top: .1em; }
figcaption { font-size: .76em; color: #$muted; margin-top: .6em; text-align: center; }
figure.chart figcaption { text-align: left; }
.page-break { break-after: page; height: 0; margin: 0; }

.fn { float: footnote; font-size: .76em; line-height: 1.45; color: #$text; text-align: left;
      font-style: normal; font-weight: normal; }
.fn::footnote-call { content: counter(footnote); vertical-align: super; font-size: .68em;
                     line-height: 0; color: #$accent; font-weight: 700; padding-left: .06em; }
.fn::footnote-marker { content: counter(footnote) ".  "; color: #$accent; font-weight: 700; }

.toc { break-after: page; }
.toc-title { font-family: $heading_stack; font-size: 1.9em; font-weight: $hw; color: #$primary;
             margin: .4cm 0 1.1em; line-height: 1.2; }
.toc ol { list-style: none; padding: 0; margin: 0; }
.toc li { margin: 0; }
.toc a { display: block; color: #$text; padding: .34em 0; }
.toc a::after { content: leader('.') "  " target-counter(attr(href), page); color: #$muted;
                font-weight: 400; }
.toc .l1 > a { font-weight: 700; color: #$primary; padding-top: .75em; }
.toc .l1:first-child > a { padding-top: 0; }
.toc .l2 > a { padding-left: 1.1em; }
.toc .l3 > a { padding-left: 2.2em; font-size: .92em; color: #$muted; }
""")

HEADING_CSS = {
    "bar": Template("""
h1 { padding-bottom: .28em; border-bottom: 2.5pt solid #$accent; }
h2 { padding-left: .55em; border-left: 3.5pt solid #$accent; }
"""),
    "rule": Template("""
h1 { padding-bottom: .35em; border-bottom: .8pt solid #$primary; }
h2::after { content: ""; display: block; width: 2.2em; border-bottom: 1.6pt solid #$accent;
            margin-top: .3em; }
"""),
}

LINED_TABLE_CSS = Template("""
th { background: none; color: #$text; border-top: 1.4pt solid #$text;
     border-bottom: .8pt solid #$text; padding: .55em .7em; }
tbody tr:nth-child(even) td { background: none; }
td { border-bottom: .4pt solid #$border; padding: .5em .7em; }
tbody tr:last-child td { border-bottom: 1.4pt solid #$text; }
""")

COVER_CSS = Template("""
.cover { page: cover; position: relative; width: ${page_w}mm; height: ${page_h}mm;
         overflow: hidden; break-after: page; color: #$text; }
.cv-kicker { font-size: 8.5pt; font-weight: 700; letter-spacing: .22em; text-transform: uppercase; }
.cv-title { font-family: $heading_stack; font-weight: $title_w; font-size: ${title_size}pt;
            line-height: 1.08; letter-spacing: -.02em; margin: .2em 0 0;
            font-variant-numeric: lining-nums; hyphens: manual; }
.cv-subtitle { font-size: 13pt; line-height: 1.4; margin-top: .7em; }
.cv-meta { display: flex; gap: 1.3cm; }
.cv-meta .lbl { display: block; font-size: 6.8pt; letter-spacing: .15em; text-transform: uppercase;
                font-weight: 700; color: #$muted; margin-bottom: .35em; }
.cv-meta .val { font-size: 10.5pt; font-weight: 600; }
.cv-logo img { max-height: 1.5cm; max-width: 6cm; }

.cover-band .cv-panel { position: absolute; top: 0; left: 0; width: ${page_w}mm; height: 58%;
                        background: #$primary; color: #$primary_on; overflow: hidden; }
.cover-band .cv-panel::before { content: ""; position: absolute; width: 13cm; height: 13cm;
  right: -4.5cm; top: -5.5cm; border-radius: 50%; background: rgba(255,255,255,.07); }
.cover-band .cv-panel::after { content: ""; position: absolute; width: 7cm; height: 7cm;
  right: -1.8cm; bottom: -4.4cm; border-radius: 50%; border: 1.2pt solid rgba(255,255,255,.2); }
.cover-band .cv-head { position: absolute; left: ${m_side}mm; right: ${m_side}mm; bottom: 2.1cm; }
.cover-band .cv-kicker { color: #$kicker_dark; }
.cover-band .cv-subtitle { color: rgba(255,255,255,.82); }
.cover-band .cv-bar { position: absolute; left: ${m_side}mm; top: 58%; width: 3.2cm; height: 7pt;
                      margin-top: -3.5pt; background: #$accent; }
.cover-band .cv-meta { position: absolute; left: ${m_side}mm; right: ${m_side}mm; top: 66%; }
.cover-band .cv-logo { position: absolute; left: ${m_side}mm; bottom: 2cm; }

.cover-full { background: linear-gradient(145deg, #$primary 0%, #$grad_end 100%);
              color: #$primary_on; }
.cover-full::before { content: ""; position: absolute; width: 16cm; height: 16cm; right: -6cm;
  top: -5cm; border-radius: 50%; background: rgba(255,255,255,.06); }
.cover-full::after { content: ""; position: absolute; width: 10cm; height: 10cm; right: -3.5cm;
  bottom: -4.6cm; border-radius: 50%; border: 1.2pt solid rgba(255,255,255,.16); }
.cover-full .cv-logo { position: absolute; left: ${m_side}mm; top: 2.2cm; }
.cover-full .cv-head { position: absolute; left: ${m_side}mm; right: ${m_side}mm; top: 36%; }
.cover-full .cv-kicker { color: #$kicker_dark; }
.cover-full .cv-rule { width: 2.6cm; height: 4pt; background: #$kicker_dark; margin-bottom: .9cm; }
.cover-full .cv-subtitle { color: rgba(255,255,255,.82); }
.cover-full .cv-meta { position: absolute; left: ${m_side}mm; right: ${m_side}mm; bottom: 2.2cm;
                       padding-top: .5cm; border-top: .8pt solid rgba(255,255,255,.28); }
.cover-full .cv-meta .lbl { color: rgba(255,255,255,.6); }

.cover-classic .cv-spine { position: absolute; left: 0; top: 0; width: 1.1cm; height: ${page_h}mm;
                           background: #$primary; }
.cover-classic .cv-spine::after { content: ""; position: absolute; left: 1.45cm; top: 0;
  width: 1.2pt; height: ${page_h}mm; background: #$accent; }
.cover-classic .cv-logo { position: absolute; left: 3.4cm; top: 2.6cm; }
.cover-classic .cv-head { position: absolute; left: 3.4cm; right: ${m_side}mm; top: 33%; }
.cover-classic .cv-kicker { color: #$accent; }
.cover-classic .cv-title { color: #$primary; }
.cover-classic .cv-rule { width: 2.4cm; height: 2pt; background: #$accent; margin-top: .8cm; }
.cover-classic .cv-subtitle { color: #$muted; margin-top: .55cm; }
.cover-classic .cv-meta { position: absolute; left: 3.4cm; right: ${m_side}mm; bottom: 2.4cm; }

.cover-minimal { text-align: center; }
.cover-minimal .cv-logo { position: absolute; left: 0; right: 0; top: 2.6cm; }
.cover-minimal .cv-logo img { margin: 0 auto; }
.cover-minimal .cv-head { position: absolute; left: 2.6cm; right: 2.6cm; top: 34%; }
.cover-minimal .cv-kicker { color: #$muted; }
.cover-minimal .cv-title { color: #$primary; letter-spacing: -.01em; font-size: ${title_size_min}pt; }
.cover-minimal .cv-rule { width: 1.6cm; height: 1.4pt; background: #$accent; margin: .75cm auto 0; }
.cover-minimal .cv-subtitle { color: #$muted; font-style: italic; margin-top: .6cm; }
.cover-minimal .cv-meta { position: absolute; left: 2cm; right: 2cm; bottom: 2.4cm;
                          justify-content: center; }
""")


def _callout_css() -> str:
    rules = []
    for kind, (color, icon) in CALLOUTS.items():
        rules.append(
            f".callout-{kind} {{ background: #{tint(color, 0.92)}; border-color: #{color}; }}\n"
            f".callout-{kind} .callout-title {{ color: #{shade(color, 0.18)}; }}\n"
            f".callout-{kind} .callout-title::before {{ content: {css_string(icon)}; "
            f"background: #{color}; }}")
    return "\n".join(rules)


def build_css(theme: Theme, setup: PageSetup, header_left: str = "", footer_left: str = "",
              title_size: float = 32) -> str:
    values = {
        "page_w": setup.width_mm, "page_h": setup.height_mm, "m_side": setup.margin_side_mm,
        "m_top": setup.margin_top_mm, "m_bottom": setup.margin_bottom_mm,
        "header_left": css_string(header_left), "footer_left": css_string(footer_left),
        "body_stack": theme.stack(theme.font_body), "heading_stack": theme.stack(theme.font_heading),
        "mono_stack": theme.stack(theme.font_mono),
        "base": theme.body_size_pt * (0.92 if setup.width_mm < 160 else 1),
        "hw": 600 if theme.font_heading == "Playfair Display" else 700,
        "title_w": 700 if theme.font_heading == "Playfair Display" else 800,
        "title_size": title_size, "title_size_min": title_size - 5,
        "text": theme.text, "muted": theme.muted, "primary": theme.primary,
        "accent": theme.accent, "surface": theme.surface, "border": theme.border,
        "zebra": theme.zebra, "primary_on": on_color(theme.primary),
        "accent_on": on_color(theme.accent), "h3": shade(theme.primary, 0.05),
        "link": shade(theme.accent, 0.1), "strong": theme.primary,
        "mark": tint("FACC15", 0.55), "quote": theme.muted,
        "align": "justify" if theme.justify else "left",
        "code_bg": highlight.CODE_BG, "code_fg": highlight.CODE_FG,
        "code_label": highlight.CODE_LABEL, "code_inline": shade(theme.accent, 0.25),
        "kicker_dark": tint(theme.accent, 0.45), "grad_end": theme.secondary,
        "callout_css": _callout_css(), "pygments_css": highlight.css(),
    }
    css = BASE_CSS.substitute(values) + HEADING_CSS[theme.heading_style].substitute(values)
    if theme.table_style == "lined":
        css += LINED_TABLE_CSS.substitute(values)
    css += COVER_CSS.substitute(values)
    families = {theme.font_body, theme.font_heading, theme.font_mono}
    return font_face_css(families) + css


# --------------------------------------------------------------------------- #
# HTML building
# --------------------------------------------------------------------------- #

def _logo_uri(logo: str | None, base_dir: Path) -> str | None:
    if not logo:
        return None
    if urlparse(logo).scheme in ("http", "https", "data"):
        return logo
    path = Path(logo).expanduser()
    path = (path if path.is_absolute() else base_dir / path).resolve()
    if not path.is_file():
        raise ValueError(f"Logo file not found: {path}")
    return path.as_uri()


def cover_html(style: str, names: dict, title: str, subtitle: str | None, author: str | None,
               date: str | None, organization: str | None, document_type: str | None,
               logo_uri: str | None) -> str:
    meta = "".join(
        f"<div><span class='lbl'>{escape(label)}</span><span class='val'>{escape(value)}</span></div>"
        for label, value in ((names["prepared_by"], author), (names["date"], date),
                             (names["organization"], organization)) if value)
    logo = f"<div class='cv-logo'><img src='{escape(logo_uri)}' alt=''></div>" if logo_uri else ""
    kicker = f"<div class='cv-kicker'>{escape(document_type)}</div>" if document_type else ""
    sub = f"<div class='cv-subtitle'>{escape(subtitle)}</div>" if subtitle else ""
    rule = "<div class='cv-rule'></div>"
    head_title = f"<div class='cv-title'>{escape(title)}</div>"
    meta_html = f"<div class='cv-meta'>{meta}</div>" if meta else ""
    if style == "band":
        return (f"<section class='cover cover-band'><div class='cv-panel'><div class='cv-head'>"
                f"{kicker}{head_title}{sub}</div></div><div class='cv-bar'></div>"
                f"{meta_html}{logo}</section>")
    if style == "full":
        return (f"<section class='cover cover-full'>{logo}<div class='cv-head'>{rule}{kicker}"
                f"{head_title}{sub}</div>{meta_html}</section>")
    if style == "classic":
        return (f"<section class='cover cover-classic'><div class='cv-spine'></div>{logo}"
                f"<div class='cv-head'>{kicker}{head_title}{rule}{sub}</div>{meta_html}</section>")
    return (f"<section class='cover cover-minimal'>{logo}<div class='cv-head'>{kicker}{head_title}"
            f"{rule}{sub}</div>{meta_html}</section>")


def toc_html(doc: ParsedDoc) -> str:
    heads = [h for h in doc.headings if h.id]
    if not heads:
        return ""
    top = min(h.level for h in heads)
    depth = 3 if sum(1 for h in heads if h.level <= top + 1) < 12 else 2
    items = "".join(
        f"<li class='l{h.level - top + 1}'><a href='#{escape(h.id)}'>{escape(h.text)}</a></li>"
        for h in heads if h.level - top < depth)
    return (f"<nav class='toc'><div class='toc-title'>{escape(doc.labels['contents'])}</div>"
            f"<ol>{items}</ol></nav>")


def _kpi_rows(items: list[dict]) -> list[list[dict]]:
    n = len(items)
    per_row = n if n <= 4 else (3 if n in (5, 6, 9) else 4)
    return [items[i:i + per_row] for i in range(0, n, per_row)]


def _kpis_html(items: list[dict]) -> str:
    rows = []
    for row in _kpi_rows(items):
        cards = []
        for k in row:
            arrow = {"up": "▲ ", "down": "▼ "}.get(k["trend"], "")
            delta = k["delta"] or ""
            if arrow and delta[:1] in "▲▼↑↓":
                arrow = ""
            cards.append(
                "<div class='kpi'>"
                f"<div class='kpi-label'>{escape(k['label'])}</div>"
                f"<div class='kpi-value'>{escape(k['value'])}</div>"
                + (f"<div class='kpi-delta {k['tone']}'>{escape(arrow + delta)}</div>" if delta else "")
                + (f"<div class='kpi-note'>{escape(k['note'])}</div>" if k["note"] else "")
                + "</div>")
        rows.append(f"<div class='kpis'>{''.join(cards)}</div>")
    return "".join(rows)


def _enhance_for_pdf(doc: ParsedDoc, theme: Theme, setup: PageSetup) -> None:
    soup = doc.soup
    for div in soup.find_all("div", class_="dm-chart"):
        spec = doc.charts[int(div["data-index"])]
        svg = render_chart(spec, theme, "svg", setup.text_width_cm, doc.lang)
        uri = "data:image/svg+xml;base64," + base64.b64encode(svg).decode()
        parts = []
        if spec["title"]:
            parts.append(f"<div class='chart-title'>{escape(str(spec['title']))}</div>")
        if spec["subtitle"]:
            parts.append(f"<div class='chart-subtitle'>{escape(str(spec['subtitle']))}</div>")
        parts.append(f"<img src='{uri}' alt=''>")
        if spec["caption"]:
            parts.append(f"<figcaption>{escape(str(spec['caption']))}</figcaption>")
        div.replace_with(BeautifulSoup(f"<figure class='chart'>{''.join(parts)}</figure>",
                                       "html.parser"))
    for div in soup.find_all("div", class_="dm-kpis"):
        div.replace_with(BeautifulSoup(_kpis_html(doc.kpis[int(div["data-index"])]),
                                       "html.parser"))
    for pre in soup.find_all("pre", attrs={"data-lang": True}):
        lang = pre["data-lang"]
        label = highlight.code_label(lang)
        body = highlight.to_html(pre.get_text().rstrip("\n"), lang)
        html = ("<div class='code'>" + (f"<div class='code-lang'>{escape(label)}</div>" if label
                                        else "") + f"<pre>{body}</pre></div>")
        pre.replace_with(BeautifulSoup(html, "html.parser"))
    for sup in soup.find_all("sup"):
        ref = sup.find("a", class_="footnote-ref")
        if ref is None:
            continue
        note = doc.footnotes.get(ref.get("href", "").lstrip("#"))
        span = soup.new_tag("span", attrs={"class": "fn"})
        if note is not None:
            for i, block in enumerate(note.find_all(["p"], recursive=False) or [note]):
                if i:
                    span.append(" ")
                for child in list(block.contents):
                    span.append(child.__copy__() if isinstance(child, Tag) else str(child))
        sup.replace_with(span)


def _ensure_native_libs() -> None:
    """Let WeasyPrint find Homebrew/MacPorts libraries on macOS."""
    if sys.platform != "darwin":
        return
    current = os.environ.get("DYLD_FALLBACK_LIBRARY_PATH")
    parts = current.split(":") if current else [str(Path.home() / "lib"), "/usr/local/lib",
                                                 "/usr/lib"]
    for extra in ("/opt/homebrew/lib", "/usr/local/lib", "/opt/local/lib"):
        if extra not in parts and Path(extra).is_dir():
            parts.append(extra)
    os.environ["DYLD_FALLBACK_LIBRARY_PATH"] = ":".join(p for p in parts if p)


def _weasyprint():
    _ensure_native_libs()
    try:
        # WeasyPrint prints install hints to stdout, which would corrupt the MCP stdio stream.
        with contextlib.redirect_stdout(sys.stderr):
            import weasyprint
            from weasyprint.text.fonts import FontConfiguration
            from weasyprint.urls import URLFetcher
    except OSError as exc:
        raise RuntimeError(f"{PANGO_HELP} Details: {exc}") from exc

    class SafeFetcher(URLFetcher):
        """Only load images, bundled fonts and inline data; never other local files."""

        def fetch(self, url, headers=None):
            scheme = urlparse(url).scheme.lower()
            if scheme == "data":
                return super().fetch(url, headers)
            if scheme == "file":
                path = Path(url2pathname(urlparse(url).path)).resolve()
                if path.is_relative_to(FONT_DIR.resolve()) or path.suffix.lower() in IMAGE_EXTS:
                    return super().fetch(url, headers)
                raise ValueError(f"Blocked local file (only images are allowed): {path}")
            if scheme in ("http", "https"):
                response = super().fetch(url, headers)
                if not response.content_type.startswith("image/"):
                    response.close()
                    raise ValueError(f"Blocked non-image URL: {url}")
                return response
            raise ValueError(f"Blocked URL scheme: {url}")

    return weasyprint, FontConfiguration, SafeFetcher


def _title_size(title: str) -> float:
    n = len(title)
    return 34 if n <= 28 else 29 if n <= 50 else 24 if n <= 80 else 20


@dataclass
class PdfResult:
    path: Path
    pages: int


def render_pdf(content: str, output_path: str, theme: Theme, title: str | None = None,
               subtitle: str | None = None, author: str | None = None, date: str | None = None,
               *, organization: str | None = None, document_type: str | None = None,
               logo: str | None = None, cover_style: str | None = None, toc: bool = False,
               page_size: str = "A4", orientation: str = "portrait", lang: str = "en",
               base_dir: str | None = None) -> PdfResult:
    weasyprint, FontConfiguration, SafeFetcher = _weasyprint()
    out = prepare_output(output_path, ".pdf")
    base = Path(base_dir).expanduser().resolve() if base_dir else out.parent
    setup = page_setup(page_size, orientation)
    style = (cover_style or theme.cover).lower()
    if style not in COVER_STYLES:
        raise ValueError(f"Unknown cover_style '{cover_style}'. Use: {', '.join(COVER_STYLES)}.")

    doc = parse_markdown(content, lang=lang, base_dir=base)
    _enhance_for_pdf(doc, theme, setup)
    names = labels(lang)
    first_h1 = next((h.text for h in doc.headings if h.level == 1), "")
    header = title or first_h1
    cover = cover_html(style, names, title, subtitle, author, date, organization, document_type,
                       _logo_uri(logo, base)) if title else ""
    body_class = f"theme-{theme.name}" + ("" if any(h.level == 1 for h in doc.headings)
                                          else " no-h1")
    meta = "".join(f"<meta name='{k}' content='{escape(v)}'>" for k, v in (
        ("author", author), ("description", subtitle), ("generator", "doc-mantap-mcp")) if v)
    html = (f"<!doctype html><html lang='{escape(lang)}'><head><meta charset='utf-8'>"
            f"<title>{escape(title or first_h1 or 'Document')}</title>{meta}</head>"
            f"<body class='{body_class}'>{cover}{toc_html(doc) if toc else ''}"
            f"<main>{doc.soup}</main></body></html>")
    css = build_css(theme, setup, header_left=header, footer_left=organization or author or "",
                    title_size=_title_size(title or ""))
    fonts = FontConfiguration()
    fetcher = SafeFetcher(timeout=15)
    rendered = weasyprint.HTML(string=html, base_url=str(base), url_fetcher=fetcher).render(
        stylesheets=[weasyprint.CSS(string=css, font_config=fonts, url_fetcher=fetcher)],
        font_config=fonts)
    rendered.write_pdf(str(out))
    return PdfResult(out, len(rendered.pages))
