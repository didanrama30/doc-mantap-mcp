"""Syntax highlighting shared by the PDF and DOCX renderers (Pygments)."""

from __future__ import annotations

from functools import lru_cache

from pygments import highlight, lex
from pygments.formatters import HtmlFormatter
from pygments.lexers import TextLexer, get_lexer_by_name
from pygments.styles import get_style_by_name
from pygments.util import ClassNotFound

STYLE_NAME = "github-dark"
CODE_BG = "0F172A"
CODE_FG = "E2E8F0"
CODE_LABEL = "94A3B8"
PLAIN_LANGS = {"", "text", "txt", "plain", "plaintext", "none", "output", "console-output"}


def _lexer(lang: str):
    if lang and lang.lower() not in PLAIN_LANGS:
        try:
            return get_lexer_by_name(lang, stripnl=False)
        except ClassNotFound:
            pass
    return TextLexer(stripnl=False)


def code_label(lang: str) -> str:
    """Short label shown on a code block ('' for plain text)."""
    if not lang or lang.lower() in PLAIN_LANGS:
        return ""
    return lang.upper() if len(lang) <= 4 else lang.capitalize()


def to_html(code: str, lang: str) -> str:
    return highlight(code, _lexer(lang), HtmlFormatter(nowrap=True))


@lru_cache(maxsize=1)
def css(selector: str = ".code pre") -> str:
    rules = HtmlFormatter(style=STYLE_NAME).get_style_defs(selector)
    # Keep token colors only; the block background comes from our own CSS.
    return "\n".join(r for r in rules.splitlines()
                     if not r.startswith((f"{selector} {{", f"{selector} .hll", "pre {", "td.", "span.linenos")))


def tokens(code: str, lang: str):
    """Yield (text, color_hex | None, bold, italic) for DOCX runs."""
    style = get_style_by_name(STYLE_NAME)
    for ttype, value in lex(code, _lexer(lang)):
        info = style.style_for_token(ttype)
        yield value, info["color"], bool(info["bold"]), bool(info["italic"])
