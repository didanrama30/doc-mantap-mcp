"""Themes shared by the PDF, DOCX and Excel renderers."""

import re
from dataclasses import dataclass, replace

_HEX_RE = re.compile(r"^#?([0-9a-fA-F]{6}|[0-9a-fA-F]{3})$")

# CSS fallback stacks for the bundled fonts (used when a glyph is missing).
FONT_STACKS = {
    "Inter": '"Inter", "Helvetica Neue", Arial, sans-serif',
    "Plus Jakarta Sans": '"Plus Jakarta Sans", "Helvetica Neue", Arial, sans-serif',
    "Source Serif 4": '"Source Serif 4", Georgia, "Times New Roman", serif',
    "Playfair Display": '"Playfair Display", Georgia, "Times New Roman", serif',
    "JetBrains Mono": '"JetBrains Mono", Menlo, Consolas, monospace',
}

COVER_STYLES = ("band", "full", "classic", "minimal")


def normalize_hex(value: str) -> str:
    """'#1e3a8a', '1E3A8A' or '#13a' -> '1E3A8A'."""
    match = _HEX_RE.match(str(value).strip())
    if not match:
        raise ValueError(f"Invalid color '{value}'. Use a hex color like #1E3A8A.")
    digits = match.group(1)
    if len(digits) == 3:
        digits = "".join(c * 2 for c in digits)
    return digits.upper()


def _rgb(color: str) -> tuple[int, int, int]:
    return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)


def mix(a: str, b: str, t: float) -> str:
    """Blend color a toward b (t=0 gives a, t=1 gives b)."""
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(_rgb(a), _rgb(b)))


def tint(color: str, t: float) -> str:
    return mix(color, "FFFFFF", t)


def shade(color: str, t: float) -> str:
    return mix(color, "000000", t)


def luminance(color: str) -> float:
    def channel(c: int) -> float:
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (channel(c) for c in _rgb(color))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def on_color(fill: str) -> str:
    """Readable text color (white or near-black) on top of a fill color."""
    return "FFFFFF" if luminance(fill) < 0.4 else "111827"


@dataclass(frozen=True)
class Theme:
    name: str
    description: str
    # Bundled fonts used by PDF output and charts
    font_body: str
    font_heading: str
    # Office-safe fonts used by DOCX and Excel output
    docx_body: str
    docx_heading: str
    # Colors (hex without '#')
    primary: str
    accent: str
    text: str
    muted: str
    surface: str
    secondary: str  # second brand color: gradients and the second chart series
    palette: tuple[str, ...]  # chart / series colors
    cover: str = "band"
    body_size_pt: float = 10.5
    justify: bool = False
    heading_style: str = "bar"  # bar | rule
    table_style: str = "filled"  # filled | lined
    font_mono: str = "JetBrains Mono"

    @property
    def border(self) -> str:
        return mix(self.surface, self.muted, 0.28)

    @property
    def zebra(self) -> str:
        return mix(self.surface, "FFFFFF", 0.45)

    def stack(self, family: str) -> str:
        return FONT_STACKS.get(family, f'"{family}", sans-serif')

    def with_colors(self, primary: str | None = None, accent: str | None = None) -> "Theme":
        """Return a copy using brand colors; surface and palette are derived."""
        if not primary and not accent:
            return self
        p = normalize_hex(primary) if primary else self.primary
        a = normalize_hex(accent) if accent else self.accent
        palette = (a, p) + tuple(c for c in self.palette if c not in (a, p))
        return replace(self, primary=p, accent=a, surface=tint(a, 0.93),
                       secondary=a, palette=palette[:6])

    def summary(self) -> dict:
        return {
            "name": self.name, "description": self.description,
            "primary": f"#{self.primary}", "accent": f"#{self.accent}",
            "fonts": {"pdf": f"{self.font_heading} / {self.font_body}",
                      "docx_xlsx": f"{self.docx_heading} / {self.docx_body}"},
            "default_cover": self.cover,
        }


THEMES: dict[str, Theme] = {t.name: t for t in (
    Theme(
        name="modern",
        description="Clean and minimal with blue accents. Good for reports and proposals.",
        font_body="Inter", font_heading="Inter",
        docx_body="Calibri", docx_heading="Calibri",
        primary="1E3A8A", accent="3B82F6", text="1E293B", muted="64748B",
        surface="F1F5F9", secondary="0EA5E9",
        palette=("2563EB", "0EA5E9", "1E3A8A", "F59E0B", "10B981", "8B5CF6"),
        cover="band",
    ),
    Theme(
        name="corporate",
        description="Formal navy and gold with serif headings. Good for business documents.",
        font_body="Inter", font_heading="Source Serif 4",
        docx_body="Calibri", docx_heading="Georgia",
        primary="0F172A", accent="B8860B", text="1F2937", muted="6B7280",
        surface="F8F5EC", secondary="1E293B",
        palette=("0F172A", "B8860B", "475569", "D4A72C", "64748B", "94A3B8"),
        cover="full", justify=True,
    ),
    Theme(
        name="academic",
        description="Classic serif with book-style tables. Good for papers, theses and journals.",
        font_body="Source Serif 4", font_heading="Source Serif 4",
        docx_body="Cambria", docx_heading="Cambria",
        primary="111827", accent="7F1D1D", text="111827", muted="4B5563",
        surface="F3F4F6", secondary="374151",
        palette=("7F1D1D", "1F2937", "B45309", "4B5563", "9CA3AF", "D97706"),
        cover="minimal", body_size_pt=11, justify=True,
        heading_style="rule", table_style="lined",
    ),
    Theme(
        name="fresh",
        description="Bright teal and friendly type. Good for learning materials and handouts.",
        font_body="Plus Jakarta Sans", font_heading="Plus Jakarta Sans",
        docx_body="Calibri", docx_heading="Calibri",
        primary="0F766E", accent="14B8A6", text="134E4A", muted="5F7D7A",
        surface="ECFDF5", secondary="0EA5E9",
        palette=("0D9488", "14B8A6", "F59E0B", "0EA5E9", "84CC16", "F43F5E"),
        cover="band",
    ),
    Theme(
        name="elegant",
        description="Plum and gold with Playfair headings. Good for proposals, invitations and portfolios.",
        font_body="Source Serif 4", font_heading="Playfair Display",
        docx_body="Cambria", docx_heading="Georgia",
        primary="5B1A2E", accent="B08D57", text="2B2024", muted="7A6A70",
        surface="F7F2EE", secondary="8C5E6D",
        palette=("5B1A2E", "B08D57", "8C5E6D", "D4B483", "3F3B3C", "A3A3A3"),
        cover="classic", justify=True, heading_style="rule", table_style="lined",
    ),
    Theme(
        name="minimal",
        description="Black and white, lots of whitespace. Good for memos, briefs and CVs.",
        font_body="Inter", font_heading="Inter",
        docx_body="Arial", docx_heading="Arial",
        primary="111111", accent="404040", text="1F1F1F", muted="6B6B6B",
        surface="F5F5F5", secondary="6B6B6B",
        palette=("111111", "737373", "A3A3A3", "404040", "D4D4D4", "8A8A8A"),
        cover="minimal", heading_style="rule", table_style="lined",
    ),
    Theme(
        name="startup",
        description="Violet-to-pink gradient, bold and modern. Good for pitch decks and product docs.",
        font_body="Plus Jakarta Sans", font_heading="Plus Jakarta Sans",
        docx_body="Calibri", docx_heading="Century Gothic",
        primary="4C1D95", accent="7C3AED", text="1E1B4B", muted="6B7280",
        surface="F5F3FF", secondary="DB2777",
        palette=("7C3AED", "EC4899", "06B6D4", "F59E0B", "10B981", "6366F1"),
        cover="full",
    ),
    Theme(
        name="warm",
        description="Terracotta and orange with serif body text. Good for stories, newsletters and plans.",
        font_body="Source Serif 4", font_heading="Plus Jakarta Sans",
        docx_body="Georgia", docx_heading="Century Gothic",
        primary="7C2D12", accent="EA580C", text="292524", muted="78716C",
        surface="FFF7ED", secondary="F59E0B",
        palette=("EA580C", "7C2D12", "F59E0B", "0F766E", "DC2626", "A16207"),
        cover="band",
    ),
)}


def get_theme(name: str, primary_color: str | None = None,
              accent_color: str | None = None) -> Theme:
    try:
        theme = THEMES[(name or "modern").strip().lower()]
    except KeyError:
        raise ValueError(
            f"Unknown theme '{name}'. Available: {', '.join(THEMES)}"
        ) from None
    return theme.with_colors(primary_color, accent_color)
