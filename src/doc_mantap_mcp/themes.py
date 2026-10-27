"""Themes shared by the PDF and DOCX renderers."""

from dataclasses import dataclass, asdict


@dataclass(frozen=True)
class Theme:
    name: str
    description: str
    # PDF fonts (CSS font stack)
    css_body: str
    css_heading: str
    # DOCX fonts (single font name)
    docx_body: str
    docx_heading: str
    # Colors (hex without '#')
    primary: str
    accent: str
    text: str
    muted: str
    surface: str
    body_size_pt: float = 11

    def to_dict(self) -> dict:
        return asdict(self)


THEMES: dict[str, Theme] = {
    "modern": Theme(
        name="modern",
        description="Clean and minimal with blue accents. Good for reports and proposals.",
        css_body='"Inter", "Helvetica Neue", Arial, sans-serif',
        css_heading='"Inter", "Helvetica Neue", Arial, sans-serif',
        docx_body="Calibri",
        docx_heading="Calibri",
        primary="1E3A8A",
        accent="3B82F6",
        text="1E293B",
        muted="64748B",
        surface="F1F5F9",
    ),
    "corporate": Theme(
        name="corporate",
        description="Formal and elegant, navy and gold. Good for business documents.",
        css_body='"Helvetica Neue", Arial, sans-serif',
        css_heading='Georgia, "Times New Roman", serif',
        docx_body="Arial",
        docx_heading="Georgia",
        primary="0F172A",
        accent="B8860B",
        text="1F2937",
        muted="6B7280",
        surface="F8F5EC",
    ),
    "academic": Theme(
        name="academic",
        description="Classic serif. Good for papers, theses and journals.",
        css_body='Georgia, "Times New Roman", serif',
        css_heading='Georgia, "Times New Roman", serif',
        docx_body="Times New Roman",
        docx_heading="Times New Roman",
        primary="111827",
        accent="7F1D1D",
        text="111827",
        muted="4B5563",
        surface="F3F4F6",
        body_size_pt=12,
    ),
    "fresh": Theme(
        name="fresh",
        description="Bright teal. Good for learning materials and handouts.",
        css_body='"Helvetica Neue", Arial, sans-serif',
        css_heading='"Helvetica Neue", Arial, sans-serif',
        docx_body="Calibri",
        docx_heading="Calibri",
        primary="0F766E",
        accent="14B8A6",
        text="134E4A",
        muted="5F7D7A",
        surface="ECFDF5",
    ),
}


def get_theme(name: str) -> Theme:
    try:
        return THEMES[name.lower()]
    except KeyError:
        raise ValueError(
            f"Unknown theme '{name}'. Available: {', '.join(THEMES)}"
        ) from None
