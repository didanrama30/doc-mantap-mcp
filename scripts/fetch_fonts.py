"""Download the bundled OFL fonts from Google Fonts and subset them for the PDF renderer.

Run from the repo root:  uv run python scripts/fetch_fonts.py

Each family/weight/style becomes one WOFF2 file in src/doc_mantap_mcp/assets/fonts,
subset to Latin (incl. Vietnamese), Greek, punctuation, currency, arrows, math, box
drawing and common symbols. OpenType layout features (kerning, ligatures, tabular
figures) are kept. A fonts.json manifest is written for the renderer.
"""

import io
import json
import re
import urllib.request
from pathlib import Path

from fontTools import subset
from fontTools.ttLib import TTFont

DEST = Path(__file__).resolve().parents[1] / "src" / "doc_mantap_mcp" / "assets" / "fonts"

# Google Fonts family -> (folder id, license dir in github.com/google/fonts, variants)
FAMILIES = {
    "Inter": ("inter", "ofl/inter",
              ["0,400", "1,400", "0,600", "0,700", "0,800"]),
    "Plus Jakarta Sans": ("plus-jakarta-sans", "ofl/plusjakartasans",
                          ["0,400", "1,400", "0,600", "0,700", "0,800"]),
    "Source Serif 4": ("source-serif-4", "ofl/sourceserif4",
                       ["0,400", "1,400", "0,600", "0,700"]),
    "Playfair Display": ("playfair-display", "ofl/playfairdisplay",
                         ["0,400", "1,400", "0,700", "0,800"]),
    "JetBrains Mono": ("jetbrains-mono", "ofl/jetbrainsmono", ["0,400", "0,700"]),
}

UNICODES = (
    "U+0000-036F,U+0370-03FF,U+1E00-1EFF,U+2000-206F,U+2070-209F,U+20A0-20CF,"
    "U+2100-214F,U+2150-218F,U+2190-21FF,U+2200-22FF,U+2300-23FF,U+2500-25FF,"
    "U+2605-2606,U+2610-2612,U+2713-2718,U+2794,U+279C,U+FB00-FB06,U+FEFF,U+FFFD"
)

UA = {"User-Agent": "Wget/1.21"}  # makes Google Fonts serve full static TTF files


def get(url: str) -> bytes:
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
        return r.read()


def main() -> None:
    manifest = []
    for family, (fid, lic_dir, variants) in FAMILIES.items():
        query = family.replace(" ", "+") + ":ital,wght@" + ";".join(sorted(variants))
        css = get(f"https://fonts.googleapis.com/css2?family={query}").decode()
        fam_dir = DEST / fid
        fam_dir.mkdir(parents=True, exist_ok=True)
        (fam_dir / "OFL.txt").write_bytes(
            get(f"https://raw.githubusercontent.com/google/fonts/main/{lic_dir}/OFL.txt"))
        for block in re.findall(r"@font-face\s*{(.*?)}", css, re.S):
            style = re.search(r"font-style:\s*(\w+)", block).group(1)
            weight = int(re.search(r"font-weight:\s*(\d+)", block).group(1))
            url = re.search(r"url\((.*?)\)", block).group(1)
            font = TTFont(io.BytesIO(get(url)))
            opts = subset.Options()
            opts.layout_features = ["*"]
            opts.flavor = "woff2"
            opts.name_IDs = ["*"]
            opts.notdef_outline = True
            sub = subset.Subsetter(opts)
            sub.populate(unicodes=subset.parse_unicodes(UNICODES))
            sub.subset(font)
            name = f"{weight}-{style}.woff2"
            font.flavor = "woff2"
            font.save(fam_dir / name)
            manifest.append({"family": family, "weight": weight, "style": style,
                             "file": f"{fid}/{name}"})
            print(f"{family:18} {weight} {style:6} {(fam_dir / name).stat().st_size / 1024:6.1f} KB")
    manifest.sort(key=lambda f: (f["family"], f["weight"], f["style"]))
    (DEST / "fonts.json").write_text(json.dumps(manifest, indent=1) + "\n")


if __name__ == "__main__":
    main()
