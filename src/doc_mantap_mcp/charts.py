"""Themed charts from a small YAML/JSON spec (```chart blocks), drawn with matplotlib.

matplotlib is imported lazily so documents without charts stay fast.
"""

from __future__ import annotations

import io
import math
import threading

import yaml

from .themes import Theme, mix, normalize_hex, on_color

_LOCK = threading.Lock()
_REGISTERED: set[str] = set()

_TYPES = {"bar", "barh", "line", "area", "pie", "donut", "scatter"}
_ALIASES = {"column": "bar", "columns": "bar", "bars": "bar", "horizontal_bar": "barh",
            "hbar": "barh", "doughnut": "donut", "lines": "line", "points": "scatter"}

# Thousands / decimal separators per language
_DOT_THOUSANDS = {"id", "de", "nl", "es", "it", "pt", "tr", "da", "vi", "el", "ro", "hr", "sl",
                  "sr"}
_SPACE_THOUSANDS = {"fr", "sv", "nb", "no", "fi", "pl", "cs", "sk", "ru", "uk", "hu", "et",
                    "lv", "lt", "bg"}
_COMPACT = {
    "en": ((1e12, "T"), (1e9, "B"), (1e6, "M"), (1e3, "K")),
    "id": ((1e12, " T"), (1e9, " M"), (1e6, " jt"), (1e3, " rb")),
}


def _lang(lang: str) -> str:
    return (lang or "en").split("-")[0].lower()


def format_number(value: float, decimals: int | None = None, lang: str = "en",
                  compact: bool = False) -> str:
    """Locale-aware number formatting, e.g. 12400.5 -> '12,400.5' (en) or '12.400,5' (id)."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    code = _lang(lang)
    suffix = ""
    if compact:
        for size, label in _COMPACT.get(code, _COMPACT["en"]):
            if abs(value) >= size:
                value, suffix = value / size, label
                break
    if decimals is None:
        if float(value).is_integer():
            decimals = 0
        else:
            decimals = 0 if abs(value) >= 100 else 1 if abs(value) >= 10 else 2
    text = f"{value:,.{decimals}f}"
    if suffix and "." in text:
        text = text.rstrip("0").rstrip(".")
    if code in _DOT_THOUSANDS:
        text = text.replace(",", "\0").replace(".", ",").replace("\0", ".")
    elif code in _SPACE_THOUSANDS:
        text = text.replace(",", "\u202f").replace(".", ",")
    return text + suffix


# --------------------------------------------------------------------------- #
# Spec parsing
# --------------------------------------------------------------------------- #

def _num(value, series: str):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        raise ValueError(f"Chart series '{series}': '{value}' is not a number.")
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace("_", "").replace(" ", ""))
    except ValueError:
        raise ValueError(
            f"Chart series '{series}': '{value}' is not a number. Use plain numbers "
            "such as 12400 or 12.5 (no thousands separators or units).") from None


def parse_chart_spec(text: str) -> dict:
    """Validate a ```chart block and return a normalised spec."""
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ValueError(f"Invalid ```chart block (expected YAML or JSON): {exc}") from None
    if not isinstance(data, dict):
        raise ValueError("A ```chart block must be a mapping with at least 'type' and "
                         "'series' (or 'data').")
    raw_type = str(data.get("type", "bar")).strip().lower().replace("-", "_").replace(" ", "_")
    stacked = bool(data.get("stacked")) or raw_type.startswith("stacked_")
    kind = raw_type.removeprefix("stacked_")
    kind = _ALIASES.get(kind, kind)
    if kind == "bar" and data.get("horizontal"):
        kind = "barh"
    if kind not in _TYPES:
        raise ValueError(f"Unknown chart type '{data.get('type')}'. Use bar, barh, line, "
                         "area, pie, donut or scatter (add stacked: true for stacked bars).")

    labels = data.get("labels") or data.get("categories") or data.get("x")
    series = data.get("series")
    single = data.get("data", data.get("values"))
    if series is None and single is not None:
        if isinstance(single, dict):
            labels, single = list(single.keys()), list(single.values())
        series = [{"name": data.get("name", ""), "data": single}]
    if isinstance(series, dict):
        series = [{"name": k, "data": v} for k, v in series.items()]
    if not series:
        raise ValueError("A chart needs 'series' (a list of {name, data}) or 'data'.")

    norm = []
    for idx, s in enumerate(series, 1):
        if isinstance(s, list):
            s = {"data": s}
        if not isinstance(s, dict) or not isinstance(s.get("data"), list):
            raise ValueError(f"Chart series {idx} needs a 'data' list.")
        name = str(s.get("name", f"Series {idx}") if s.get("name") is not None else "")
        if kind == "scatter":
            try:
                points = [(float(p[0]), float(p[1])) for p in s["data"]]
            except (TypeError, ValueError, IndexError):
                raise ValueError(f"Scatter series '{name}' needs data as [[x, y], ...].") from None
            values = points
        else:
            values = [_num(v, name) for v in s["data"]]
        series_type = str(s.get("type", "")).lower() or None
        if series_type and series_type not in ("bar", "line", "area"):
            raise ValueError(f"Series '{name}': type must be bar, line or area.")
        norm.append({"name": name, "data": values, "type": series_type,
                     "secondary": bool(s.get("secondary")),
                     "color": normalize_hex(s["color"]) if s.get("color") else None})

    if kind != "scatter":
        longest = max(len(s["data"]) for s in norm)
        labels = [str(x) for x in (labels or range(1, longest + 1))]
        for s in norm:
            if len(s["data"]) != len(labels):
                raise ValueError(f"Chart series '{s['name']}' has {len(s['data'])} values but "
                                 f"there are {len(labels)} labels.")
    if kind in ("pie", "donut"):
        norm = norm[:1]
        if any(v is not None and v < 0 for v in norm[0]["data"]):
            raise ValueError("Pie/donut values must be zero or positive.")

    legend = str(data.get("legend", "auto")).lower()
    if legend in ("false", "no", "off"):
        legend = "none"
    colors = data.get("colors")
    return {
        "type": kind, "stacked": stacked, "labels": labels, "series": norm,
        "title": data.get("title"), "subtitle": data.get("subtitle"),
        "caption": data.get("caption") or data.get("source"),
        "x_label": data.get("x_label"), "y_label": data.get("y_label"),
        "prefix": str(data.get("prefix") or ""), "suffix": str(data.get("suffix") or ""),
        "decimals": int(data["decimals"]) if data.get("decimals") is not None else None,
        "data_labels": bool(data.get("data_labels", data.get("show_values", False))),
        "legend": legend, "height": float(data["height"]) if data.get("height") else None,
        "colors": [normalize_hex(c) for c in colors] if isinstance(colors, list) else None,
        "y_min": data.get("y_min"), "y_max": data.get("y_max"),
        "center": data.get("center"),
    }


# --------------------------------------------------------------------------- #
# Rendering
# --------------------------------------------------------------------------- #

def _ensure_fonts(theme: Theme) -> None:
    from matplotlib import font_manager

    from .fonts import ttf_files

    for path in ttf_files({theme.font_body, theme.font_heading}):
        if str(path) not in _REGISTERED:
            font_manager.fontManager.addfont(str(path))
            _REGISTERED.add(str(path))


def render_chart(spec: dict, theme: Theme, fmt: str = "svg", width_cm: float = 16.0,
                 lang: str = "en") -> bytes:
    """Render a parsed chart spec to SVG or PNG bytes."""
    import matplotlib
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    from matplotlib.ticker import FuncFormatter, MaxNLocator

    kind = spec["type"]
    muted, text = f"#{theme.muted}", f"#{theme.text}"
    grid = f"#{mix(theme.surface, theme.muted, 0.22)}"
    colors = [f"#{c}" for c in (spec["colors"] or theme.palette)]
    pre, suf, dec = spec["prefix"], spec["suffix"], spec["decimals"]
    rc = {
        "font.family": [theme.font_body, "DejaVu Sans"], "font.size": 8.4,
        "svg.fonttype": "path", "axes.edgecolor": grid, "axes.labelcolor": muted,
        "xtick.color": muted, "ytick.color": muted, "text.color": text,
        "axes.spines.top": False, "axes.spines.right": False, "axes.linewidth": 0.8,
        "legend.fontsize": 8, "axes.labelsize": 8.2,
    }

    def label(v: float, compact_hint: bool = False) -> str:
        return pre + format_number(v, dec, lang, compact=compact_hint or abs(v) >= 1e6) + suf

    with _LOCK:
        _ensure_fonts(theme)
        with matplotlib.rc_context(rc):
            n_labels = len(spec["labels"] or [])
            default_h = {"pie": 6.4, "donut": 6.4,
                         "barh": min(14.0, max(4.6, 1.4 + 0.75 * n_labels))}.get(kind, 7.2)
            height = spec["height"] or default_h
            fig = Figure(figsize=(width_cm / 2.54, height / 2.54))
            FigureCanvasAgg(fig)
            ax = fig.add_subplot(111)
            handles = None

            if kind in ("pie", "donut"):
                values = [v or 0 for v in spec["series"][0]["data"]]
                total = sum(values) or 1
                wedge_colors = [colors[i % len(colors)] for i in range(len(values))]
                wedges, *_ = ax.pie(
                    values, colors=wedge_colors, startangle=90, counterclock=False,
                    wedgeprops={"linewidth": 2, "edgecolor": "white",
                                "width": 0.36 if kind == "donut" else 1})
                ax.set_aspect("equal")
                if kind == "donut" and spec["center"] is not None:
                    center = label(total) if str(spec["center"]).lower() == "total" \
                        else spec["center"]
                    ax.text(0, 0, str(center), ha="center", va="center", fontsize=13,
                            fontweight="bold", color=f"#{theme.primary}")
                legend_labels = [f"{lab}  ·  {format_number(v / total * 100, 0, lang)}%"
                                 for lab, v in zip(spec["labels"], values)]
                ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1.02, 0.5),
                          frameon=False, handlelength=0.9, handleheight=0.9, fontsize=8.4)
            elif kind == "scatter":
                for i, s in enumerate(spec["series"]):
                    xs, ys = zip(*s["data"]) if s["data"] else ((), ())
                    ax.scatter(xs, ys, s=30, color=f"#{s['color']}" if s["color"] else
                               colors[i % len(colors)], alpha=0.88, edgecolor="white",
                               linewidth=0.6, zorder=3, label=s["name"])
                ax.grid(True, color=grid, linewidth=0.7)
                ax.set_axisbelow(True)
                ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: label(v)))
                ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: format_number(v, None, lang)))
            else:
                handles = _draw_cartesian(ax, spec, colors, theme, label, lang)

            if kind not in ("pie", "donut"):
                ax.tick_params(length=0, pad=4)
                ax.spines["left"].set_visible(kind == "scatter")
                if kind != "barh":
                    ax.yaxis.set_major_locator(MaxNLocator(nbins=5))
                if spec["x_label"]:
                    ax.set_xlabel(str(spec["x_label"]), labelpad=6)
                if spec["y_label"]:
                    ax.set_ylabel(str(spec["y_label"]), labelpad=6)
                if spec["y_min"] is not None or spec["y_max"] is not None:
                    target = ax.set_xlim if kind == "barh" else ax.set_ylim
                    target(spec["y_min"], spec["y_max"])
                named = [s for s in spec["series"] if s["name"]]
                show = spec["legend"] != "none" and (len(named) > 1 or spec["legend"] != "auto")
                if show and named:
                    h, lab = handles or ax.get_legend_handles_labels()
                    pos = spec["legend"]
                    opts = {"frameon": False, "handlelength": 1.0, "handleheight": 0.8,
                            "columnspacing": 1.4, "borderaxespad": 0.2}
                    if pos == "right":
                        ax.legend(h, lab, loc="center left", bbox_to_anchor=(1.01, 0.5), **opts)
                    elif pos == "bottom":
                        ax.legend(h, lab, loc="upper center", bbox_to_anchor=(0.5, -0.14),
                                  ncol=min(len(lab), 5), **opts)
                    else:
                        ax.legend(h, lab, loc="lower left", bbox_to_anchor=(0, 1.01),
                                  ncol=min(len(lab), 5), **opts)

            buf = io.BytesIO()
            fig.savefig(buf, format=fmt, dpi=220, bbox_inches="tight", pad_inches=0.04,
                        transparent=fmt == "svg", facecolor="white")
            return buf.getvalue()


def _draw_cartesian(ax, spec, colors, theme: Theme, label, lang):
    """Bar, horizontal bar, line, area and bar+line combos."""
    import numpy as np
    from matplotlib.ticker import FuncFormatter

    kind, stacked = spec["type"], spec["stacked"]
    labels = spec["labels"]
    x = np.arange(len(labels))
    grid = f"#{mix(theme.surface, theme.muted, 0.22)}"
    default_type = {"barh": "bar"}.get(kind, kind)
    series = [dict(s, type=s["type"] or default_type,
                   color=f"#{s['color']}" if s["color"] else colors[i % len(colors)])
              for i, s in enumerate(spec["series"])]
    bars = [s for s in series if s["type"] == "bar"]
    others = [s for s in series if s["type"] != "bar"]
    ax2 = ax.twinx() if any(s["secondary"] for s in others) else None
    horizontal = kind == "barh"
    all_vals = [v for s in series for v in s["data"] if v is not None]
    compact = bool(all_vals) and max(abs(v) for v in all_vals) >= 1e5
    value_fmt = FuncFormatter(lambda v, _: label(v, compact))

    # Bars
    group = 0.74 if len(bars) > 1 else 0.6
    width = group if stacked else group / max(len(bars), 1)
    pos_base, neg_base = np.zeros(len(labels)), np.zeros(len(labels))
    for i, s in enumerate(bars):
        vals = np.array([np.nan if v is None else v for v in s["data"]], dtype=float)
        if stacked:
            clean = np.nan_to_num(vals)
            base = np.where(clean >= 0, pos_base, neg_base)
            offset = x
        else:
            base, offset = None, x + (i - (len(bars) - 1) / 2) * width
        draw = ax.barh if horizontal else ax.bar
        kw = {"left" if horizontal else "bottom": base} if stacked else {}
        cont = draw(offset, vals, width * (0.98 if stacked else 0.86), color=s["color"],
                    zorder=3, label=s["name"], linewidth=0, **kw)
        if stacked:
            clean = np.nan_to_num(vals)
            pos_base += np.where(clean >= 0, clean, 0)
            neg_base += np.where(clean < 0, clean, 0)
        if spec["data_labels"]:
            ax.bar_label(cont, labels=["" if np.isnan(v) else label(v, compact) for v in vals],
                         padding=0 if stacked else 2, fontsize=7.2,
                         color=f"#{on_color(s['color'][1:])}" if stacked else f"#{theme.text}",
                         label_type="center" if stacked else "edge")

    # Lines / areas
    stack_base = np.zeros(len(labels))
    for s in others:
        target = ax2 if (ax2 is not None and s["secondary"]) else ax
        vals = np.array([np.nan if v is None else v for v in s["data"]], dtype=float)
        if s["type"] == "area" and stacked:
            top = stack_base + np.nan_to_num(vals)
            target.fill_between(x, stack_base, top, color=s["color"], alpha=0.85, zorder=2,
                                linewidth=0, label=s["name"])
            stack_base = top
            continue
        target.plot(x, vals, color=s["color"], linewidth=2.2, marker="o", markersize=3.6,
                    markerfacecolor="white", markeredgewidth=1.5, zorder=4, label=s["name"])
        if s["type"] == "area":
            target.fill_between(x, vals, color=s["color"], alpha=0.12, zorder=2, linewidth=0)
        if spec["data_labels"]:
            for xi, v in zip(x, vals):
                if not np.isnan(v):
                    target.annotate(label(v, compact), (xi, v), textcoords="offset points",
                                    xytext=(0, 6), ha="center", fontsize=7.2,
                                    color=f"#{theme.text}")

    if horizontal:
        ax.set_yticks(x, labels)
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(value_fmt)
        ax.xaxis.grid(True, color=grid, linewidth=0.7)
        ax.spines["bottom"].set_visible(False)
    else:
        ax.set_xticks(x, labels)
        ax.yaxis.set_major_formatter(value_fmt)
        ax.yaxis.grid(True, color=grid, linewidth=0.7)
        if len(labels) > 8 and max(len(str(lab)) for lab in labels) > 4:
            for t in ax.get_xticklabels():
                t.set_rotation(35)
                t.set_ha("right")
    ax.set_axisbelow(True)
    if (bars or any(s["type"] == "area" for s in others)) and min(all_vals, default=0) >= 0:
        if horizontal:
            ax.set_xlim(left=0)
        else:
            ax.set_ylim(bottom=0)
    if ax2 is not None:
        ax2.tick_params(length=0, pad=4)
        for side in ("top", "left", "bottom"):
            ax2.spines[side].set_visible(False)
        ax2.spines["right"].set_visible(False)
        sec = [v for s in others if s["secondary"] for v in s["data"] if v is not None]
        sec_compact = bool(sec) and max(abs(v) for v in sec) >= 1e5
        ax2.yaxis.set_major_formatter(FuncFormatter(
            lambda v, _: format_number(v, spec["decimals"], lang, sec_compact)))
        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        return h1 + h2, l1 + l2
    return None
