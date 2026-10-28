"""Render a declarative workbook spec -> designed .xlsx with XlsxWriter.

Layout per sheet: banner (title/subtitle), KPI cards, an assumptions block, Excel tables
(with totals, conditional formats, dropdowns, sparklines), native charts and notes. Colors
and fonts come from the theme, which is also written as the workbook's Office theme so
built-in table styles and Excel's color picker match the brand.
"""

from __future__ import annotations

import datetime as dt
import math
import re
from dataclasses import dataclass, field
from html import escape
from io import StringIO
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.theme import THEME_XML_2007
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from .common import prepare_output
from .themes import Theme, mix, normalize_hex, on_color, shade, tint
from .xlsx_formula import Evaluator, FormulaError, Unsupported, compile_formula
from .xlsx_spec import Chart, Column, Input, Kpi, Sheet, Table

UNKNOWN = object()
ROW_PX = 24  # default row height (18 pt)
NUMERIC = {"number", "integer", "currency", "accounting", "percent", "multiple"}
TYPE_NAMES = NUMERIC | {"text", "date", "datetime", "time", "boolean", "url"}
_NO_DECIMALS = {"rp", "idr", "¥", "jpy", "₩", "krw", "vnd", "₫", "huf", "clp", "isk"}
_COMMA_DECIMAL = {"id", "de", "nl", "es", "it", "pt", "tr", "da", "vi", "fr", "pl", "cs", "ru",
                  "sv", "fi", "nb", "ro", "hu", "el"}
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_ISO_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?$")
_BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_STYLES = {"good": ("C6EFCE", "006100"), "bad": ("FFC7CE", "9C0006"),
           "neutral": ("FFEB9C", "9C5700")}
TEXT = {
    "en": {"assumptions": "Assumptions", "data": "data", "total": "Total"},
    "id": {"assumptions": "Asumsi", "data": "data", "total": "Total"},
}


def _texts(lang: str) -> dict:
    return TEXT.get((lang or "en").split("-")[0].lower(), TEXT["en"])


# --------------------------------------------------------------------------- #
# Values and number formats
# --------------------------------------------------------------------------- #

def number_format(kind: str | None, decimals: int | None, currency: str) -> str | None:
    if kind is None or kind in ("text", "boolean", "url"):
        return "@" if kind == "text" else None
    cur = currency.replace('"', "")
    sym = f'"{cur} "' if (len(cur) > 1 or cur.isalpha()) else f'"{cur}"'
    default = 0 if cur.lower().strip(". ") in _NO_DECIMALS else 2

    def frac(n: int) -> str:
        return "." + "0" * n if n else ""

    if kind == "integer":
        return "#,##0"
    if kind == "number":
        return "#,##0" + frac(2 if decimals is None else decimals)
    if kind == "currency":
        d = frac(default if decimals is None else decimals)
        return f'{sym}#,##0{d};[Red]-{sym}#,##0{d};{sym}"-"'
    if kind == "accounting":
        d = frac(default if decimals is None else decimals)
        return f'_({sym}* #,##0{d}_);_({sym}* (#,##0{d});_({sym}* "-"_);_(@_)'
    if kind == "percent":
        return "0" + frac(1 if decimals is None else decimals) + "%"
    if kind == "multiple":
        return "0" + frac(1 if decimals is None else decimals) + '"x"'
    if kind == "date":
        return "dd mmm yyyy"
    if kind == "datetime":
        return "dd mmm yyyy hh:mm"
    if kind == "time":
        return "hh:mm"
    return None


def resolve_format(fmt: str | None, currency: str, decimals: int | None = None):
    """'currency' -> (kind, excel format); '0.0"x"' -> (None, '0.0"x"')."""
    if not fmt:
        return None, None
    key = fmt.strip().lower()
    if key in TYPE_NAMES:
        return key, number_format(key, decimals, currency)
    return None, fmt


def parse_number(text: str, lang: str) -> float | None:
    s = str(text).strip()
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")") or s.lstrip().startswith(("-", "−"))
    percent = s.endswith("%")
    digits = re.sub(r"[^\d.,]", "", s)
    if not digits or not re.search(r"\d", digits):
        return None
    if "." in digits and "," in digits:
        dec = "." if digits.rfind(".") > digits.rfind(",") else ","
    elif "," in digits:
        dec = "," if (lang in _COMMA_DECIMAL and not re.fullmatch(r"\d{1,3}(,\d{3})+", digits)) \
            else ("," if not re.fullmatch(r"\d{1,3}(,\d{3})+", digits) else None)
    elif "." in digits:
        dec = None if re.fullmatch(r"\d{1,3}(\.\d{3})+", digits) and lang in _COMMA_DECIMAL \
            else "."
    else:
        dec = None
    thousands = {".": ",", ",": "."}.get(dec, ".,")
    for ch in thousands:
        digits = digits.replace(ch, "")
    if dec:
        digits = digits.replace(dec, ".")
    try:
        value = float(digits)
    except ValueError:
        return None
    value = -value if negative else value
    return value / 100 if percent else value


def to_datetime(value) -> dt.datetime | dt.date | None:
    if isinstance(value, (dt.date, dt.datetime)):
        return value
    s = str(value).strip()
    try:
        if _ISO_DATE.match(s):
            return dt.date.fromisoformat(s)
        if _ISO_DATETIME.match(s):
            return dt.datetime.fromisoformat(s.replace(" ", "T"))
    except ValueError:
        return None
    return None


def detect_type(values: list) -> str:
    vals = [v for v in values if v is not None and v != ""
            and not (isinstance(v, str) and v.startswith("="))]
    if not vals:
        return "text"
    if all(isinstance(v, bool) for v in vals):
        return "boolean"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
        return "integer" if all(float(v).is_integer() for v in vals) else "number"
    if all(isinstance(v, str) and _ISO_DATE.match(v.strip()) for v in vals):
        return "date"
    if all(isinstance(v, str) and _ISO_DATETIME.match(v.strip()) for v in vals):
        return "datetime"
    if all(isinstance(v, str) and v.strip().lower().startswith(("http://", "https://"))
           for v in vals):
        return "url"
    return "text"


def _display_len(value, kind: str | None, currency: str) -> int:
    if value is None or value is UNKNOWN:
        return 0
    if isinstance(value, (dt.date, dt.datetime)):
        return 11 if kind != "datetime" else 17
    if isinstance(value, bool):
        return 5
    if isinstance(value, (int, float)):
        if kind == "percent":
            return len(f"{value * 100:,.1f}%")
        text = f"{value:,.0f}" if abs(value) >= 1000 or float(value).is_integer() \
            else f"{value:,.2f}"
        extra = len(currency) + 1 if kind in ("currency", "accounting") else 0
        return len(text) + extra + (2 if kind == "accounting" else 0)
    return max((len(line) for line in str(value).split("\n")), default=0)


def _name_ok(name: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.]*", name)) and not re.fullmatch(
        r"[A-Za-z]{1,3}\d+|[RrCc]\d*|[Rr]\d+[Cc]\d+", name) and len(name) <= 250


def make_name(text: str, used: set[str], prefix: str) -> str:
    words = re.findall(r"[A-Za-z0-9]+", text or "")
    base = "".join(w[:1].upper() + w[1:] for w in words) or prefix
    if not _name_ok(base):
        base = f"{prefix}{base}"
    if not _name_ok(base):
        base = prefix
    name, n = base[:60], 2
    while name.lower() in used:
        name, n = f"{base[:56]}{n}", n + 1
    used.add(name.lower())
    return name


# --------------------------------------------------------------------------- #
# Model (resolved spec) and formula support
# --------------------------------------------------------------------------- #

@dataclass
class TableModel:
    spec: Table
    sheet: str
    name: str
    columns: list[str]
    col_specs: list[Column]
    kinds: list[str]
    formats: list[str | None]
    rows: list[list[Any]]
    col_formulas: dict[int, Any] = field(default_factory=dict)
    cell_formulas: dict[tuple[int, int], Any] = field(default_factory=dict)
    totals: dict[int, tuple[str, str]] = field(default_factory=dict)  # col -> (kind, text)
    header_row: int = 0
    first_col: int = 1
    _memo: dict = field(default_factory=dict)
    _busy: set = field(default_factory=set)

    @property
    def key(self) -> str:
        return self.name.lower()

    @property
    def n_rows(self) -> int:
        return max(len(self.rows), 1)

    @property
    def has_total(self) -> bool:
        if self.spec.total_row is not None:
            return self.spec.total_row
        return bool(self.totals)

    def index(self, col: str) -> int | None:
        target = col.strip().lower()
        for i, c in enumerate(self.columns):
            if c.lower() == target:
                return i
        return None

    def value(self, r: int, c: int, ev: Evaluator):
        if (r, c) in self._memo:
            v = self._memo[(r, c)]
        else:
            node = self.col_formulas.get(c) or self.cell_formulas.get((r, c))
            if node is None:
                v = self.rows[r][c] if r < len(self.rows) else None
            else:
                if (r, c) in self._busy:
                    raise Unsupported("circular reference")
                self._busy.add((r, c))
                try:
                    v = ev.value(node, self.key, r)
                finally:
                    self._busy.discard((r, c))
                v = UNKNOWN if v is None else v
            self._memo[(r, c)] = v
        if v is UNKNOWN:
            raise Unsupported("value unknown")
        if isinstance(v, (dt.date, dt.datetime)):
            raise Unsupported("dates are not evaluated")
        return v

    def cached(self, r: int, c: int, ev: Evaluator):
        try:
            return self.value(r, c, ev)
        except Unsupported:
            return None

    def column(self, c: int, ev: Evaluator) -> list:
        return [self.value(r, c, ev) for r in range(len(self.rows))]


def _theme_xml(theme: Theme) -> str:
    accents = [theme.primary, theme.accent] + [c for c in theme.palette
                                                if c not in (theme.primary, theme.accent)]
    accents = (accents + list(theme.palette) * 2)[:6]
    colors = {"dk2": theme.text, "lt2": theme.surface, "hlink": shade(theme.accent, 0.1),
              "folHlink": theme.secondary}
    colors.update({f"accent{i + 1}": c for i, c in enumerate(accents)})
    xml = THEME_XML_2007
    for tag, color in colors.items():
        xml = re.sub(rf'(<a:{tag}><a:srgbClr val=")[0-9A-Fa-f]{{6}}("/>)', rf"\g<1>{color}\g<2>",
                     xml)
    xml = xml.replace('<a:latin typeface="Cambria"/>',
                      f'<a:latin typeface="{escape(theme.docx_heading)}"/>', 1)
    xml = xml.replace('<a:latin typeface="Calibri"/>',
                      f'<a:latin typeface="{escape(theme.docx_body)}"/>', 1)
    return xml.replace('name="Office Theme"', f'name="doc-mantap {theme.name}"', 1)


@dataclass
class XlsxResult:
    path: Path
    summary: list[str]
    warnings: list[str]


class WorkbookBuilder:
    def __init__(self, path: Path, theme: Theme, currency: str, lang: str):
        self.t, self.currency, self.lang = theme, currency, (lang or "en").split("-")[0].lower()
        self.texts = _texts(self.lang)
        self.wb = xlsxwriter.Workbook(str(path), {
            "use_future_functions": True,
            "default_format_properties": {"font_name": theme.docx_body, "font_size": 10,
                                          "font_scheme": "minor"},
        })
        self.wb.use_custom_theme(StringIO(_theme_xml(theme)))
        self._formats: dict = {}
        self.tables: dict[str, TableModel] = {}
        self.titles: dict[str, TableModel] = {}
        self.names: dict[str, Any] = {}
        self.warnings: list[str] = []
        self.summary: list[str] = []
        self.sheet_slots: dict[str, tuple] = {}
        self.ev = Evaluator(self.tables, self.names)

    # ---- formats -----------------------------------------------------------
    def fmt(self, **props):
        base = {"font_name": self.t.docx_body, "font_size": 10, "font_color": f"#{self.t.text}",
                "valign": "vcenter"}
        base.update({k: v for k, v in props.items() if v is not None})
        return self._cached(base)

    def hfmt(self, **props):
        """Table header format without colors, so the table style supplies them."""
        base = {"font_name": self.t.docx_body, "font_size": 10, "bold": True,
                "valign": "vcenter", "text_wrap": True}
        base.update({k: v for k, v in props.items() if v is not None})
        return self._cached(base)

    def _cached(self, props: dict):
        key = tuple(sorted(props.items()))
        if key not in self._formats:
            f = self.wb.add_format(props)
            f._dm_props = dict(props)
            self._formats[key] = f
        return self._formats[key]

    # ---- model building ------------------------------------------------------
    def build_models(self, sheets: list[Sheet]) -> list[tuple[str, Sheet]]:
        used_sheets: set[str] = set()
        used_names: set[str] = set()
        out = []
        for idx, sheet in enumerate(sheets):
            name = _BAD_SHEET_CHARS.sub("-", sheet.name).strip("'").strip()[:31] or f"Sheet{idx + 1}"
            base, n = name, 2
            while name.lower() in used_sheets:
                name = f"{base[:28]} {n}"
                n += 1
            if name != sheet.name:
                self.warnings.append(f"Sheet '{sheet.name}' was renamed to '{name}' "
                                     "(Excel limits: 31 characters, no []:*?/\\, unique).")
            used_sheets.add(name.lower())
            out.append((name, sheet))

        # Inline chart data becomes a small table on the same sheet.
        for name, sheet in out:
            for n_chart, chart in enumerate(sheet.charts, 1):
                if chart.table or not chart.series:
                    continue
                labels = chart.labels or [str(i + 1) for i in range(len(chart.series[0].data))]
                for s in chart.series:
                    if len(s.data) != len(labels):
                        raise ValueError(f"Chart '{chart.title or n_chart}' on sheet '{name}': "
                                         f"series '{s.name}' has {len(s.data)} values but "
                                         f"there are {len(labels)} labels.")
                title = f"{chart.title or 'Chart'} ({self.texts['data']})"
                data = Table(title=title, columns=["Label"] + [s.name for s in chart.series],
                             rows=[[lab] + [s.data[i] for s in chart.series]
                                   for i, lab in enumerate(labels)])
                sheet.tables.append(data)
                chart.table = title

        # Assumptions -> defined names (needed before evaluating table formulas).
        for name, sheet in out:
            for inp in sheet.inputs:
                if inp.name and not _name_ok(inp.name):
                    raise ValueError(f"Input name '{inp.name}' is not a valid Excel name. Use "
                                     "letters, digits and _ (start with a letter, e.g. TaxRate).")
                key = inp.name or make_name(inp.label, used_names, "Input")
                if inp.name:
                    if inp.name.lower() in used_names:
                        raise ValueError(f"Input name '{inp.name}' is used twice.")
                    used_names.add(inp.name.lower())
                inp.name = key
                kind, _ = resolve_format(inp.format, self.currency)
                value = self._coerce(inp.value, kind or detect_type([inp.value]),
                                     f"input '{inp.label}'")
                inp.value = value
                if not (isinstance(value, str) and value.startswith("=")):
                    self.names[key] = value

        for name, sheet in out:
            for table in sheet.tables:
                model = self._table_model(name, table, used_names)
                self.tables[model.key] = model
                for alias in filter(None, (table.title, table.name)):
                    self.titles.setdefault(alias.strip().lower(), model)
        self.ev.names.update({k.lower(): v for k, v in self.names.items()})

        # Validate every formula now, so mistakes come back as clear errors.
        for model in self.tables.values():
            for c, node in model.col_formulas.items():
                self._check(node, model, f"column '{model.columns[c]}'")
            for (r, c), node in model.cell_formulas.items():
                self._check(node, model, f"row {r + 1}, column '{model.columns[c]}'")
        for name, sheet in out:
            for kpi in sheet.kpis:
                for part in (kpi.value, kpi.delta):
                    if isinstance(part, str) and part.startswith("="):
                        self._check(self._compile(part, f"KPI '{kpi.label}'"), None,
                                    f"KPI '{kpi.label}'")
        return out

    def _compile(self, formula: str, where: str):
        try:
            return compile_formula(formula)
        except FormulaError as exc:
            raise ValueError(f"Formula error in {where}: {exc} Formula: {formula}") from None
        except Unsupported:
            return None

    def _check(self, node, model: TableModel | None, where: str) -> None:
        if node is None:
            return
        try:
            self.ev.check(node, model.key if model else None)
        except FormulaError as exc:
            prefix = f"table '{model.name}', " if model else ""
            raise ValueError(f"Formula error in {prefix}{where}: {exc}") from None

    def _coerce(self, value, kind: str, where: str):
        if value is None or (isinstance(value, str) and value.startswith("=")):
            return value
        if kind in NUMERIC:
            if isinstance(value, bool):
                return value
            if isinstance(value, (int, float)):
                return value
            num = parse_number(value, self.lang)
            if num is None:
                self.warnings.append(f"{where}: '{value}' is not a number; written as text.")
                return str(value)
            if kind != "percent" and str(value).strip().endswith("%"):
                return num
            return num
        if kind in ("date", "datetime", "time"):
            d = to_datetime(value)
            if d is None and value != "":
                self.warnings.append(f"{where}: '{value}' is not an ISO date (YYYY-MM-DD); "
                                     "written as text.")
                return str(value)
            return d
        return value

    def _table_model(self, sheet: str, table: Table, used: set[str]) -> TableModel:
        specs = [c if isinstance(c, Column) else Column(header=str(c)) for c in table.columns]
        headers = [c.header.strip() or f"Column{i + 1}" for i, c in enumerate(specs)]
        seen = set()
        for h in headers:
            if h.lower() in seen:
                raise ValueError(f"Table '{table.title or table.name}' on sheet '{sheet}' has "
                                 f"duplicate column '{h}'. Headers must be unique.")
            seen.add(h.lower())
        if table.name and not _name_ok(table.name):
            raise ValueError(f"Table name '{table.name}' is invalid. Use letters, digits and _ "
                             "(start with a letter, no spaces), e.g. 'Sales2026'.")
        if table.name:
            if table.name.lower() in used:
                raise ValueError(f"Name '{table.name}' is used twice (tables and inputs share "
                                 "one namespace).")
            used.add(table.name.lower())
            name = table.name
        else:
            name = make_name(table.title or "", used, "Table")
        ncols = len(specs)
        rows = []
        for r_idx, row in enumerate(table.rows):
            if len(row) > ncols:
                raise ValueError(f"Table '{name}' row {r_idx + 1} has {len(row)} values but "
                                 f"only {ncols} columns.")
            rows.append(list(row) + [None] * (ncols - len(row)))
        kinds, formats = [], []
        model = TableModel(spec=table, sheet=sheet, name=name, columns=headers, col_specs=specs,
                           kinds=kinds, formats=formats, rows=rows)
        for c, spec in enumerate(specs):
            raw = [row[c] for row in rows]
            kind, custom = resolve_format(spec.format, self.currency, spec.decimals)
            kind = spec.type or kind or ("text" if spec.sparkline else
                                         "number" if spec.formula else detect_type(raw))
            if spec.formula:
                model.col_formulas[c] = self._compile(
                    spec.formula if spec.formula.startswith("=") else "=" + spec.formula,
                    f"table '{name}', column '{headers[c]}'")
            where = f"Table '{name}', column '{headers[c]}'"
            for r in range(len(rows)):
                rows[r][c] = None if (spec.formula or spec.sparkline) else \
                    self._coerce(rows[r][c], kind, where)
                v = rows[r][c]
                if isinstance(v, str) and v.startswith("="):
                    model.cell_formulas[(r, c)] = self._compile(v, f"{where}, row {r + 1}")
            if kind == "percent" and not spec.formula:
                nums = [v for v in (row[c] for row in rows)
                        if isinstance(v, (int, float)) and not isinstance(v, bool)]
                if nums and all(1 < abs(v) <= 100 for v in nums):
                    for row in rows:
                        if isinstance(row[c], (int, float)) and not isinstance(row[c], bool):
                            row[c] = row[c] / 100
                    self.warnings.append(f"{where}: values looked like whole percentages "
                                         "(e.g. 15 for 15%) and were divided by 100. Pass "
                                         "fractions (0.15) to avoid this.")
            kinds.append(kind)
            formats.append(custom or number_format(kind, spec.decimals, self.currency))
            if spec.total:
                total = spec.total.strip()
                low = total.lower()
                if low in ("sum", "average", "count", "min", "max", "count_nums", "std_dev",
                           "var"):
                    model.totals[c] = ("function", low)
                elif total.startswith("="):
                    model.totals[c] = ("formula", total)
                else:
                    model.totals[c] = ("label", total)
        if not rows and table.blank_rows == 0:
            rows.append([None] * ncols)
        for _ in range(table.blank_rows):
            rows.append([None] * ncols)
        return model

    # ---- sheet layout ------------------------------------------------------------
    def column_widths(self, sheet: Sheet, models: list[TableModel], span: int) -> list[float]:
        widths = [0.0] * span
        for m in models:
            for c, spec in enumerate(m.col_specs):
                if spec.width:
                    widths[c] = max(widths[c], spec.width)
                    continue
                longest = max([_display_len(m.cached(r, c, self.ev), m.kinds[c], self.currency)
                               for r in range(len(m.rows))] or [0])
                header = len(m.columns[c]) + 4  # room for the filter button
                need = max(header, longest + 3)
                if m.kinds[c] in NUMERIC and longest == 0:
                    need = max(need, 12)
                if spec.sparkline:
                    need = max(need, 16)
                widths[c] = max(widths[c], min(need, 40 if spec.wrap else 55))
        if sheet.inputs:
            widths[0] = max(widths[0], min(max(len(i.label) for i in sheet.inputs) + 3, 45))
            longest = 0
            for inp in sheet.inputs:
                kind, _ = resolve_format(inp.format, self.currency)
                longest = max(longest, _display_len(inp.value, kind or detect_type([inp.value]),
                                                    self.currency))
            widths[1] = max(widths[1], min(longest + 4, 40), 14)
        return [w if w else 13.0 for w in widths]

    def write_sheet(self, name: str, sheet: Sheet, sheet_index: int) -> None:
        t = self.t
        ws = self.wb.add_worksheet(name)
        models = [self.tables[self._model_key(tb)] for tb in sheet.tables]
        span = max([6] + [len(m.columns) for m in models] + [3 if sheet.inputs else 0])
        widths = self.column_widths(sheet, models, span)
        ws.set_default_row(18)
        ws.set_column(0, 0, 2.6)
        for c, w in enumerate(widths, start=1):
            ws.set_column(c, c, w)
        ws.set_column(span + 1, span + 1, 2.6)
        if not sheet.gridlines:
            ws.hide_gridlines(2)
        ws.set_zoom(sheet.zoom or 100)
        ws.set_tab_color("#" + normalize_hex(sheet.tab_color) if sheet.tab_color
                         else "#" + t.palette[sheet_index % len(t.palette)])
        first, last = 1, span
        row = 0
        ws.set_row(row, 8)
        row += 1

        if sheet.title:
            on = on_color(t.primary)
            ws.set_row(row, 36)
            ws.merge_range(row, first, row, last, sheet.title, self.fmt(
                bold=True, font_size=18, font_name=t.docx_heading, font_color=f"#{on}",
                bg_color=f"#{t.primary}", indent=1))
            row += 1
            if sheet.subtitle:
                ws.set_row(row, 20)
                ws.merge_range(row, first, row, last, sheet.subtitle, self.fmt(
                    font_size=10, font_color=f"#{mix(on, t.primary, 0.25)}",
                    bg_color=f"#{t.primary}", indent=1, italic=True))
                row += 1
            ws.set_row(row, 4)
            ws.merge_range(row, first, row, last, "", self.fmt(bg_color=f"#{t.accent}"))
            row += 1
            ws.set_row(row, 12)
            row += 1
        elif sheet.subtitle:
            ws.merge_range(row, first, row, last, sheet.subtitle,
                           self.fmt(italic=True, font_color=f"#{t.muted}"))
            row += 2

        if sheet.kpis:
            row = self.write_kpis(ws, sheet.kpis, row, widths)
        if sheet.inputs:
            row = self.write_inputs(ws, name, sheet, row)
        right_charts = [c for c in sheet.charts if c.position == "right"]
        first_table_row = None
        for m in models:
            row, header_row = self.write_table(ws, m, row)
            if first_table_row is None:
                first_table_row = header_row
        content_px = sum(int(w * 7 + 5) for w in widths)
        below = [c for c in sheet.charts if c.position != "right"]
        self._chart_slots = []
        if below:
            per_row = 2 if len(below) > 1 and content_px >= 2 * 380 + 16 else 1
            default_w = max(360, (content_px - 16) // per_row) if per_row == 2 \
                else min(max(content_px, 480), 760)
            row += 1
            col_px = [int(w * 7 + 5) for w in widths]
            for i in range(0, len(below), per_row):
                group = below[i:i + per_row]
                tallest = 0
                x_px = 0
                for chart in group:
                    w = chart.width or default_w
                    h = chart.height or 300
                    col, offset, acc = first, x_px, 0
                    for k, cw in enumerate(col_px):
                        if acc + cw > x_px:
                            col, offset = first + k, x_px - acc
                            break
                        acc += cw
                    self._chart_slots.append((chart, row, col, offset, w, h))
                    x_px += w + 16
                    tallest = max(tallest, h)
                row += math.ceil((tallest + 12) / ROW_PX) + 1
        if right_charts:
            r = first_table_row if first_table_row is not None else row
            for chart in right_charts:
                w, h = chart.width or 520, chart.height or 300
                self._chart_slots.append((chart, max(r - 1, 1), span + 2, 0, w, h))
                r += math.ceil((h + 12) / ROW_PX) + 1
        for note in sheet.notes:
            ws.write_string(row, first, note, self.fmt(italic=True, font_size=8.5,
                                                       font_color=f"#{t.muted}"))
            row += 1

        # Freeze the header of a table near the top.
        freeze = sheet.freeze
        if freeze is None:
            freeze = first_table_row is not None and first_table_row <= 7 and not right_charts
        if freeze and first_table_row is not None:
            ws.freeze_panes(first_table_row + 1, 0)

        # Print setup
        total_chars = sum(widths)
        landscape = sheet.landscape if sheet.landscape is not None else \
            (total_chars > 100 or bool(right_charts))
        if landscape:
            ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
        ws.center_horizontally()
        if first_table_row is not None:
            ws.repeat_rows(first_table_row)
        footer_title = (sheet.title or name).replace("&", "&&")
        ws.set_footer(f"&L&8{footer_title}&R&8&P / &N")
        self.sheet_slots[name] = (ws, self._chart_slots)

    def _model_key(self, table: Table) -> str:
        for key, model in self.tables.items():
            if model.spec is table:
                return key
        raise KeyError(table.title)

    def write_kpis(self, ws, kpis: list[Kpi], row: int, widths: list[float]) -> int:
        t = self.t
        rows = [kpis[i:i + 4] for i in range(0, len(kpis), 4)]
        for group in rows:
            n = len(group)
            spans = _partition(widths, n)
            ws.set_row(row, 22)
            ws.set_row(row + 1, 34)
            ws.set_row(row + 2, 20)
            gap = {"left": 2, "left_color": "#FFFFFF", "right": 2, "right_color": "#FFFFFF"}
            for kpi, (c0, c1) in zip(group, spans):
                c0, c1 = c0 + 1, c1 + 1
                fill = f"#{t.surface}"
                label_f = self.fmt(bold=True, font_size=8, font_color=f"#{t.muted}", bg_color=fill,
                                   top=5, top_color=f"#{t.accent}", indent=1, **gap)
                self._merge(ws, row, c0, c1, kpi.label.upper(), label_f)
                kind, custom = resolve_format(kpi.format, self.currency)
                value = kpi.value
                if isinstance(value, str) and value.startswith("="):
                    node = self._compile(value, f"KPI '{kpi.label}'")
                    sample = self.ev.value(node) if node is not None else None
                else:
                    sample = self._coerce(value, kind or detect_type([value]),
                                          f"KPI '{kpi.label}'")
                value_f = self.fmt(bold=True, font_size=20, font_name=t.docx_heading,
                                   font_color=f"#{t.primary}", bg_color=fill, indent=1,
                                   num_format=self._kpi_format(kind, custom, sample),
                                   align="left", shrink=True, **gap)
                self._merge_value(ws, row + 1, c0, c1, value, kind, value_f,
                                  f"KPI '{kpi.label}'")
                self._kpi_delta(ws, row + 2, c0, c1, kpi, fill, gap)
            row += 4
        return row

    def _merge(self, ws, r, c0, c1, value, fmt) -> None:
        if c1 > c0:
            ws.merge_range(r, c0, r, c1, value, fmt)
        else:
            ws.write(r, c0, value, fmt)

    def _merge_value(self, ws, r, c0, c1, value, kind, fmt, where: str) -> None:
        if isinstance(value, str) and value.startswith("="):
            node = self._compile(value, where)
            cached = self.ev.value(node) if node is not None else None
            if not _props(fmt).get("num_format"):
                fmt = self.fmt(**{**_props(fmt), "num_format": _auto_format(cached, kind,
                                                                            self.currency)})
            if c1 > c0:
                ws.merge_range(r, c0, r, c1, "", fmt)
            ws.write_formula(r, c0, value, fmt, cached if cached is not None else 0)
            return
        value = self._coerce(value, kind or detect_type([value]), where)
        if isinstance(value, (dt.date, dt.datetime)):
            fmt = self.fmt(**{**_props(fmt), "num_format": "dd mmm yyyy"})
        elif not _props(fmt).get("num_format"):
            fmt = self.fmt(**{**_props(fmt), "num_format": _auto_format(value, kind,
                                                                        self.currency)})
        if c1 > c0:
            ws.merge_range(r, c0, r, c1, "", fmt)
        if isinstance(value, (dt.date, dt.datetime)):
            ws.write_datetime(r, c0, value, fmt)
        elif value is None:
            ws.write_blank(r, c0, None, fmt)
        elif isinstance(value, str):
            ws.write_string(r, c0, value, fmt)
        else:
            ws.write(r, c0, value, fmt)

    def _kpi_delta(self, ws, r, c0, c1, kpi: Kpi, fill: str, gap: dict) -> None:
        t = self.t
        up, down = ("#047857", "#B91C1C") if not kpi.lower_is_better else ("#B91C1C", "#047857")
        muted = f"#{t.muted}"
        base = {"font_size": 9, "bg_color": fill, "indent": 1, "align": "left", **gap}
        note = f"  {kpi.note}" if kpi.note else ""
        d = kpi.delta
        if d is None:
            self._merge(ws, r, c0, c1, kpi.note or "", self.fmt(**base, font_color=muted))
            return
        kind, num = resolve_format(kpi.delta_format or "percent", self.currency)
        if isinstance(d, (int, float)) and not isinstance(d, bool):
            text = self._delta_text(abs(d), kind)
            if text is not None:
                arrow = "▲ " if d > 0 else "▼ " if d < 0 else "– "
                color = up if d > 0 else down if d < 0 else muted
                self._merge(ws, r, c0, c1, arrow + text + note,
                            self.fmt(**base, bold=True, font_color=color))
                return
        if isinstance(d, (int, float)) and not isinstance(d, bool) or \
                (isinstance(d, str) and d.startswith("=")):
            num = num or "0.0%"
            lit = f' "{kpi.note.replace(chr(34), "")}"' if kpi.note else ""
            fmt = self.fmt(**base, bold=True, font_color=muted,
                           num_format=f'"▲ "{num}{lit};"▼ "{num}{lit};"– "{num}{lit}')
            self._merge_value(ws, r, c0, c1, d, kind, fmt, f"KPI '{kpi.label}' delta")
            for criteria, color in ((">", up), ("<", down)):
                ws.conditional_format(r, c0, r, c1, {
                    "type": "cell", "criteria": criteria, "value": 0,
                    "format": self.wb.add_format({"font_color": color})})
            return
        first = str(d).strip()[:1]
        color = up if first in "+▲↑" else down if first in "-−▼↓" else muted
        self._merge(ws, r, c0, c1, str(d) + note, self.fmt(**base, bold=True, font_color=color))

    def _delta_text(self, value: float, kind: str | None) -> str | None:
        from .charts import format_number
        if kind in (None, "percent"):
            return format_number(value * 100, 1, self.lang) + "%"
        if kind == "integer":
            return format_number(value, 0, self.lang)
        if kind == "number":
            return format_number(value, 2, self.lang)
        if kind in ("currency", "accounting"):
            decimals = 0 if self.currency.lower().strip(". ") in _NO_DECIMALS else 2
            return f"{self.currency} " + format_number(value, decimals, self.lang)
        if kind == "multiple":
            return format_number(value, 1, self.lang) + "x"
        return None

    def _kpi_format(self, kind: str | None, custom: str | None, value) -> str | None:
        """Number format for a KPI value; large money/numbers are shown compactly."""
        if custom and kind is None:
            return custom
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return custom
        if kind in (None, "number", "integer", "currency", "accounting") and abs(value) >= 1e6:
            scale, suffix = ((",,,", " M" if self.lang == "id" else "B") if abs(value) >= 1e9
                             else (",,", " jt" if self.lang == "id" else "M"))
            if kind in ("currency", "accounting"):
                cur = self.currency.replace('"', "")
                sym = f'"{cur} "' if (len(cur) > 1 or cur.isalpha()) else f'"{cur}"'
                return f'{sym}#,##0.0{scale}"{suffix}"'
            return f'#,##0.0{scale}"{suffix}"'
        return custom or _auto_format(value, kind, self.currency)

    def write_inputs(self, ws, sheet_name: str, sheet: Sheet, row: int) -> int:
        t = self.t
        title = sheet.inputs_title or self.texts["assumptions"]
        ws.set_row(row, 24)
        ws.write_string(row, 1, title, self.fmt(bold=True, font_size=12,
                                                font_name=t.docx_heading,
                                                font_color=f"#{t.primary}"))
        row += 1
        border = f"#{t.border}"
        for inp in sheet.inputs:
            kind, num_fmt = resolve_format(inp.format, self.currency)
            value = inp.value
            kind = kind or detect_type([value])
            num_fmt = num_fmt or _auto_format(value, kind, self.currency)
            ws.write_string(row, 1, inp.label, self.fmt(bottom=1, bottom_color=border))
            vf = self.fmt(font_color="#0000FF", bg_color="#FFF9DB", border=1, border_color=border,
                          num_format=num_fmt, align="right")
            if isinstance(value, (dt.date, dt.datetime)):
                ws.write_datetime(row, 2, value, vf)
            elif isinstance(value, str) and value.startswith("="):
                node = self._compile(value, f"input '{inp.label}'")
                cached = self.ev.value(node) if node is not None else None
                ws.write_formula(row, 2, value, self.fmt(**{**_props(vf), "font_color": "#000000"}),
                                 cached if cached is not None else 0)
            elif value is None:
                ws.write_blank(row, 2, None, vf)
            else:
                ws.write(row, 2, value, vf)
            self.wb.define_name(inp.name, f"='{sheet_name}'!{xl_rowcol_to_cell(row, 2, True, True)}")
            if inp.note:
                ws.write_string(row, 3, inp.note, self.fmt(italic=True, font_size=8.5,
                                                           font_color=f"#{t.muted}", indent=1))
            row += 1
        self.summary.append(f"  inputs: {', '.join(i.name for i in sheet.inputs)} "
                            f"(defined names, cells C{row - len(sheet.inputs) + 1}:C{row})")
        return row + 1

    def write_table(self, ws, m: TableModel, row: int) -> tuple[int, int]:
        t, spec = self.t, m.spec
        width = len(m.columns)
        first, last = 1, width
        if spec.title:
            ws.set_row(row, 24)
            ws.write_string(row, first, spec.title, self.fmt(
                bold=True, font_size=12, font_name=t.docx_heading, font_color=f"#{t.primary}",
                bottom=1, bottom_color=f"#{t.border}"))
            for c in range(first + 1, last + 1):
                ws.write_blank(row, c, None, self.fmt(bottom=1, bottom_color=f"#{t.border}"))
            row += 1
        if spec.description:
            ws.write_string(row, first, spec.description,
                            self.fmt(font_size=9, font_color=f"#{t.muted}"))
            row += 1
        if spec.title or spec.description:
            ws.set_row(row, 6)
            row += 1
        header = row
        m.header_row, m.first_col = header, first
        n = m.n_rows
        last_row = header + n + (1 if m.has_total else 0)
        columns = []
        for c, cs in enumerate(m.col_specs):
            props = {"num_format": m.formats[c]}
            if cs.align or m.kinds[c] in ("date", "datetime", "time", "boolean"):
                props["align"] = cs.align or "center"
            if cs.wrap:
                props["text_wrap"] = True
                props["valign"] = "top"
            if cs.input:
                props.update(font_color="#0000FF", bg_color="#FFF9DB")
            col = {"header": m.columns[c], "format": self.fmt(**props),
                   "header_format": self.hfmt(align="right" if m.kinds[c] in NUMERIC
                                              else None)}
            if cs.formula:
                col["formula"] = _this_row(cs.formula)
            total = m.totals.get(c)
            if total and total[0] == "function":
                col["total_function"] = total[1]
                col["total_value"] = self._total_value(m, c, total[1]) or 0
            elif total and total[0] == "label":
                col["total_string"] = total[1]
            elif total and total[0] == "formula":
                col["total_function"] = total[1]
                node = self._compile(total[1], f"total of '{m.columns[c]}'")
                if node is not None:
                    self._check(node, m, f"total of '{m.columns[c]}'")
                cached = self.ev.value(node, m.key) if node is not None else None
                col["total_value"] = cached if cached is not None else 0
            elif c == 0 and m.has_total and not m.totals.get(0):
                col["total_string"] = self.texts["total"]
            columns.append(col)
        style = spec.style or ("Table Style Light 2" if t.table_style == "lined"
                               else "Table Style Medium 2")
        ws.add_table(header, first, last_row, last, {
            "name": m.name, "columns": columns, "style": style, "total_row": m.has_total,
            "autofilter": True, "banded_rows": True,
        })
        # Values with the right writer (dates, urls, cached formula results).
        for r in range(len(m.rows)):
            xr = header + 1 + r
            for c, cs in enumerate(m.col_specs):
                xc = first + c
                fmt = columns[c]["format"]
                if cs.formula:
                    cached = m.cached(r, c, self.ev)
                    ws.write_formula(xr, xc, _this_row(cs.formula), fmt,
                                     cached if cached is not None else 0)
                    continue
                if cs.sparkline:
                    continue
                v = m.rows[r][c]
                if v is None:
                    ws.write_blank(xr, xc, None, fmt)
                elif (r, c) in m.cell_formulas:
                    cached = m.cached(r, c, self.ev)
                    ws.write_formula(xr, xc, _this_row(v), fmt,
                                     cached if cached is not None else 0)
                elif isinstance(v, (dt.date, dt.datetime)):
                    ws.write_datetime(xr, xc, v, fmt)
                elif m.kinds[c] == "url" and isinstance(v, str):
                    ws.write_url(xr, xc, v, self.fmt(**{**_props(fmt), "font_color":
                                                       f"#{shade(t.accent, 0.1)}",
                                                       "underline": 1}), string=v)
                elif isinstance(v, str):
                    ws.write_string(xr, xc, v, fmt)
                else:
                    ws.write(xr, xc, v, fmt)
        data_first, data_last = header + 1, header + n
        for c, cs in enumerate(m.col_specs):
            xc = first + c
            if cs.conditional:
                self._conditional(ws, data_first, xc, data_last, cs.conditional)
            if cs.options:
                source = [str(o) for o in cs.options]
                if any("," in s for s in source):
                    raise ValueError(f"Dropdown options for '{m.columns[c]}' cannot contain "
                                     "commas (Excel uses them as separators).")
                if sum(len(s) + 1 for s in source) > 255:
                    raise ValueError(f"Dropdown options for '{m.columns[c]}' exceed Excel's "
                                     "255-character limit; shorten or reduce them.")
                ws.data_validation(data_first, xc, data_last, xc, {
                    "validate": "list", "source": source, "error_title": m.columns[c],
                    "error_message": "Choose a value from the list."})
            if cs.note:
                ws.write_comment(header, xc, cs.note, {"x_scale": 1.6, "y_scale": 1.0,
                                                       "author": "doc-mantap-mcp"})
            if cs.sparkline:
                a, b = m.index(cs.sparkline.from_column), m.index(cs.sparkline.to_column)
                if a is None or b is None:
                    raise ValueError(f"Sparkline in '{m.columns[c]}' refers to unknown columns "
                                     f"'{cs.sparkline.from_column}'/'{cs.sparkline.to_column}'.")
                for r in range(data_first, data_last + 1):
                    rng = (f"'{m.sheet}'!{xl_rowcol_to_cell(r, first + min(a, b))}:"
                           f"{xl_rowcol_to_cell(r, first + max(a, b))}")
                    ws.add_sparkline(r, xc, {
                        "range": rng, "type": cs.sparkline.type,
                        "series_color": f"#{t.accent}", "markers": cs.sparkline.type == "line",
                        "high_point": True, "low_point": True, "negative_points": True,
                        "markers_color": f"#{t.primary}", "high_color": f"#{t.primary}",
                        "low_color": "#DC2626", "negative_color": "#DC2626"})
        for rule in spec.highlight:
            self._highlight(ws, m, rule, data_first, data_last)
        row = last_row + 1
        if spec.note:
            ws.write_string(row, first, spec.note, self.fmt(italic=True, font_size=8.5,
                                                            font_color=f"#{t.muted}"))
            row += 1
        rng = f"{xl_rowcol_to_cell(header, first)}:{xl_rowcol_to_cell(last_row, last)}"
        self.summary.append(f"  table {m.name}: {rng} ({len(m.rows)} rows"
                            + (", total row" if m.has_total else "") + ")")
        return row + 1, header

    def _total_value(self, m: TableModel, c: int, func: str):
        try:
            values = [v for v in m.column(c, self.ev)
                      if isinstance(v, (int, float)) and not isinstance(v, bool)]
        except Unsupported:
            return None
        if func == "count":
            return float(sum(1 for r in range(len(m.rows)) if m.cached(r, c, self.ev)
                             not in (None, "")))
        if not values:
            return None
        if func == "sum":
            return float(sum(values))
        if func == "average":
            return sum(values) / len(values)
        if func == "min":
            return min(values)
        if func == "max":
            return max(values)
        if func == "count_nums":
            return float(len(values))
        return None

    def _conditional(self, ws, r0, c, r1, kind: str) -> None:
        t = self.t
        if kind == "data_bar":
            opts = {"type": "data_bar", "bar_color": f"#{tint(t.accent, 0.15)}",
                    "bar_solid": True, "bar_negative_color": "#F87171", "data_bar_2010": True,
                    "bar_border_color": f"#{tint(t.accent, 0.15)}"}
        elif kind == "heatmap":
            opts = {"type": "2_color_scale", "min_color": "#FFFFFF",
                    "max_color": f"#{tint(t.accent, 0.35)}"}
        elif kind == "color_scale":
            opts = {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84",
                    "max_color": "#63BE7B"}
        elif kind == "color_scale_reverse":
            opts = {"type": "3_color_scale", "min_color": "#63BE7B", "mid_color": "#FFEB84",
                    "max_color": "#F8696B"}
        elif kind == "arrows":
            opts = {"type": "icon_set", "icon_style": "3_arrows"}
        elif kind == "traffic_lights":
            opts = {"type": "icon_set", "icon_style": "3_traffic_lights"}
        elif kind == "negative_red":
            opts = {"type": "cell", "criteria": "<", "value": 0,
                    "format": self.wb.add_format({"font_color": "#C00000"})}
        else:  # positive_negative
            ws.conditional_format(r0, c, r1, c, {"type": "cell", "criteria": ">", "value": 0,
                                                 "format": self.wb.add_format(
                                                     {"font_color": "#047857"})})
            opts = {"type": "cell", "criteria": "<", "value": 0,
                    "format": self.wb.add_format({"font_color": "#C00000"})}
        ws.conditional_format(r0, c, r1, c, opts)

    def _highlight(self, ws, m: TableModel, rule, r0: int, r1: int) -> None:
        t = self.t
        idx = m.index(rule.column)
        if idx is None:
            raise ValueError(f"Highlight rule refers to unknown column '{rule.column}' in table "
                             f"'{m.name}'. Columns: {', '.join(m.columns)}.")
        bg, fg = _STYLES.get(rule.style, (tint(t.accent, 0.82), shade(t.accent, 0.3)))
        fmt = self.wb.add_format({"bg_color": f"#{bg}", "font_color": f"#{fg}"})
        c = m.first_col + idx
        value = rule.value

        def lit(v):
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return str(v)
            return '"' + str(v).replace('"', '""') + '"'

        when = rule.when
        if rule.entire_row:
            ref = f"${xl_col_to_name(c)}{r0 + 1}"
            if when in ("<", "<=", ">", ">=", "="):
                formula = f"={ref}{when}{lit(value)}"
            elif when == "!=":
                formula = f"={ref}<>{lit(value)}"
            elif when == "between":
                lo, hi = _pair(value, rule.column)
                formula = f"=AND({ref}>={lit(lo)},{ref}<={lit(hi)})"
            elif when == "contains":
                formula = f"=ISNUMBER(SEARCH({lit(value)},{ref}))"
            elif when == "blank":
                formula = f"=LEN(TRIM({ref}))=0"
            else:
                raise ValueError(f"Highlight '{when}' cannot be used with entire_row.")
            ws.conditional_format(r0, m.first_col, r1, m.first_col + len(m.columns) - 1,
                                  {"type": "formula", "criteria": formula, "format": fmt})
            return
        if when in ("<", "<=", ">", ">=", "=", "!="):
            criteria = {"!=": "<>"}.get(when, when)
            v = value if isinstance(value, (int, float)) and not isinstance(value, bool) \
                else lit(value)
            opts = {"type": "cell", "criteria": criteria, "value": v}
        elif when == "between":
            lo, hi = _pair(value, rule.column)
            opts = {"type": "cell", "criteria": "between", "minimum": lo, "maximum": hi}
        elif when == "contains":
            opts = {"type": "text", "criteria": "containing", "value": str(value)}
        elif when in ("top", "bottom"):
            opts = {"type": when, "value": int(value or 3)}
        elif when in ("above_average", "below_average"):
            opts = {"type": "average", "criteria": when.split("_")[0]}
        elif when == "duplicate":
            opts = {"type": "duplicate"}
        else:
            opts = {"type": "blanks"}
        opts["format"] = fmt
        ws.conditional_format(r0, c, r1, c, opts)

    # ---- charts -------------------------------------------------------------------
    def write_charts(self) -> None:
        for sheet_name, (ws, slots) in self.sheet_slots.items():
            for chart, row, col, offset, w, h in slots:
                model = self._chart_table(chart, sheet_name)
                xl_chart = self._chart(chart, model)
                xl_chart.set_size({"width": w, "height": h})
                ws.insert_chart(row, col, xl_chart, {"x_offset": offset, "y_offset": 4,
                                                     "object_position": 1})

    def _chart_table(self, chart: Chart, sheet_name: str) -> TableModel:
        if chart.table:
            key = chart.table.strip().lower()
            model = self.tables.get(key) or self.titles.get(key)
            if model is None:
                known = ", ".join(m.name for m in self.tables.values()) or "none"
                raise ValueError(f"Chart '{chart.title or chart.type}' refers to unknown table "
                                 f"'{chart.table}'. Tables: {known}.")
            return model
        for model in self.tables.values():
            if model.sheet == sheet_name:
                return model
        raise ValueError(f"Chart '{chart.title or chart.type}' on sheet '{sheet_name}' has no "
                         "table to plot. Add a table, or set 'table', or give labels + series.")

    def _chart(self, chart: Chart, m: TableModel):
        t = self.t
        kind = chart.type
        sub = None
        if kind.startswith("stacked_"):
            kind, sub = kind[8:], "stacked"
        elif kind.startswith("percent_stacked_"):
            kind, sub = kind[16:], "percent_stacked"
        x_idx = m.index(chart.x) if chart.x else 0
        if x_idx is None:
            raise ValueError(f"Chart '{chart.title}': table '{m.name}' has no column '{chart.x}'.")
        secondary = [s.lower() for s in (chart.secondary or [])]
        if chart.y:
            ys = []
            for y in chart.y + [s for s in (chart.secondary or []) if s.lower() not in
                                [v.lower() for v in chart.y]]:
                i = m.index(y)
                if i is None:
                    raise ValueError(f"Chart '{chart.title}': table '{m.name}' has no column "
                                     f"'{y}'. Columns: {', '.join(m.columns)}.")
                ys.append(i)
        else:
            ys = [i for i, k in enumerate(m.kinds) if i != x_idx and k in NUMERIC
                  and not m.col_specs[i].sparkline]
            if not ys:
                raise ValueError(f"Chart '{chart.title}': table '{m.name}' has no numeric "
                                 "columns; set 'y'.")
            if kind in ("pie", "doughnut"):
                ys = ys[:1]
        opts = {"type": kind}
        if sub:
            opts["subtype"] = sub
        if kind == "scatter":
            opts["subtype"] = "marker_only"
        xl = self.wb.add_chart(opts)
        line2 = None
        r0, r1 = m.header_row + 1, m.header_row + len(m.rows)
        colors = [c for c in ([t.primary, t.accent] + list(t.palette))]
        seen, palette = set(), []
        for c in colors:
            if c not in seen:
                seen.add(c)
                palette.append(c)
        font = {"name": t.docx_body, "size": 9, "color": f"#{t.muted}"}
        num_fmt = chart.number_format or (m.formats[ys[0]] if ys else None)
        for n, i in enumerate(ys):
            color = f"#{palette[n % len(palette)]}"
            series = {
                "name": [m.sheet, m.header_row, m.first_col + i],
                "categories": [m.sheet, r0, m.first_col + x_idx, r1, m.first_col + x_idx],
                "values": [m.sheet, r0, m.first_col + i, r1, m.first_col + i],
            }
            on_secondary = m.columns[i].lower() in secondary
            if kind in ("pie", "doughnut"):
                series["points"] = [{"fill": {"color": f"#{palette[k % len(palette)]}"},
                                     "border": {"color": "#FFFFFF", "width": 1.5}}
                                    for k in range(len(m.rows))]
                series["data_labels"] = {"percentage": True, "font": {**font, "color": "#FFFFFF",
                                                                      "bold": True}}
            elif kind in ("line", "scatter") or on_secondary:
                series["line"] = {"color": color, "width": 2.25} if kind != "scatter" \
                    else {"none": True}
                series["marker"] = {"type": "circle", "size": 6, "fill": {"color": "#FFFFFF"
                                    if kind != "scatter" else color},
                                    "border": {"color": color, "width": 1.5}}
            elif kind == "area":
                series["fill"] = {"color": color, "transparency": 25 if sub is None else 0}
                series["line"] = {"color": color, "width": 1.5}
            else:
                series["fill"] = {"color": color}
                series["border"] = {"none": True}
                series["gap"] = 70 if len(ys) > 1 else 90
                if sub:
                    series["overlap"] = 100
            if chart.data_labels and kind not in ("pie", "doughnut"):
                series["data_labels"] = {"value": True, "num_format": num_fmt or "General",
                                         "font": {**font, "color": f"#{t.text}"}}
            if on_secondary:
                series["y2_axis"] = True
            if on_secondary and kind in ("column", "bar", "area"):
                if line2 is None:
                    line2 = self.wb.add_chart({"type": "line"})
                line2.add_series(series)
            else:
                xl.add_series(series)
        if line2 is not None:
            xl.combine(line2)
            line2.set_y2_axis({"num_font": font, "line": {"none": True},
                               "major_tick_mark": "none"})
        if chart.title:
            xl.set_title({"name": chart.title, "overlay": False,
                          "name_font": {"name": t.docx_heading, "size": 12, "bold": True,
                                        "color": f"#{t.text}"}})
        else:
            xl.set_title({"none": True})
        if chart.legend == "none" or (len(ys) == 1 and kind not in ("pie", "doughnut")
                                      and chart.legend == "bottom"):
            xl.set_legend({"none": True})
        else:
            xl.set_legend({"position": chart.legend, "font": font})
        grid = f"#{mix(t.surface, t.muted, 0.2)}"
        if kind not in ("pie", "doughnut"):
            cat_axis = {"num_font": font, "line": {"color": grid}, "major_tick_mark": "none",
                        "name": chart.x_title, "name_font": font}
            val_axis = {"num_font": font, "line": {"none": True}, "major_tick_mark": "none",
                        "major_gridlines": {"visible": True, "line": {"color": grid,
                                                                      "width": 0.75}},
                        "name": chart.y_title, "name_font": font}
            if num_fmt and sub != "percent_stacked":
                val_axis["num_format"] = num_fmt
            xl.set_x_axis(cat_axis)
            xl.set_y_axis(val_axis)
        else:
            if kind == "doughnut":
                xl.set_hole_size(55)
        xl.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
        xl.set_plotarea({"border": {"none": True}})
        return xl


def _props(fmt) -> dict:
    """Recover the properties a cached format was created with."""
    return dict(getattr(fmt, "_dm_props", {}))


def _auto_format(value, kind: str | None, currency: str) -> str | None:
    """Number format for a value that has no explicit format."""
    if kind and kind in TYPE_NAMES and kind not in ("number", "integer"):
        return number_format(kind, None, currency)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return "#,##0" if float(value).is_integer() or abs(value) >= 1000 else "#,##0.00"


def _this_row(formula: str) -> str:
    """'=[@Qty]*[@[Unit Price]]' -> '=[[#This Row],[Qty]]*[[#This Row],[Unit Price]]'.

    Excel stores the Excel 2007 form in the file; string literals are left alone.
    """
    f = formula if formula.startswith("=") else "=" + formula
    out, i, in_str = [], 0, False
    while i < len(f):
        ch = f[i]
        if ch == '"':
            in_str = not in_str
        if not in_str and f.startswith("[@", i):
            j = i + 2
            if j < len(f) and f[j] == "[":
                out.append("[[#This Row],")
                i = j
                continue
            end = f.find("]", j)
            if end < 0:
                raise ValueError(f"Unbalanced brackets in formula {formula}")
            name = f[j:end].strip()
            out.append(f"[[#This Row],[{name}]]" if name else "[#This Row]")
            i = end + 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _pair(value, column: str):
    if not isinstance(value, (list, tuple)) or len(value) != 2:
        raise ValueError(f"Highlight 'between' on '{column}' needs value: [min, max].")
    return value[0], value[1]


def _partition(widths: list[float], n: int) -> list[tuple[int, int]]:
    """Split column indexes into n contiguous groups of similar total width."""
    total = len(widths)
    if n >= total:
        return [(i, i) for i in range(n)]
    target = sum(widths) / n
    groups, start, acc = [], 0, 0.0
    for i, w in enumerate(widths):
        acc += w
        remaining_groups = n - len(groups) - 1
        remaining_cols = total - i - 1
        if (acc >= target * (len(groups) + 1) and remaining_groups > 0) or \
                remaining_cols == remaining_groups and remaining_groups > 0:
            groups.append((start, i))
            start = i + 1
    groups.append((start, total - 1))
    return groups


def render_xlsx(output_path: str, sheets: list[Sheet], theme: Theme, title: str | None = None,
                author: str | None = None, company: str | None = None,
                currency: str | None = None, lang: str = "en") -> XlsxResult:
    if not sheets:
        raise ValueError("Provide at least one sheet.")
    out = prepare_output(output_path, ".xlsx")
    code = (lang or "en").split("-")[0].lower()
    currency = currency or ("Rp" if code == "id" else "$")
    builder = WorkbookBuilder(out, theme, currency, lang)
    builder.sheet_slots = {}
    try:
        resolved = builder.build_models([s.model_copy(deep=True) for s in sheets])
        for i, (name, sheet) in enumerate(resolved):
            builder.summary.append(f"sheet '{name}'")
            builder.write_sheet(name, sheet, i)
        builder.write_charts()
        for name, sheet in resolved:
            if sheet.charts:
                builder.summary.append(f"  {len(sheet.charts)} chart(s) on '{name}'")
        builder.wb.set_properties({"title": title or resolved[0][1].title or "",
                                   "author": author or "", "company": company or "",
                                   "comments": "Created with doc-mantap-mcp"})
        builder.wb.close()
    except Exception:
        try:
            builder.wb.fileclosed = True  # do not write a half-built file
        finally:
            if out.exists() and out.stat().st_size == 0:
                out.unlink()
        raise
    return XlsxResult(out, builder.summary, builder.warnings)
