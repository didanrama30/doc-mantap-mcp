"""Declarative workbook spec for create_excel (also the tool's input schema)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

ValueType = Literal["text", "number", "integer", "currency", "accounting", "percent", "date",
                    "datetime", "time", "multiple", "boolean", "url"]
Conditional = Literal["data_bar", "heatmap", "color_scale", "color_scale_reverse", "arrows",
                      "traffic_lights", "negative_red", "positive_negative"]


class _Spec(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Sparkline(_Spec):
    from_column: str = Field(description="First column (header) of the row range to plot.")
    to_column: str = Field(description="Last column (header) of the row range to plot.")
    type: Literal["line", "column", "win_loss"] = "line"


class Column(_Spec):
    header: str = Field(description="Column header. Must be unique within the table.")
    type: ValueType | None = Field(
        None, description="Value type; picks the number format. Omit to detect from the data. "
        "Percent values are fractions (0.15 shows as 15.0%).")
    format: str | None = Field(None, description="Custom Excel number format, overrides type, "
                               "e.g. '0.0\"x\"' or '#,##0 \"kg\"'.")
    decimals: int | None = Field(None, ge=0, le=6, description="Decimal places for numbers.")
    width: float | None = Field(None, gt=0, le=120,
                                description="Width in characters. Default: fit to content.")
    formula: str | None = Field(
        None, description="Calculated column, same formula in every row, using structured "
        "references, e.g. '=[@Qty]*[@[Unit Price]]'. Leave this column out of the row values "
        "or put null there.")
    total: str | None = Field(
        None, description="Total-row cell: sum, average, count, min, max, a label such as "
        "'Total', or a formula. Adding any total turns the total row on.")
    conditional: Conditional | None = Field(
        None, description="Visual rule for the column: data_bar, heatmap, color_scale "
        "(red-yellow-green), color_scale_reverse, arrows, traffic_lights, negative_red, "
        "positive_negative.")
    options: list[str] | None = Field(
        None, description="Allowed values; adds a dropdown list (data validation).")
    input: bool = Field(False, description="Style as cells the user fills in (blue text on "
                        "light yellow), the financial-model convention.")
    align: Literal["left", "center", "right"] | None = None
    wrap: bool = Field(False, description="Wrap long text inside the cells.")
    note: str | None = Field(None, description="Comment shown on the header cell, e.g. the "
                             "source or definition of the column.")
    sparkline: Sparkline | None = Field(
        None, description="Draw an in-cell sparkline from other columns of the same row "
        "(e.g. Jan..Dec). Row values for this column are ignored.")


class Highlight(_Spec):
    column: str = Field(description="Header of the column the rule tests.")
    when: Literal["<", "<=", ">", ">=", "=", "!=", "between", "contains", "top", "bottom",
                  "above_average", "below_average", "duplicate", "blank"] = Field(
        description="Condition. 'top'/'bottom' take a count as value; 'between' takes [min, max].")
    value: Any = Field(None, description="Comparison value (number or text).")
    style: Literal["good", "bad", "neutral", "accent"] = Field(
        "bad", description="good = green, bad = red, neutral = yellow, accent = theme color.")
    entire_row: bool = Field(False, description="Color the whole table row, not just the cell.")


class Table(_Spec):
    name: str | None = Field(None, description="Excel table name used in formulas, e.g. 'Sales' "
                             "in =SUM(Sales[Revenue]). Default: derived from the title.")
    title: str | None = Field(None, description="Section heading shown above the table.")
    description: str | None = Field(None, description="Short text under the heading.")
    columns: list[Column | str] = Field(
        description="Column definitions, or plain header strings.", min_length=1)
    rows: list[list[Any]] = Field(
        default_factory=list, description="Row values in column order. Numbers must be numbers, "
        "dates ISO strings ('2026-10-28'), formulas strings starting with '='.")
    blank_rows: int = Field(0, ge=0, le=500, description="Extra empty rows for data entry.")
    total_row: bool | None = Field(None, description="Force the total row on or off.")
    style: str | None = Field(None, description="Excel table style name, e.g. 'Table Style "
                              "Medium 9'. Default follows the theme.")
    highlight: list[Highlight] = Field(default_factory=list,
                                       description="Rule-based cell or row coloring.")
    note: str | None = Field(None, description="Small note under the table, e.g. the source.")


class InlineSeries(_Spec):
    name: str
    data: list[float | None]


class Chart(_Spec):
    type: Literal["column", "bar", "line", "area", "pie", "doughnut", "scatter",
                  "stacked_column", "stacked_bar", "stacked_area", "percent_stacked_column"] = \
        Field("column", description="column = vertical bars, bar = horizontal bars.")
    title: str | None = None
    table: str | None = Field(None, description="Name or title of the table to plot (any sheet). "
                              "Default: the first table on this sheet.")
    x: str | None = Field(None, description="Category column header. Default: first column.")
    y: list[str] | None = Field(None, description="Value column headers. Default: all numeric "
                                "columns.")
    secondary: list[str] | None = Field(
        None, description="Value columns drawn as lines on a second axis (combo chart).")
    labels: list[str] | None = Field(None, description="Inline categories when there is no "
                                     "table; a small data table is added to the sheet.")
    series: list[InlineSeries] | None = Field(None, description="Inline values for 'labels'.")
    position: Literal["below", "right"] = Field(
        "below", description="Place below all tables, or to the right of the first table.")
    width: int | None = Field(None, ge=240, le=1600, description="Width in pixels.")
    height: int | None = Field(None, ge=160, le=1000, description="Height in pixels.")
    data_labels: bool = False
    legend: Literal["bottom", "top", "right", "none"] = "bottom"
    number_format: str | None = Field(None, description="Axis/label format, e.g. '#,##0'.")
    x_title: str | None = None
    y_title: str | None = None


class Kpi(_Spec):
    label: str
    value: Any = Field(description="Number, text, or a formula such as "
                       "'=SUM(Sales[Revenue])'.")
    format: str | None = Field(None, description="A value type (currency, percent, integer, "
                               "number, ...) or a custom Excel number format.")
    delta: Any = Field(None, description="Change vs previous period: a number (fraction for "
                       "percent), a formula, or text like '+12%'.")
    delta_format: str | None = Field(None, description="Type or format for a numeric delta. "
                                     "Default: percent.")
    note: str | None = Field(None, description="Small caption, e.g. 'vs Q2'.")
    lower_is_better: bool = Field(False, description="Show decreases in green (costs, churn).")


class Input(_Spec):
    label: str
    value: Any = Field(description="The assumption's value (number, text or date).")
    name: str | None = Field(None, description="Defined name for formulas, e.g. 'TaxRate' in "
                             "=[@Revenue]*TaxRate. Default: derived from the label.")
    format: str | None = Field(None, description="Value type or custom number format.")
    note: str | None = Field(None, description="Source or explanation shown next to the value.")


class Sheet(_Spec):
    name: str = Field(description="Tab name (max 31 characters).")
    title: str | None = Field(None, description="Banner title at the top of the sheet.")
    subtitle: str | None = None
    kpis: list[Kpi] = Field(default_factory=list, description="Metric cards under the banner.")
    inputs: list[Input] = Field(default_factory=list, description="Assumptions block: labeled "
                                "input cells with defined names that formulas can use.")
    inputs_title: str | None = None
    tables: list[Table] = Field(default_factory=list)
    charts: list[Chart] = Field(default_factory=list, description="Native Excel charts.")
    notes: list[str] = Field(default_factory=list, description="Footnotes at the bottom.")
    tab_color: str | None = Field(None, description="Hex color for the sheet tab.")
    landscape: bool | None = Field(None, description="Print orientation. Default: automatic.")
    freeze: bool | None = Field(None, description="Freeze the first table's header row. "
                                "Default: automatic.")
    gridlines: bool = Field(False, description="Show Excel gridlines (hidden by default for a "
                            "clean look).")
    zoom: int | None = Field(None, ge=10, le=400)
