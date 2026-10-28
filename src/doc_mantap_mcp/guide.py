"""Authoring guide returned by the design_guide tool."""

DOCUMENTS = r'''
# Documents (markdown_to_pdf / markdown_to_docx)

Write normal Markdown. The theme handles fonts, colors, spacing and page layout.
PDF gives the most polished result; DOCX gives an editable Word file with the same structure.

## Cover, metadata and layout
- `title` creates a designed cover page. Do not repeat the title as a `#` heading.
- Optional cover fields: `subtitle`, `document_type` (small label above the title, e.g.
  "Quarterly Report"), `author`, `date`, `organization`, `logo` (image path).
- `cover_style`: band | full | classic | minimal (each theme has a default).
- `toc: true` adds a table of contents (worth it from about 4 sections).
- `lang`: "id" for Indonesian, "en" for English. Sets hyphenation, labels
  ("Daftar Isi", "Catatan"...) and number formatting in charts.
- `page_size`: A4 | Letter | Legal | A5, `orientation`: portrait | landscape.
- `primary_color` / `accent_color`: brand colors (hex), applied on top of the theme.
- Relative image paths resolve against `base_dir` (default: the output folder).

## Structure
- `#` = chapter (shown in the running header), `##` = section, `###` = subsection.
- Page break: a line containing only `\pagebreak` (or `<!-- pagebreak -->`).
- Lists, tables and code fences may follow a paragraph directly; 2-space nesting works.

## Callouts
```
> [!NOTE]
> Neutral information.

> [!TIP] Custom title
> Advice or a recommendation.
```
Types: NOTE, INFO, TIP, SUCCESS, IMPORTANT, WARNING, CAUTION, DANGER.

## KPI cards (a row of big numbers)
````
```kpi
- label: Revenue
  value: $12.4M
  delta: +18%
  note: vs Q2
- label: Churn
  value: 2.1%
  delta: -0.4 pp
  lower_is_better: true
```
````
Values are shown exactly as written. Up to 4 cards per row (more wrap to a new row).
Delta color follows the sign; `lower_is_better: true` makes decreases green.

## Charts
````
```chart
type: bar            # bar | barh | line | area | pie | donut | scatter
title: Revenue by quarter
subtitle: IDR billion
labels: [Q1, Q2, Q3, Q4]
series:
  - name: 2025
    data: [8.2, 9.1, 10.5, 11.0]
  - name: 2026
    data: [9.0, 10.5, 12.4, 13.1]
caption: "Source: ERP"
```
````
Options: `stacked: true`, `data_labels: true`, `prefix: "$"`, `suffix: "%"`,
`decimals: 1`, `legend: top|bottom|right|none`, `height: 8` (cm), `colors: ["#hex", ...]`,
`x_label`, `y_label`, `y_min`, `y_max`. Single series shorthand: `data: {Q1: 10, Q2: 12}`.
Combo chart: give a series `type: line` and `secondary: true`.
Donut: `center: total` (or any text) writes a label in the hole.
Scatter: `data: [[x, y], ...]` per series. Use plain numbers (12400, not "12.400").

## Other syntax
- Footnotes: `text[^1]` + `[^1]: Note text.` (real page footnotes in PDF).
- Task lists: `- [x] Done` / `- [ ] Todo`.
- `==highlight==`, `~~strike~~`, `H~2~O`, `m^2^`, `` `code` ``, definition lists (`Term` then
  `:   Definition`), `*[HTML]: abbreviation`.
- Images: `![alt](path/to/img.png "Caption")`; width: `![alt](img.png){ width=60% }`.
- Code blocks with a language get syntax highlighting.
- PDF only: raw HTML and inline `<style>` are allowed for special layouts.

## Design rules that make documents look professional
1. Start with a 2-4 sentence summary, then a KPI row if there are headline numbers.
2. One idea per section; keep paragraphs to 2-5 sentences.
3. Tables: 2-7 columns, short headers, units in the header ("Revenue (IDR m)"). Numbers are
   right-aligned automatically.
4. Use a chart when a trend or comparison matters more than exact values; keep the exact
   values in a table if readers need them.
5. Use at most one or two callouts per page, for what the reader must not miss.
6. Prefer bold for key figures over long bold sentences.
7. After rendering, call preview_document on the PDF and fix anything that looks off
   (awkward page breaks, empty pages, cramped tables).
'''

EXCEL = r'''
# Spreadsheets (create_excel)

Describe the workbook; layout and styling are automatic: banner title, KPI cards, an
assumptions block, real Excel tables (filters, banded rows, total row), conditional
formatting, dropdowns, sparklines, native Excel charts, frozen headers and print setup.

## Shape of `sheets`
```json
[{
  "name": "Sales",
  "title": "Sales Dashboard Q3 2026",
  "subtitle": "Data as of 30 Sep 2026",
  "kpis": [
    {"label": "Revenue", "value": "=SUM(Sales[Revenue])", "format": "currency",
     "delta": 0.18, "note": "vs Q2"},
    {"label": "Avg margin", "value": "=AVERAGE(Sales[Margin])", "format": "percent"}
  ],
  "inputs": [
    {"label": "Tax rate", "value": 0.11, "format": "percent", "name": "TaxRate",
     "note": "Source: tax office"}
  ],
  "tables": [{
    "name": "Sales",
    "title": "Performance by region",
    "columns": [
      {"header": "Region", "total": "Total"},
      {"header": "Revenue", "type": "currency", "total": "sum", "conditional": "data_bar"},
      {"header": "Cost", "type": "currency", "total": "sum"},
      {"header": "Margin", "type": "percent",
       "formula": "=IFERROR(([@Revenue]-[@Cost])/[@Revenue],0)", "conditional": "color_scale"},
      {"header": "Tax", "type": "currency", "formula": "=[@Revenue]*TaxRate"},
      {"header": "Status", "options": ["On track", "At risk"]}
    ],
    "rows": [["West", 4200000, 2600000, null, null, "On track"],
             ["East", 5100000, 4900000, null, null, "At risk"]],
    "highlight": [{"column": "Status", "when": "=", "value": "At risk",
                   "style": "bad", "entire_row": true}],
    "note": "Source: ERP export"
  }],
  "charts": [{"type": "column", "title": "Revenue vs cost", "x": "Region",
              "y": ["Revenue", "Cost"]}],
  "notes": ["All amounts in USD."]
}]
```

## Columns
- `type`: text, number, integer, currency, accounting, percent, date, datetime, time,
  multiple, boolean, url. Omit to detect it from the data.
- Percent values are fractions: 0.15 shows as 15.0%.
- Dates: ISO strings "2026-10-28". Numbers: real numbers, no thousands separators.
- `formula`: calculated column using structured references (`[@Col]`, `[@[Col name]]`,
  `Table[Col]`) and input names. Put null for that column in `rows`.
- Any row value starting with "=" is written as a formula.
- `total`: sum | average | count | min | max | a label | a formula.
- `conditional`: data_bar, heatmap, color_scale, color_scale_reverse, arrows,
  traffic_lights, negative_red, positive_negative.
- `options` (dropdown), `input: true` (cells to fill in), `note` (header comment), `width`,
  `wrap`, `align`, `decimals`, `format` (custom Excel format), `sparkline`
  ({"from_column": "Jan", "to_column": "Dec", "type": "line"}).

## Charts
`type`: column, bar (horizontal), line, area, pie, doughnut, scatter, stacked_column,
stacked_bar, stacked_area, percent_stacked_column. Point at a table with `table`, `x`, `y`;
`secondary: ["Margin"]` adds a line on a second axis. Without a table, pass `labels` and
`series` and a small data table is created. `position: "right"` puts it next to the table.

## Rules for workbooks people trust
1. Use formulas for anything derived (totals, margins, KPIs) so the sheet recalculates.
2. Put assumptions in `inputs` and reference them by name (=[@Revenue]*TaxRate),
   never hardcode them inside formulas.
3. Prefer SUMIFS, INDEX/MATCH, IFERROR over XLOOKUP/FILTER for compatibility.
4. One table per topic; name tables so formulas read well (Sales[Revenue]).
5. Name units in headers ("Revenue (USD)"), keep percent as fractions, and guard divisions
   with IFERROR.
6. Add a `note` with the data source.
Formula mistakes (unknown table or column, unbalanced brackets) are reported back as errors
before the file is written. Excel recalculates every formula when the file opens.
'''


def guide(topic: str = "all") -> str:
    topic = (topic or "all").lower()
    if topic.startswith("doc"):
        return DOCUMENTS.strip()
    if topic.startswith(("excel", "xls", "sheet", "spread")):
        return EXCEL.strip()
    return DOCUMENTS.strip() + "\n\n" + EXCEL.strip()
