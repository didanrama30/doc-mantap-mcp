# doc-mantap-mcp

An MCP server that makes documents written by AI look beautiful.

The AI writes Markdown or a small declarative spec, and doc-mantap-mcp turns it into a designed PDF, Word document or Excel workbook: theme, cover page, running header and footer, table of contents, styled tables, callouts, KPI cards and native charts.

## Tools

- `list_themes`: available themes with colors, fonts and default cover
- `design_guide`: syntax reference (callouts, KPI cards, charts, Excel spec) and design rules
- `markdown_to_pdf`: Markdown to a polished PDF
- `markdown_to_docx`: Markdown to an editable Word document
- `create_excel`: designed workbook with real Excel tables, formulas and charts
- `preview_document`: render pages of a generated file as images to check the layout

Themes: `modern`, `corporate`, `academic`, `fresh`, `elegant`, `minimal`, `startup`, `warm`. Brand colors can be overridden per call with `primary_color` and `accent_color`.

## Markdown extras

```markdown
> [!TIP]
> Callouts: note, info, tip, success, important, warning, caution, danger.

\`\`\`kpi
- value: Rp 1.2 M
  label: Revenue
  delta: 0.12
\`\`\`

\`\`\`chart
type: bar
title: Revenue by quarter
labels: [Q1, Q2, Q3, Q4]
series:
  - name: 2026
    data: [120, 135, 158, 172]
\`\`\`
```

Also supported: tables, footnotes, task lists, definition lists, images with captions, syntax-highlighted code, `\pagebreak`, `toc=true`, page size/orientation and `lang` (`en`, `id`, ...) for labels and number formats.

## Setup

Requires [uv](https://docs.astral.sh/uv/).

- PDF output also needs Pango (`brew install pango` on macOS, `libpango-1.0-0 libpangoft2-1.0-0` on Debian/Ubuntu). DOCX and Excel output work without it.
- `preview_document` can open PDFs directly; previewing DOCX/XLSX also needs LibreOffice.

Add this to your MCP client config:

```json
{
  "mcpServers": {
    "doc-mantap-mcp": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/didanrama30/doc-mantap-mcp", "doc-mantap-mcp"]
    }
  }
}
```

Then ask your AI something like: "Write a Q3 sales report as a PDF using the corporate theme" or "Build an Excel budget with monthly charts".

## Development

```bash
uv sync
uv run doc-mantap-mcp
```

Bundled fonts are OFL licensed; regenerate them with `uv run python scripts/fetch_fonts.py`.

## License

MIT
