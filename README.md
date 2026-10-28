# doc-mantap-mcp

An MCP server that makes documents written by AI look beautiful.

The AI writes Markdown or a small declarative spec, and doc-mantap-mcp turns it into a designed PDF, Word document or Excel workbook: theme, cover page, running header and footer, table of contents, styled tables, callouts, KPI cards and native charts.

## Tools

- `list_themes`: available themes with colors, fonts and default cover
- `design_guide`: syntax reference (callouts, KPI cards, charts, Excel spec) and design rules
- `markdown_to_pdf`: Markdown to a polished PDF
- `markdown_to_docx`: Markdown to an editable Word document
- `docx_read` / `docx_outline`: inspect Word text, structure, formatting and placeholders
- `docx_patch`: make anchored, in-place-safe edits to an existing DOCX/DOCM
- `docx_insert_image`: insert screenshots into paragraphs or table cells
- `docx_find` / `docx_replace_all`: search or replace across body, tables and headers/footers
- `docx_fill_template`: fill `{{key}}` placeholders in one file or a batch
- `pdf_read_text`: extract PDF text page by page, with optional Tesseract OCR fallback
- `doctor`: check dependencies, Office converters, OCR languages and fonts
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

To add Markdown to an existing Word template, call `markdown_to_docx` with
`mode="append"`, `target_path="template.docx"`, and an optional `start_after` paragraph
anchor. Without `output_path`, it writes a sibling `.appended.docx`; use the target path as
`output_path` to update it with a `.bak` backup. Existing document styles and headers/footers
are retained. `docx_patch` and `docx_insert_image` write a new `.patched.docx` by default;
explicit in-place edits create a `.bak` first. Package parts not involved in image insertion
or XML patching remain untouched.

`pdf_read_text` extracts native page text without Office software. Scanned pages use
Tesseract when available (`brew install tesseract tesseract-lang` on macOS or
`apt install tesseract-ocr tesseract-ocr-ind` on Debian/Ubuntu).

`preview_document` tries LibreOffice, Microsoft Word automation (macOS/Windows), then
macOS QuickLook for Office files. It returns rendered image content plus JSON metadata
(page numbers, image paths/URIs and dimensions). Page selection, width/DPI, PNG/JPEG,
and `max_pages` are configurable; cached previews invalidate when file content changes.

`docx://{absolute-path}#outline` exposes the compact DOCX outline as an MCP resource.
`docx_patch` accepts a `batch` of `{path, operations, output_path?}` items and skips
already-applied insert/fill operations by default. `docx_fill_template` accepts a batch
of `{path, data, output_path?}` items; repeat fills are safe no-ops.

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
