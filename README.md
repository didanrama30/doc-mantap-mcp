# doc-mantap-mcp

An MCP server that makes documents written by AI look beautiful.

The AI writes Markdown, and doc-mantap-mcp turns it into a polished PDF or DOCX with a theme, cover page, page numbers, styled tables, quotes and code blocks.

## Tools

- `list_themes`: show the available themes (`modern`, `corporate`, `academic`, `fresh`)
- `markdown_to_pdf`: Markdown to PDF
- `markdown_to_docx`: Markdown to an editable Word document

## Setup

Requires [uv](https://docs.astral.sh/uv/). PDF output also needs Pango (`brew install pango` on macOS).

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

Then ask your AI something like: "Write a Q3 sales report as a PDF using the corporate theme."

## License

MIT
