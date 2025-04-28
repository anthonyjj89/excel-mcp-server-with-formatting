# Excel MCP for Claude with Formatting

This project extends the standard Excel MCP integration for Claude Desktop with comprehensive formatting capabilities.

## Features

- Read Excel sheet names and data
- Read and write Excel formulas
- Format cells (colors, fonts, alignment)
- Add borders to cells
- Merge and unmerge cells
- Add worksheets
- Apply complex formatting styles
- Automatically refresh Excel to view changes

## Forked From

This project is an enhancement of the original Excel MCP. The original idea and base implementation are credited to Anthropic's Claude MCP team.

## Usage

1. Install dependencies:
```bash
npm install @modelcontextprotocol/sdk@latest exceljs zod
```

2. Configure in Claude Desktop:
```json
"excel": {
  "command": "node",
  "args": ["/path/to/excel-mcp-formatting/excel-mcp.js"],
  "env": {
    "NODE_ENV": "production"
  }
}
```

3. Use in Claude with commands like:
```
Format cells A1:B5 in the "Sheet1" sheet of the Excel file at /path/to/file.xlsx with bold green text and yellow background
```
