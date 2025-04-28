# Excel MCP Server with Formatting v0.1.0

An enhanced Model Context Protocol (MCP) server for Excel integration with Claude Desktop, adding comprehensive formatting capabilities.

## Enhanced Features

This fork extends the original Excel MCP server with powerful formatting tools:

### Styling Capabilities
- **Font Formatting**: Bold, italic, font size, and custom colors
- **Cell Appearance**: Background colors, borders with custom styles and colors
- **Advanced Layout**: Cell merging/unmerging
- **Worksheet Management**: Add new worksheets easily
- **Comprehensive Styling**: Apply multiple style attributes in a single operation

### Excel Refresh
- Automatic refresh functionality to immediately see changes in Excel

## Original Features

Built upon the excellent [excel-mcp-server](https://github.com/negokaz/excel-mcp-server) by negokaz, this version adds powerful styling tools to make your Excel automation truly professional while maintaining all the core functionality:

- Read sheet names from Excel files
- Read data from Excel sheets
- Read formulas from Excel sheets
- Write data to Excel sheets
- Write formulas to Excel sheets

## Usage

1. Install dependencies:
```bash
npm install @modelcontextprotocol/sdk@latest exceljs zod
```

2. Configure in Claude Desktop:
```json
"excel": {
  "command": "node",
  "args": ["/path/to/excel-mcp-server-with-formatting/excel-mcp.js"],
  "env": {
    "NODE_ENV": "production"
  }
}
```

3. Use in Claude with commands like:
```
Format cells A1:B5 in the "Sheet1" sheet of the Excel file at /path/to/file.xlsx with bold green text and yellow background
```

```
Add a bold black border around cells A1:B10 in the Sample sheet
```

```
Merge cells A1:C1 and center the text in the Header sheet
```

## Credits

Original Excel MCP Server by [negokaz](https://github.com/negokaz).  
Formatting enhancements by [anthonyjj89](https://github.com/anthonyjj89).
