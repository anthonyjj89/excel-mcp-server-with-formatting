# Excel MCP Server with Formatting v0.2.0

An enhanced Model Context Protocol (MCP) server for Excel integration with Claude Desktop, adding comprehensive formatting capabilities.

## v0.2.0 Features

### Formula Calculation Fix
- Automatic formula calculation when Excel files are opened
- No more manual clicking or entering formulas to activate them
- Formulas work correctly immediately upon file open

### Auto-Fit Column Width
- Automatically adjust column widths to fit content
- Add `autoFit: true` parameter to write operations
- Dedicated `autofit_columns` tool for precise control
- Support for custom padding and min/max width constraints

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

```
Write formulas to cells A1:B3 in Sheet1 that will calculate the sum, average, and max of values in column C
```

```
Auto-fit columns A through E in the data sheet to match their content width
```

## Credits

Original Excel MCP Server by [negokaz](https://github.com/negokaz).  
Formatting enhancements by [anthonyjj89](https://github.com/anthonyjj89).
