# Excel MCP v0.2 Updates

This document outlines the changes and new features included in version 0.2 of the Excel MCP.

## Major Enhancements

### 1. Formula Calculation Fix

We've fixed the issue where formulas were not being properly calculated in Excel. The fix includes:

- Setting `workbook.calcProperties.fullCalcOnLoad = true` to force Excel to recalculate all formulas when the file is opened
- Setting `cell.model.result = undefined` for each formula cell to ensure it's marked for recalculation
- Improved formula handling across all tools that write to Excel files

### 2. Auto-fit Column Width Feature

We've added the ability to automatically adjust column widths based on content:

- Added `autoFit` parameter to `write_sheet_data` and `write_sheet_formula` tools
- Created a dedicated `autofit_columns` tool for more precise control
- Intelligent width calculation based on content length
- Support for constraints like minimum/maximum width and custom padding

## Integration Guide

### Using Formula Fix

The formula fix is automatically applied whenever you write formulas to a sheet. To ensure proper calculation:

```javascript
// Example of writing formulas with auto-calculation
await excel_mcp.write_sheet_formula({
  fileAbsolutePath: '/path/to/file.xlsx',
  sheetName: 'Sheet1',
  range: 'A1:B3',
  formulas: [
    ['=SUM(C1:C10)', '=AVERAGE(D1:D10)'],
    ['=TODAY()', '=NOW()'],
    ['=A1+B1', '=A2-B2']
  ]
});
```

### Using Auto-fit Column Width

You can enable auto-fit column width in two ways:

1. **Use with existing write operations**:

```javascript
// Example of writing data with auto-fit columns
await excel_mcp.write_sheet_data({
  fileAbsolutePath: '/path/to/file.xlsx',
  sheetName: 'Sheet1',
  range: 'A1:C10',
  data: yourData,
  autoFit: true  // Enable auto-fit columns
});
```

2. **Use the dedicated tool for more control**:

```javascript
// Example of using the dedicated auto-fit tool
await excel_mcp.autofit_columns({
  fileAbsolutePath: '/path/to/file.xlsx',
  sheetName: 'Sheet1',
  columns: ['A', 'B', 'C'],  // Specific columns to adjust
  padding: 3,                // Extra characters of padding
  minWidth: 10,              // Minimum column width
  maxWidth: 50               // Maximum column width
});
```

## Implementation Notes

### How to Apply These Updates

1. The `fix-formula-issues.js` file contains the implementation of these features
2. Review and incorporate these changes into your main `excel-mcp.js` file
3. Update version number to 0.2.0 in both `package.json` and `excel-mcp.js`
4. Test thoroughly with various formula scenarios and data types

### Backward Compatibility

These updates are fully backward compatible. The `autoFit` parameter is optional, so existing code will continue to work without modification.

## Next Steps

1. Integrate these updates into the main codebase
2. Update tests to verify formula calculation and auto-width behavior
3. Update documentation to reflect the new features
4. Create examples demonstrating the new capabilities
