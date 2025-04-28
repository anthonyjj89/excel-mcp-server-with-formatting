#!/usr/bin/env node

const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const ExcelJS = require('exceljs');
const z = require('zod');

// Create utility functions since ExcelJS.utils doesn't exist
function columnNameToNumber(name) {
  let result = 0;
  for (let i = 0; i < name.length; i++) {
    result = result * 26 + (name.charCodeAt(i) - 64);
  }
  return result;
}

function numberToColumnName(num) {
  let result = '';
  while (num > 0) {
    const modulo = (num - 1) % 26;
    result = String.fromCharCode(65 + modulo) + result;
    num = Math.floor((num - modulo) / 26);
  }
  return result;
}

function parseRange(range) {
  const [start, end] = range.split(':');
  const startMatch = start.match(/([A-Z]+)(\d+)/);
  const endMatch = end.match(/([A-Z]+)(\d+)/);
  
  if (!startMatch || !endMatch) {
    throw new Error(`Invalid range format: ${range}`);
  }
  
  return {
    startCol: startMatch[1],
    startRow: parseInt(startMatch[2]),
    endCol: endMatch[1],
    endRow: parseInt(endMatch[2])
  };
}

async function main() {
  // Create the MCP server
  const server = new McpServer({
    name: "Excel MCP",
    version: "1.0.0"
  }, {
    capabilities: {
      tools: { listChanged: true } 
    }
  });

  // Register the read_sheet_names tool
  server.tool(
    'read_sheet_names',
    'List all sheet names in an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file')
    },
    async ({ fileAbsolutePath }) => {
      try {
        console.error(`Reading sheet names from ${fileAbsolutePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const sheetNames = workbook.worksheets.map(sheet => sheet.name);
        console.error(`Found ${sheetNames.length} sheets: ${sheetNames.join(', ')}`);
        
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({ sheetNames }) 
          }]
        };
      } catch (error) {
        console.error(`Error reading sheet names: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to read sheet names: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register the read_sheet_data tool
  server.tool(
    'read_sheet_data',
    'Read data from Excel sheet with pagination',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().optional().describe('Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]'),
      knownPagingRanges: z.array(z.string()).optional().describe('List of already read paging ranges')
    },
    async ({ fileAbsolutePath, sheetName, range, knownPagingRanges }) => {
      try {
        console.error(`Reading data from ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range || 'all'}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        let data = [];
        
        if (range) {
          const { startCol, startRow, endCol, endRow } = parseRange(range);
          const startColNum = columnNameToNumber(startCol);
          const endColNum = columnNameToNumber(endCol);
          
          for (let row = startRow; row <= endRow; row++) {
            const rowData = [];
            for (let col = startColNum; col <= endColNum; col++) {
              const cellAddress = `${numberToColumnName(col)}${row}`;
              const cell = worksheet.getCell(cellAddress);
              rowData.push(cell.value);
            }
            data.push(rowData);
          }
        } else {
          // Read all data
          worksheet.eachRow((row, rowNum) => {
            const rowData = [];
            row.eachCell((cell) => {
              rowData.push(cell.value);
            });
            data.push(rowData);
          });
        }
        
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({ data }) 
          }]
        };
      } catch (error) {
        console.error(`Error reading sheet data: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to read sheet data: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register read_sheet_formula tool
  server.tool(
    'read_sheet_formula',
    'Read formulas from Excel sheet with pagination',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().optional().describe('Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]'),
      knownPagingRanges: z.array(z.string()).optional().describe('List of already read paging ranges')
    },
    async ({ fileAbsolutePath, sheetName, range, knownPagingRanges }) => {
      try {
        console.error(`Reading formulas from ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range || 'all'}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        let formulas = [];
        
        if (range) {
          const { startCol, startRow, endCol, endRow } = parseRange(range);
          const startColNum = columnNameToNumber(startCol);
          const endColNum = columnNameToNumber(endCol);
          
          for (let row = startRow; row <= endRow; row++) {
            const rowFormulas = [];
            for (let col = startColNum; col <= endColNum; col++) {
              const cellAddress = `${numberToColumnName(col)}${row}`;
              const cell = worksheet.getCell(cellAddress);
              rowFormulas.push(cell.formula || null);
            }
            formulas.push(rowFormulas);
          }
        } else {
          // Read all formulas
          worksheet.eachRow((row, rowNum) => {
            const rowFormulas = [];
            row.eachCell((cell) => {
              rowFormulas.push(cell.formula || null);
            });
            formulas.push(rowFormulas);
          });
        }
        
        return { 
          content: [{ 
            type: "text", 
            text: JSON.stringify({ formulas }) 
          }]
        };
      } catch (error) {
        console.error(`Error reading sheet formulas: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to read sheet formulas: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register write_sheet_data tool
  server.tool(
    'write_sheet_data',
    'Write data to the Excel sheet',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe('Data to write to the Excel sheet')
    },
    async ({ fileAbsolutePath, sheetName, range, data }) => {
      try {
        console.error(`Writing data to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        
        // Try to read existing file, create new if doesn't exist
        try {
          await workbook.xlsx.readFile(fileAbsolutePath);
        } catch (e) {
          console.error(`File ${fileAbsolutePath} doesn't exist. Creating a new workbook.`);
        }
        
        // Get or create worksheet
        let worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          worksheet = workbook.addWorksheet(sheetName);
        }
        
        // Parse range and write data
        const { startCol, startRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          for (let j = 0; j < row.length; j++) {
            const cellAddress = `${numberToColumnName(startColNum + j)}${startRow + i}`;
            worksheet.getCell(cellAddress).value = row[j];
          }
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully wrote data to ${sheetName} in range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error writing sheet data: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to write sheet data: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register write_sheet_formula tool
  server.tool(
    'write_sheet_formula',
    'Write formulas to the Excel sheet',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
      formulas: z.array(z.array(z.string())).describe('Formulas to write to the Excel sheet (e.g., "=A1+B1")')
    },
    async ({ fileAbsolutePath, sheetName, range, formulas }) => {
      try {
        console.error(`Writing formulas to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        
        // Try to read existing file, create new if doesn't exist
        try {
          await workbook.xlsx.readFile(fileAbsolutePath);
        } catch (e) {
          console.error(`File ${fileAbsolutePath} doesn't exist. Creating a new workbook.`);
        }
        
        // Get or create worksheet
        let worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          worksheet = workbook.addWorksheet(sheetName);
        }
        
        // Parse range and write formulas
        const { startCol, startRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        
        for (let i = 0; i < formulas.length; i++) {
          const row = formulas[i];
          for (let j = 0; j < row.length; j++) {
            const cellAddress = `${numberToColumnName(startColNum + j)}${startRow + i}`;
            if (row[j].startsWith('=')) {
              worksheet.getCell(cellAddress).value = { formula: row[j].substring(1) };
            } else {
              worksheet.getCell(cellAddress).value = { formula: row[j] };
            }
          }
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully wrote formulas to ${sheetName} in range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error writing sheet formulas: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to write sheet formulas: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register add_borders tool
  server.tool(
    'add_borders',
    'Add borders to cells in an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells to add borders to (e.g., "A1:C10")'),
      borderStyle: z.object({
        top: z.object({ style: z.string(), color: z.string() }).optional(),
        bottom: z.object({ style: z.string(), color: z.string() }).optional(),
        left: z.object({ style: z.string(), color: z.string() }).optional(),
        right: z.object({ style: z.string(), color: z.string() }).optional()
      }).describe('Border style options')
    },
    async ({ fileAbsolutePath, sheetName, range, borderStyle }) => {
      try {
        console.error(`Adding borders to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        const { startCol, startRow, endCol, endRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Create border style object for ExcelJS
        const border = {};
        if (borderStyle.top) {
          border.top = { style: borderStyle.top.style, color: { argb: borderStyle.top.color } };
        }
        if (borderStyle.bottom) {
          border.bottom = { style: borderStyle.bottom.style, color: { argb: borderStyle.bottom.color } };
        }
        if (borderStyle.left) {
          border.left = { style: borderStyle.left.style, color: { argb: borderStyle.left.color } };
        }
        if (borderStyle.right) {
          border.right = { style: borderStyle.right.style, color: { argb: borderStyle.right.color } };
        }
        
        // Apply border to each cell in the range
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColNum; col <= endColNum; col++) {
            const cellAddress = `${numberToColumnName(col)}${row}`;
            worksheet.getCell(cellAddress).border = border;
          }
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully added borders to ${sheetName} in range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error adding borders: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to add borders: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register format_cells tool
  server.tool(
    'format_cells',
    'Format cells in an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells to format (e.g., "A1:C10")'),
      formatting: z.object({
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        fontSize: z.number().optional(),
        fontColor: z.string().optional(),
        fillColor: z.string().optional(),
        alignment: z.object({
          horizontal: z.string().optional(),
          vertical: z.string().optional()
        }).optional()
      }).describe('Formatting options')
    },
    async ({ fileAbsolutePath, sheetName, range, formatting }) => {
      try {
        console.error(`Formatting cells in ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        const { startCol, startRow, endCol, endRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Apply formatting to each cell in the range
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColNum; col <= endColNum; col++) {
            const cellAddress = `${numberToColumnName(col)}${row}`;
            const cell = worksheet.getCell(cellAddress);
            
            // Font formatting
            if (!cell.font) cell.font = {};
            if (formatting.bold !== undefined) cell.font.bold = formatting.bold;
            if (formatting.italic !== undefined) cell.font.italic = formatting.italic;
            if (formatting.fontSize !== undefined) cell.font.size = formatting.fontSize;
            if (formatting.fontColor !== undefined) cell.font.color = { argb: formatting.fontColor };
            
            // Fill formatting
            if (formatting.fillColor !== undefined) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: formatting.fillColor }
              };
            }
            
            // Alignment formatting
            if (formatting.alignment) {
              cell.alignment = {};
              if (formatting.alignment.horizontal) cell.alignment.horizontal = formatting.alignment.horizontal;
              if (formatting.alignment.vertical) cell.alignment.vertical = formatting.alignment.vertical;
            }
          }
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully formatted cells in ${sheetName} range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error formatting cells: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to format cells: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register merge_cells tool
  server.tool(
    'merge_cells',
    'Merge cells in an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells to merge (e.g., "A1:C10")')
    },
    async ({ fileAbsolutePath, sheetName, range }) => {
      try {
        console.error(`Merging cells in ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        // Merge the cells in the specified range
        worksheet.mergeCells(range);
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully merged cells in ${sheetName} range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error merging cells: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to merge cells: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register unmerge_cells tool
  server.tool(
    'unmerge_cells',
    'Unmerge previously merged cells in an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells to unmerge (e.g., "A1:C10")')
    },
    async ({ fileAbsolutePath, sheetName, range }) => {
      try {
        console.error(`Unmerging cells in ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        // Unmerge the cells in the specified range
        worksheet.unMergeCells(range);
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully unmerged cells in ${sheetName} range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error unmerging cells: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to unmerge cells: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register add_worksheet tool
  server.tool(
    'add_worksheet',
    'Add a new worksheet to an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Name for the new worksheet')
    },
    async ({ fileAbsolutePath, sheetName }) => {
      try {
        console.error(`Adding worksheet ${sheetName} to ${fileAbsolutePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        // Check if worksheet already exists
        if (workbook.getWorksheet(sheetName)) {
          throw new Error(`Worksheet "${sheetName}" already exists`);
        }
        
        // Add new worksheet
        workbook.addWorksheet(sheetName);
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully added worksheet "${sheetName}" to ${fileAbsolutePath}` 
          }]
        };
      } catch (error) {
        console.error(`Error adding worksheet: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to add worksheet: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register apply_styles tool
  server.tool(
    'apply_styles',
    'Apply multiple styles to cells in an Excel file',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells to style (e.g., "A1:C10")'),
      styles: z.object({
        font: z.object({
          name: z.string().optional(),
          size: z.number().optional(),
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.boolean().optional(),
          color: z.string().optional()
        }).optional(),
        fill: z.object({
          type: z.string().optional(),
          pattern: z.string().optional(),
          color: z.string().optional()
        }).optional(),
        border: z.object({
          top: z.object({ style: z.string(), color: z.string() }).optional(),
          bottom: z.object({ style: z.string(), color: z.string() }).optional(),
          left: z.object({ style: z.string(), color: z.string() }).optional(),
          right: z.object({ style: z.string(), color: z.string() }).optional()
        }).optional(),
        alignment: z.object({
          horizontal: z.string().optional(),
          vertical: z.string().optional(),
          wrapText: z.boolean().optional()
        }).optional()
      }).describe('Style options to apply')
    },
    async ({ fileAbsolutePath, sheetName, range, styles }) => {
      try {
        console.error(`Applying styles to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        const { startCol, startRow, endCol, endRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Apply styles to each cell in the range
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startColNum; col <= endColNum; col++) {
            const cellAddress = `${numberToColumnName(col)}${row}`;
            const cell = worksheet.getCell(cellAddress);
            
            // Apply font styles
            if (styles.font) {
              cell.font = cell.font || {};
              if (styles.font.name !== undefined) cell.font.name = styles.font.name;
              if (styles.font.size !== undefined) cell.font.size = styles.font.size;
              if (styles.font.bold !== undefined) cell.font.bold = styles.font.bold;
              if (styles.font.italic !== undefined) cell.font.italic = styles.font.italic;
              if (styles.font.underline !== undefined) cell.font.underline = styles.font.underline;
              if (styles.font.color !== undefined) cell.font.color = { argb: styles.font.color };
            }
            
            // Apply fill styles
            if (styles.fill) {
              cell.fill = {
                type: styles.fill.type || 'pattern',
                pattern: styles.fill.pattern || 'solid',
                fgColor: styles.fill.color ? { argb: styles.fill.color } : undefined
              };
            }
            
            // Apply border styles
            if (styles.border) {
              cell.border = cell.border || {};
              if (styles.border.top) {
                cell.border.top = { 
                  style: styles.border.top.style, 
                  color: { argb: styles.border.top.color } 
                };
              }
              if (styles.border.bottom) {
                cell.border.bottom = { 
                  style: styles.border.bottom.style, 
                  color: { argb: styles.border.bottom.color } 
                };
              }
              if (styles.border.left) {
                cell.border.left = { 
                  style: styles.border.left.style, 
                  color: { argb: styles.border.left.color } 
                };
              }
              if (styles.border.right) {
                cell.border.right = { 
                  style: styles.border.right.style, 
                  color: { argb: styles.border.right.color } 
                };
              }
            }
            
            // Apply alignment styles
            if (styles.alignment) {
              cell.alignment = cell.alignment || {};
              if (styles.alignment.horizontal !== undefined) cell.alignment.horizontal = styles.alignment.horizontal;
              if (styles.alignment.vertical !== undefined) cell.alignment.vertical = styles.alignment.vertical;
              if (styles.alignment.wrapText !== undefined) cell.alignment.wrapText = styles.alignment.wrapText;
            }
          }
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully applied styles to ${sheetName} range ${range}` 
          }]
        };
      } catch (error) {
        console.error(`Error applying styles: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to apply styles: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register refresh_excel_file tool
  server.tool(
    'refresh_excel_file',
    'Close and reopen Excel file to refresh changes using AppleScript',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file to refresh')
    },
    async ({ fileAbsolutePath }) => {
      try {
        console.error(`Refreshing Excel file: ${fileAbsolutePath}`);
        
        // Construct the AppleScript to close and reopen the file WITHOUT saving
        // This avoids potentially overwriting MCP changes with Excel's cached version
        const path = require('path');
        const fileName = path.basename(fileAbsolutePath);
        const appleScript = `
          tell application "Microsoft Excel"
            set fileToRefresh to "${fileAbsolutePath}"
            set fileName to "${fileName}"
            
            -- Check if the file is open
            set isOpen to false
            set targetWorkbook to null
            repeat with i from 1 to count of workbooks
              if name of workbook i is fileName then
                set isOpen to true
                set targetWorkbook to workbook i
                exit repeat
              end if
            end repeat
            
            if isOpen then
              -- Close without saving to load the actual file from disk
              close targetWorkbook saving no
              
              -- Reopen the file
              open fileToRefresh
              
              return "Successfully refreshed " & fileName
            else
              -- File not open, just open it
              open fileToRefresh
              return "Opened " & fileName
            end if
          end tell
        `;
        
        // Execute the AppleScript using osascript
        const { execSync } = require('child_process');
        const result = execSync(`osascript -e '${appleScript.replace(/'/g, "'\''")}' 2>&1`).toString();
        
        console.error(`AppleScript result: ${result}`);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully refreshed Excel file: ${fileName}` 
          }]
        };
      } catch (error) {
        console.error(`Error refreshing Excel file: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to refresh Excel file: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Connect the server to the transport
  console.error('Excel MCP starting...');
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Excel MCP connected to transport');
}

main().catch(error => {
  console.error('Error:', error);
  process.exit(1);
});
