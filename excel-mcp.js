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

// Smart deletion system helper functions
function analyzeSheetContent(worksheet) {
  const analysis = {
    totalRows: worksheet.rowCount,
    sections: [],
    contentMap: new Map()
  };

  // Scan all rows for content patterns
  for (let rowNum = 1; rowNum <= worksheet.rowCount; rowNum++) {
    const row = worksheet.getRow(rowNum);
    if (row.hasValues) {
      const rowContent = [];
      row.eachCell((cell, colNumber) => {
        rowContent.push(cell.text || '');
      });
      
      const rowSignature = rowContent.join('|').trim();
      if (rowSignature) {
        analysis.contentMap.set(rowNum, {
          signature: rowSignature,
          content: rowContent,
          isEmpty: false
        });
      }
    } else {
      analysis.contentMap.set(rowNum, {
        signature: '',
        content: [],
        isEmpty: true
      });
    }
  }

  return analysis;
}

function detectDuplicateSections(analysis, options) {
  const duplicateRanges = [];
  const signatureGroups = new Map();
  
  // Group rows by their content signature
  for (const [rowNum, rowData] of analysis.contentMap) {
    if (!rowData.isEmpty && rowData.signature) {
      if (!signatureGroups.has(rowData.signature)) {
        signatureGroups.set(rowData.signature, []);
      }
      signatureGroups.get(rowData.signature).push(rowNum);
    }
  }
  
  // Find duplicate sections (2+ consecutive rows with same signatures)
  for (const [signature, rows] of signatureGroups) {
    if (rows.length > 1) {
      // Group consecutive rows into sections
      const sections = [];
      let currentSection = [rows[0]];
      
      for (let i = 1; i < rows.length; i++) {
        if (rows[i] === rows[i-1] + 1) {
          currentSection.push(rows[i]);
        } else {
          if (currentSection.length > 0) {
            sections.push([...currentSection]);
          }
          currentSection = [rows[i]];
        }
      }
      if (currentSection.length > 0) {
        sections.push(currentSection);
      }
      
      // Find multi-row duplicate sections
      if (sections.length > 1) {
        const sectionGroups = groupConsecutiveSections(sections, analysis);
        
        for (const group of sectionGroups) {
          if (group.length > 1) {
            // Keep first, mark others for deletion
            for (let i = 1; i < group.length; i++) {
              const section = group[i];
              const startRow = Math.min(...section);
              const endRow = Math.max(...section);
              
              duplicateRanges.push({
                range: `A${startRow}:D${endRow}`,
                startRow,
                endRow,
                reason: `Duplicate of section starting at row ${Math.min(...group[0])}`
              });
            }
          }
        }
      }
    }
  }
  
  return duplicateRanges;
}

function groupConsecutiveSections(sections, analysis) {
  const groups = [];
  
  for (const section of sections) {
    // Check if this section is part of a larger duplicate block
    const sectionSignatures = section.map(rowNum => 
      analysis.contentMap.get(rowNum).signature
    );
    
    let matched = false;
    for (const group of groups) {
      const groupSignatures = group[0].map(rowNum => 
        analysis.contentMap.get(rowNum).signature
      );
      
      if (arraysEqual(sectionSignatures, groupSignatures)) {
        group.push(section);
        matched = true;
        break;
      }
    }
    
    if (!matched) {
      groups.push([section]);
    }
  }
  
  return groups;
}

function arraysEqual(a, b) {
  return a.length === b.length && a.every((val, i) => val === b[i]);
}

async function executeSurgicalDeletion(worksheet, duplicateRanges) {
  let cellsCleared = 0;
  const operations = [];
  
  for (const duplicate of duplicateRanges) {
    const { startRow, endRow } = duplicate;
    
    // Clear each row in the duplicate section
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      row.eachCell((cell) => {
        cell.value = null;
        cellsCleared++;
      });
    }
    
    operations.push(`Cleared rows ${startRow}-${endRow}: ${duplicate.reason}`);
  }
  
  return {
    cellsCleared,
    operations,
    rangesProcessed: duplicateRanges.length
  };
}

function formatDeletionReport(analysis, duplicateRanges, deletionReport, isDryRun) {
  let report = `Smart Deletion Analysis Results:\n\n`;
  report += `📊 Sheet Analysis:\n`;
  report += `- Total rows scanned: ${analysis.totalRows}\n`;
  report += `- Content rows found: ${Array.from(analysis.contentMap.values()).filter(r => !r.isEmpty).length}\n\n`;
  
  if (duplicateRanges.length > 0) {
    report += `🔍 Duplicate Sections Detected:\n`;
    duplicateRanges.forEach((dup, i) => {
      report += `${i + 1}. Range ${dup.range}: ${dup.reason}\n`;
    });
    report += `\n`;
    
    if (isDryRun) {
      report += `🔬 DRY RUN MODE - No changes made\n`;
      report += `Would clear ${duplicateRanges.length} duplicate sections\n`;
    } else {
      report += `✅ Surgical Deletion Complete:\n`;
      report += `- Sections processed: ${deletionReport.rangesProcessed}\n`;
      report += `- Total cells cleared: ${deletionReport.cellsCleared}\n`;
      report += `- Operations performed:\n`;
      deletionReport.operations.forEach(op => {
        report += `  • ${op}\n`;
      });
    }
  } else {
    report += `✅ No duplicate sections found - sheet is already clean!\n`;
  }
  
  return report;
}

// Color formatting helper function
function formatExcelColor(color) {
  if (!color) return undefined;
  
  // Remove # prefix if present
  let colorValue = color;
  if (colorValue.startsWith('#')) {
    colorValue = colorValue.replace('#', '');
  }
  
  // Add FF alpha channel prefix for opacity if needed
  return colorValue.length === 6 ? 'FF' + colorValue : colorValue;
}

async function main() {
  // Create the MCP server
  const server = new McpServer({
    name: "Excel MCP",
    version: "0.2.1"
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
        .describe('Data to write to the Excel sheet'),
      autoFit: z.boolean().optional().describe('Automatically adjust column widths to fit content')
    },
    async ({ fileAbsolutePath, sheetName, range, data, autoFit = false }) => {
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
        const { startCol, startRow, endCol, endRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Track content length for auto-fit
        const columnWidths = {};
        
        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          for (let j = 0; j < row.length; j++) {
            const cellAddress = `${numberToColumnName(startColNum + j)}${startRow + i}`;
            const cell = worksheet.getCell(cellAddress);
            
            // Handle possible formula values (strings starting with =)
            if (typeof row[j] === 'string' && row[j].startsWith('=')) {
              cell.value = { formula: row[j].substring(1) };
              // Force Excel to recalculate
              if (cell.model) {
                cell.model.result = undefined;
              }
            } else {
              cell.value = row[j];
            }
            
            // Track content length for auto-fit
            if (autoFit && row[j] !== null && row[j] !== undefined) {
              const colLetter = numberToColumnName(startColNum + j);
              const contentLength = String(row[j]).length;
              
              if (!columnWidths[colLetter] || contentLength > columnWidths[colLetter]) {
                columnWidths[colLetter] = contentLength;
              }
            }
          }
        }
        
        // Apply auto-fit if requested
        if (autoFit) {
          // Apply column widths with padding
          for (let colNum = startColNum; colNum <= endColNum; colNum++) {
            const colLetter = numberToColumnName(colNum);
            if (columnWidths[colLetter]) {
              worksheet.getColumn(colLetter).width = columnWidths[colLetter] + 2; // Add padding
            }
          }
          console.error(`Auto-fitted columns ${startCol}-${endCol}`);
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully wrote data to ${sheetName} in range ${range}${autoFit ? ' and auto-fitted columns' : ''}` 
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
      formulas: z.array(z.array(z.string())).describe('Formulas to write to the Excel sheet (e.g., "=A1+B1")'),
      autoFit: z.boolean().optional().describe('Automatically adjust column widths to fit content')
    },
    async ({ fileAbsolutePath, sheetName, range, formulas, autoFit = false }) => {
      try {
        console.error(`Writing formulas to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
        const workbook = new ExcelJS.Workbook();
        
        // Enable formula calculation on load - FIX for formula issues
        workbook.calcProperties = workbook.calcProperties || {};
        workbook.calcProperties.fullCalcOnLoad = true;
        
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
        const { startCol, startRow, endCol, endRow } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Track content length for auto-fit
        const columnWidths = {};
        
        for (let i = 0; i < formulas.length; i++) {
          const row = formulas[i];
          for (let j = 0; j < row.length; j++) {
            const cellAddress = `${numberToColumnName(startColNum + j)}${startRow + i}`;
            const cell = worksheet.getCell(cellAddress);
            
            if (row[j]) {
              if (row[j].startsWith('=')) {
                cell.value = { formula: row[j].substring(1) };
              } else {
                cell.value = { formula: row[j] };
              }
              
              // Force Excel to recalculate this formula when the file is opened - FIX for formula issues
              if (cell.model) {
                cell.model.result = undefined;
              }
              
              // Track content length for auto-fit
              if (autoFit) {
                const colLetter = numberToColumnName(startColNum + j);
                const textLength = row[j].length;
                
                if (!columnWidths[colLetter] || textLength > columnWidths[colLetter]) {
                  columnWidths[colLetter] = textLength;
                }
              }
            }
          }
        }
        
        // Apply auto-fit if requested
        if (autoFit) {
          // Apply column widths with padding
          for (let colNum = startColNum; colNum <= endColNum; colNum++) {
            const colLetter = numberToColumnName(colNum);
            if (columnWidths[colLetter]) {
              worksheet.getColumn(colLetter).width = columnWidths[colLetter] + 2; // Add padding
            }
          }
          console.error(`Auto-fitted columns ${startCol}-${endCol}`);
        }
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return { 
          content: [{ 
            type: "text", 
            text: `Successfully wrote formulas to ${sheetName} in range ${range}${autoFit ? ' and auto-fitted columns' : ''}` 
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

  // Add a dedicated tool for adjusting column widths
  server.tool(
    'autofit_columns',
    'Automatically adjust column widths based on content',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      columns: z.array(z.string()).optional().describe('Columns to auto-fit (e.g., ["A", "B", "C"]). Default: all columns'),
      padding: z.number().optional().describe('Additional padding to add (in characters). Default: 2'),
      minWidth: z.number().optional().describe('Minimum column width'),
      maxWidth: z.number().optional().describe('Maximum column width')
    },
    async ({ fileAbsolutePath, sheetName, columns, padding = 2, minWidth, maxWidth }) => {
      try {
        console.error(`Auto-fitting columns in ${fileAbsolutePath}, sheet: ${sheetName}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        // Store max length for each column
        const columnWidths = {};
        
        // Process all rows to find the max content width for each column
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const column = worksheet.getColumn(colNumber);
            const columnLetter = column.letter;
            
            // Skip if we're only processing specific columns and this isn't one of them
            if (columns && !columns.includes(columnLetter)) {
              return;
            }
            
            let contentLength = 0;
            
            if (cell.text) {
              contentLength = cell.text.toString().length;
            } else if (cell.value !== null && cell.value !== undefined) {
              contentLength = cell.value.toString().length;
            }
            
            // Account for header/column name length too
            const headerLength = column.header ? column.header.toString().length : 0;
            
            // Update max width if needed
            if (!columnWidths[columnLetter] || contentLength > columnWidths[columnLetter]) {
              columnWidths[columnLetter] = Math.max(contentLength, headerLength);
            }
          });
        });
        
        // Set column widths based on content length plus padding
        Object.keys(columnWidths).forEach(columnLetter => {
          const column = worksheet.getColumn(columnLetter);
          let width = columnWidths[columnLetter] + padding;
          
          // Apply min/max constraints if provided
          if (minWidth !== undefined && width < minWidth) {
            width = minWidth;
          }
          
          if (maxWidth !== undefined && width > maxWidth) {
            width = maxWidth;
          }
          
          column.width = width;
        });
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        return {
          content: [{
            type: "text",
            text: `Successfully auto-fit columns in ${sheetName}. Columns adjusted: ${Object.keys(columnWidths).join(', ')}`
          }]
        };
      } catch (error) {
        console.error(`Error auto-fitting columns: ${error.message}`);
        return {
          content: [{
            type: "text",
            text: `Failed to auto-fit columns: ${error.message}`
          }],
          isError: true
        };
      }
    }
  );

  // Add borders to cells in an Excel file
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
          border.top = { 
            style: borderStyle.top.style, 
            color: { argb: formatExcelColor(borderStyle.top.color) }
          };
        }
        if (borderStyle.bottom) {
          border.bottom = { 
            style: borderStyle.bottom.style, 
            color: { argb: formatExcelColor(borderStyle.bottom.color) }
          };
        }
        if (borderStyle.left) {
          border.left = {
            style: borderStyle.left.style, 
            color: { argb: formatExcelColor(borderStyle.left.color) }
          };
        }
        if (borderStyle.right) {
          border.right = { 
            style: borderStyle.right.style, 
            color: { argb: formatExcelColor(borderStyle.right.color) }
          };
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

  // Format cells in an Excel file
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
            if (formatting.fontColor !== undefined) {
              cell.font.color = { argb: formatExcelColor(formatting.fontColor) };
            }
            
            // Fill formatting with proper color handling
            if (formatting.fillColor !== undefined) {
              const formattedColor = formatExcelColor(formatting.fillColor);
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: formattedColor },
                bgColor: { argb: formattedColor }
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

  // Merge cells in an Excel file
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

  // Unmerge cells in an Excel file
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

  // Add a new worksheet to an Excel file
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

  // Apply multiple styles to cells in an Excel file
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
              if (styles.font.color !== undefined) {
                cell.font.color = { argb: formatExcelColor(styles.font.color) };
              }
            }
            
            // Apply fill styles
            if (styles.fill && styles.fill.color) {
              const formattedColor = formatExcelColor(styles.fill.color);
              cell.fill = {
                type: styles.fill.type || 'pattern',
                pattern: styles.fill.pattern || 'solid',
                fgColor: { argb: formattedColor },
                bgColor: { argb: formattedColor }
              };
            }
            
            // Apply border styles
            if (styles.border) {
              cell.border = cell.border || {};
              if (styles.border.top) {
                cell.border.top = { 
                  style: styles.border.top.style, 
                  color: { argb: formatExcelColor(styles.border.top.color) } 
                };
              }
              if (styles.border.bottom) {
                cell.border.bottom = { 
                  style: styles.border.bottom.style, 
                  color: { argb: formatExcelColor(styles.border.bottom.color) } 
                };
              }
              if (styles.border.left) {
                cell.border.left = { 
                  style: styles.border.left.style, 
                  color: { argb: formatExcelColor(styles.border.left.color) } 
                };
              }
              if (styles.border.right) {
                cell.border.right = { 
                  style: styles.border.right.style, 
                  color: { argb: formatExcelColor(styles.border.right.color) } 
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

  // Refresh Excel file
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

  // Register the delete_rows tool
  server.tool(
    'delete_rows',
    'Delete rows from an Excel worksheet',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      startRow: z.number().describe('Starting row number (1-based index)'),
      deleteCount: z.number().optional().describe('Number of rows to delete (default: 1)')
    },
    async ({ fileAbsolutePath, sheetName, startRow, deleteCount = 1 }) => {
      try {
        console.error(`Deleting ${deleteCount} rows starting from row ${startRow} in ${fileAbsolutePath}, sheet: ${sheetName}`);
        
        if (startRow < 1) {
          throw new Error('Row number must be 1 or greater');
        }
        
        if (deleteCount < 1) {
          throw new Error('Delete count must be 1 or greater');
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        // Check if we're trying to delete beyond existing rows
        const actualRowCount = worksheet.actualRowCount;
        if (startRow > actualRowCount) {
          throw new Error(`Cannot delete row ${startRow}: sheet only has ${actualRowCount} rows`);
        }
        
        // Adjust delete count if it would exceed available rows
        const adjustedDeleteCount = Math.min(deleteCount, actualRowCount - startRow + 1);
        
        // Use spliceRows to delete rows (convert 1-based to 0-based indexing)
        worksheet.spliceRows(startRow - 1, adjustedDeleteCount);
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        const message = adjustedDeleteCount < deleteCount 
          ? `Successfully deleted ${adjustedDeleteCount} rows (adjusted from ${deleteCount} to avoid exceeding sheet bounds)` 
          : `Successfully deleted ${deleteCount} rows starting from row ${startRow}`;
        
        console.error(message);
        
        return { 
          content: [{ 
            type: "text", 
            text: message
          }]
        };
      } catch (error) {
        console.error(`Error deleting rows: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to delete rows: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register the delete_columns tool
  server.tool(
    'delete_columns',
    'Delete columns from an Excel worksheet',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      startColumn: z.string().describe('Starting column letter (e.g., "A", "B", "AA")'),
      deleteCount: z.number().optional().describe('Number of columns to delete (default: 1)')
    },
    async ({ fileAbsolutePath, sheetName, startColumn, deleteCount = 1 }) => {
      try {
        console.error(`Deleting ${deleteCount} columns starting from column ${startColumn} in ${fileAbsolutePath}, sheet: ${sheetName}`);
        
        if (deleteCount < 1) {
          throw new Error('Delete count must be 1 or greater');
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        // Convert column letter to number
        const startColNum = columnNameToNumber(startColumn.toUpperCase());
        
        // Check if we're trying to delete beyond existing columns
        const actualColumnCount = worksheet.actualColumnCount;
        if (startColNum > actualColumnCount) {
          throw new Error(`Cannot delete column ${startColumn}: sheet only has ${actualColumnCount} columns`);
        }
        
        // Adjust delete count if it would exceed available columns
        const adjustedDeleteCount = Math.min(deleteCount, actualColumnCount - startColNum + 1);
        
        // Use spliceColumns to delete columns
        worksheet.spliceColumns(startColNum, adjustedDeleteCount);
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        const endColumn = numberToColumnName(startColNum + adjustedDeleteCount - 1);
        const message = adjustedDeleteCount < deleteCount 
          ? `Successfully deleted ${adjustedDeleteCount} columns (adjusted from ${deleteCount} to avoid exceeding sheet bounds)` 
          : `Successfully deleted ${deleteCount} columns from ${startColumn} to ${endColumn}`;
        
        console.error(message);
        
        return { 
          content: [{ 
            type: "text", 
            text: message
          }]
        };
      } catch (error) {
        console.error(`Error deleting columns: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to delete columns: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register the clear_range tool with smart duplicate detection
  server.tool(
    'clear_range',
    'Clear cell values and optionally formatting from a range in Excel worksheet. Auto-detects and clears complete duplicate sections when "smart" range is used.',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells to clear (e.g., "A1:C10") or "smart" to auto-detect duplicates'),
      clearFormatting: z.boolean().optional().describe('Also clear formatting (default: false)')
    },
    async ({ fileAbsolutePath, sheetName, range, clearFormatting = false }) => {
      try {
        console.error(`Clearing range ${range} in ${fileAbsolutePath}, sheet: ${sheetName}, including formatting: ${clearFormatting}`);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        let clearedCells = 0;
        let message = '';
        
        // Smart duplicate detection mode
        if (range.toLowerCase() === 'smart') {
          console.error(`Smart duplicate detection mode activated`);
          
          // Phase 1: Analyze sheet content
          const analysis = analyzeSheetContent(worksheet);
          
          // Phase 2: Detect duplicates
          const duplicateRanges = detectDuplicateSections(analysis, {});
          
          if (duplicateRanges.length === 0) {
            message = `Smart analysis: No duplicate sections found in sheet "${sheetName}"`;
          } else {
            // Phase 3: Clear all detected duplicates
            const operations = [];
            
            for (const duplicate of duplicateRanges) {
              const { startRow, endRow } = duplicate;
              
              for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
                const row = worksheet.getRow(rowNum);
                row.eachCell((cell) => {
                  cell.value = null;
                  if (clearFormatting) {
                    cell.style = {};
                  }
                  clearedCells++;
                });
              }
              
              operations.push(`Cleared duplicate section rows ${startRow}-${endRow}`);
            }
            
            await workbook.xlsx.writeFile(fileAbsolutePath);
            
            message = `Smart deletion complete:\n`;
            message += `• Duplicates detected: ${duplicateRanges.length} sections\n`;
            message += `• Cells cleared: ${clearedCells}\n`;
            message += `• Operations:\n`;
            operations.forEach(op => {
              message += `  - ${op}\n`;
            });
          }
        } else {
          // Traditional range clearing
          const { startCol, startRow, endCol, endRow } = parseRange(range);
          const startColNum = columnNameToNumber(startCol);
          const endColNum = columnNameToNumber(endCol);
          
          for (let row = startRow; row <= endRow; row++) {
            for (let col = startColNum; col <= endColNum; col++) {
              const cellAddress = `${numberToColumnName(col)}${row}`;
              const cell = worksheet.getCell(cellAddress);
              
              // Clear value
              cell.value = null;
              
              // Clear formatting if requested
              if (clearFormatting) {
                cell.style = {};
              }
              
              clearedCells++;
            }
          }
          
          await workbook.xlsx.writeFile(fileAbsolutePath);
          
          message = clearFormatting 
            ? `Successfully cleared ${clearedCells} cells (values and formatting) in range ${range}`
            : `Successfully cleared ${clearedCells} cell values in range ${range}`;
        }
        
        console.error(message);
        
        return { 
          content: [{ 
            type: "text", 
            text: message
          }]
        };
      } catch (error) {
        console.error(`Error clearing range: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to clear range: ${error.message}` 
          }],
          isError: true
        };
      }
    }
  );

  // Register the delete_worksheet tool
  server.tool(
    'delete_worksheet',
    'Delete an entire worksheet from an Excel workbook',
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Name of the sheet to delete')
    },
    async ({ fileAbsolutePath, sheetName }) => {
      try {
        console.error(`Deleting worksheet "${sheetName}" from ${fileAbsolutePath}`);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const worksheet = workbook.getWorksheet(sheetName);
        if (!worksheet) {
          throw new Error(`Sheet "${sheetName}" not found`);
        }
        
        // Check if this is the only worksheet
        if (workbook.worksheets.length === 1) {
          throw new Error('Cannot delete the only worksheet in the workbook');
        }
        
        // Remove the worksheet
        workbook.removeWorksheet(worksheet.id);
        
        // Save the workbook
        await workbook.xlsx.writeFile(fileAbsolutePath);
        
        const message = `Successfully deleted worksheet "${sheetName}"`;
        console.error(message);
        
        return { 
          content: [{ 
            type: "text", 
            text: message
          }]
        };
      } catch (error) {
        console.error(`Error deleting worksheet: ${error.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Failed to delete worksheet: ${error.message}` 
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
