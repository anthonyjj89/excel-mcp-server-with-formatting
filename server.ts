import { Server as McpServer } from '@modelcontextprotocol/sdk/dist/cjs/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/dist/cjs/server/stdio.js';
import ExcelJS from 'exceljs';
import * as z from 'zod';

// Create utility functions since ExcelJS.utils doesn't exist
function columnNameToNumber(name: string): number {
  let result = 0;
  for (let i = 0; i < name.length; i++) {
    result = result * 26 + (name.charCodeAt(i) - 64);
  }
  return result;
}

function numberToColumnName(num: number): string {
  let result = '';
  while (num > 0) {
    const modulo = (num - 1) % 26;
    result = String.fromCharCode(65 + modulo) + result;
    num = Math.floor((num - modulo) / 26);
  }
  return result;
}

function parseRange(range: string): { startRow: number; startCol: string; endRow: number; endCol: string } {
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
  const server = new McpServer({
    name: "Excel MCP",
    version: "1.0.0"
  });

  // Register the read_sheet_names tool
  server.tool(
    "read_sheet_names",
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file')
    },
    async ({ fileAbsolutePath }) => {
      try {
        console.error(`Reading sheet names from ${fileAbsolutePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fileAbsolutePath);
        
        const sheetNames = workbook.worksheets.map(sheet => sheet.name);
        
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
    "read_sheet_data",
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().optional().describe('Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]'),
      knownPagingRanges: z.array(z.string()).optional().describe('List of already read paging ranges')
    },
    async ({ fileAbsolutePath, sheetName, range, knownPagingRanges }) => {
      try {
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
    "read_sheet_formula",
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().optional().describe('Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]'),
      knownPagingRanges: z.array(z.string()).optional().describe('List of already read paging ranges')
    },
    async ({ fileAbsolutePath, sheetName, range, knownPagingRanges }) => {
      try {
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
    "write_sheet_data",
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
      data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe('Data to write to the Excel sheet')
    },
    async ({ fileAbsolutePath, sheetName, range, data }) => {
      try {
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
    "write_sheet_formula",
    {
      fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
      sheetName: z.string().describe('Sheet name in the Excel file'),
      range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
      formulas: z.array(z.array(z.string())).describe('Formulas to write to the Excel sheet (e.g., "=A1+B1")')
    },
    async ({ fileAbsolutePath, sheetName, range, formulas }) => {
      try {
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

  // Connect the server
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(error => {
  console.error('Error:', error);
  process.exit(1);
});
