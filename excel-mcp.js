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
    version: "0.1.1"
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
        const workbook = new Ex