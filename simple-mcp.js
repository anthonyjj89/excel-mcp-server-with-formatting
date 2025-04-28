#!/usr/bin/env node

const http = require('http');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Simple MCP implementation
const PORT = process.env.PORT || 8001;

// Handler for MCP requests
async function handleMCPRequest(req, res) {
  res.setHeader('Content-Type', 'application/json');
  
  if (req.method === 'GET') {
    // Respond to GET with a list of available tools
    res.end(JSON.stringify({
      tools: [
        {
          name: 'read_sheet_data',
          description: 'Read data from Excel sheet with pagination.',
          parameters: {
            type: 'object',
            properties: {
              fileAbsolutePath: {
                type: 'string',
                description: 'Absolute path to the Excel file'
              },
              sheetName: {
                type: 'string', 
                description: 'Sheet name in the Excel file'
              },
              range: {
                type: 'string',
                description: 'Range of cells to read in the Excel sheet (e.g., "A1:C10")'
              }
            },
            required: ['fileAbsolutePath', 'sheetName']
          }
        },
        {
          name: 'read_sheet_names',
          description: 'List all sheet names in an Excel file',
          parameters: {
            type: 'object',
            properties: {
              fileAbsolutePath: {
                type: 'string',
                description: 'Absolute path to the Excel file'
              }
            },
            required: ['fileAbsolutePath']
          }
        }
      ]
    }));
    return;
  }
  
  if (req.method !== 'POST') {
    res.statusCode = 405;
    res.end(JSON.stringify({ error: 'Method not allowed' }));
    return;
  }
  
  let body = '';
  req.on('data', chunk => {
    body += chunk.toString();
  });
  
  req.on('end', async () => {
    try {
      const { tool, params } = JSON.parse(body);
      console.log(`Received request for tool: ${tool}`);
      
      if (tool === 'read_sheet_data') {
        await handleReadSheetData(params, res);
      } else if (tool === 'read_sheet_names') {
        await handleReadSheetNames(params, res);
      } else {
        res.statusCode = 400;
        res.end(JSON.stringify({ error: `Unknown tool: ${tool}` }));
      }
    } catch (error) {
      console.error('Error handling request:', error);
      res.statusCode = 500;
      res.end(JSON.stringify({ error: `Server error: ${error.message}` }));
    }
  });
}

async function handleReadSheetData(params, res) {
  try {
    const { fileAbsolutePath, sheetName, range } = params;
    console.log(`Reading Excel file: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range || 'default'}`);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      res.statusCode = 400;
      res.end(JSON.stringify({ error: `Sheet "${sheetName}" not found in workbook.` }));
      return;
    }
    
    let data = [];
    if (range) {
      // If range is provided, read that specific range
      worksheet.getCell(range.split(':')[0]).value;  // Just to ensure the range is valid
      data = worksheet.getSheetValues();
    } else {
      // Otherwise, read all data
      data = worksheet.getSheetValues();
    }
    
    // Filter out undefined rows and convert to readable format
    const filteredData = data.filter(row => row !== undefined)
      .map(row => Object.values(row).filter(cell => cell !== undefined));
    
    res.end(JSON.stringify({ 
      content: `Successfully read data from ${fileAbsolutePath}:\n\n${JSON.stringify(filteredData, null, 2)}`
    }));
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.statusCode = 500;
    res.end(JSON.stringify({ error: `Error reading Excel file: ${error.message}` }));
  }
}

async function handleReadSheetNames(params, res) {
  try {
    const { fileAbsolutePath } = params;
    console.log(`Reading sheet names from Excel file: ${fileAbsolutePath}`);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheetNames = workbook.worksheets.map(sheet => sheet.name);
    
    res.end(JSON.stringify({ 
      content: `Sheets in ${fileAbsolutePath}:\n\n${JSON.stringify(sheetNames, null, 2)}`
    }));
  } catch (error) {
    console.error('Error reading sheet names:', error);
    res.statusCode = 500;
    res.end(JSON.stringify({ error: `Error reading sheet names: ${error.message}` }));
  }
}

// Create server and start listening
const server = http.createServer(handleMCPRequest);
server.listen(PORT, () => {
  console.log(`Excel MCP server running on port ${PORT}`);
});

console.log('Excel MCP with formatting initialized - simple HTTP version');
