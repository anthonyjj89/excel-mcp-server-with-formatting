import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';
import path from 'path';

export const readSheetData: Tool = {
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
        description: 'Range of cells to read in the Excel sheet (e.g., "A1:C10"). [default: first paging range]'
      },
      knownPagingRanges: {
        type: 'array',
        items: {
          type: 'string'
        },
        description: 'List of already read paging ranges'
      }
    },
    required: ['fileAbsolutePath', 'sheetName']
  },
  handler: async ({ fileAbsolutePath, sheetName, range, knownPagingRanges }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`Worksheet "${sheetName}" not found`);
      }

      // If no range is specified, read all data
      if (!range) {
        const data: any[][] = [];
        sheet.eachRow((row, rowNumber) => {
          const rowData: any[] = [];
          row.eachCell((cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
          });
          data[rowNumber - 1] = rowData;
        });
        
        return {
          data,
          sheetName,
          readRange: `A1:${sheet.lastColumn?.letter}${sheet.rowCount}`,
          allRangesRead: true
        };
      }

      // Parse range
      const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

      // Read data from the specified range
      const data: any[][] = [];
      for (let i = startRowNum; i <= endRowNum; i++) {
        const rowData: any[] = [];
        for (let j = startColNum; j <= endColNum; j++) {
          const cell = sheet.getCell(i, j);
          rowData.push(cell.value);
        }
        data.push(rowData);
      }

      const allRangesRead = !knownPagingRanges || knownPagingRanges.length === 0 || 
                           (knownPagingRanges.includes(range) && 
                            endRowNum >= sheet.rowCount && 
                            endColNum >= sheet.columnCount);

      return {
        data,
        sheetName,
        readRange: range,
        allRangesRead
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to read Excel data: ${error.message}`);
      }
      throw error;
    }
  }
};
