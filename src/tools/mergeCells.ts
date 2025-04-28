import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';

export const mergeCells: Tool = {
  name: 'merge_cells',
  description: 'Merge cells in Excel sheet',
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
        description: 'Range of cells to merge (e.g., "A1:C10")'
      },
      alignContent: {
        type: 'object',
        description: 'Alignment options for merged content',
        properties: {
          horizontal: {
            type: 'string',
            enum: ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'],
            description: 'Horizontal alignment'
          },
          vertical: {
            type: 'string',
            enum: ['top', 'middle', 'bottom', 'distributed', 'justify'],
            description: 'Vertical alignment'
          }
        }
      }
    },
    required: ['fileAbsolutePath', 'sheetName', 'range']
  },
  handler: async ({ fileAbsolutePath, sheetName, range, alignContent }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`Worksheet "${sheetName}" not found`);
      }

      // Merge cells
      sheet.mergeCells(range);
      
      // Get the top-left cell of the merged range for alignment
      if (alignContent) {
        const rangeMatch = range.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/);
        if (rangeMatch) {
          const [_, startCol, startRow] = rangeMatch;
          const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
          const startRowNum = parseInt(startRow);
          
          const cell = sheet.getCell(startRowNum, startColNum);
          
          // Apply alignment
          cell.alignment = {
            horizontal: alignContent.horizontal,
            vertical: alignContent.vertical
          };
        }
      }

      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);

      return {
        status: 'success',
        message: `Cells merged in ${sheetName} at range ${range}`,
        sheetName,
        mergedRange: range
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to merge cells: ${error.message}`);
      }
      throw error;
    }
  }
};

export const unmergeCells: Tool = {
  name: 'unmerge_cells',
  description: 'Unmerge previously merged cells in Excel sheet',
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
        description: 'Range of cells to unmerge (e.g., "A1:C10")'
      }
    },
    required: ['fileAbsolutePath', 'sheetName', 'range']
  },
  handler: async ({ fileAbsolutePath, sheetName, range }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`Worksheet "${sheetName}" not found`);
      }

      // Unmerge cells
      sheet.unMergeCells(range);

      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);

      return {
        status: 'success',
        message: `Cells unmerged in ${sheetName} at range ${range}`,
        sheetName,
        unmergedRange: range
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to unmerge cells: ${error.message}`);
      }
      throw error;
    }
  }
};
