import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';

export const readSheetNames: Tool = {
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
  },
  handler: async ({ fileAbsolutePath }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      // Get all sheet names
      const sheetNames = workbook.worksheets.map(sheet => sheet.name);

      return sheetNames;
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to read sheet names: ${error.message}`);
      }
      throw error;
    }
  }
};
