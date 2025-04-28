import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';

export const writeSheetData: Tool = {
  name: 'write_sheet_data',
  description: 'Write data to the Excel sheet',
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
        description: 'Range of cells in the Excel sheet (e.g., "A1:C10")'
      },
      data: {
        type: 'array',
        items: {
          type: 'array',
          items: {
            type: ['string', 'number', 'boolean', 'null']
          }
        },
        description: 'Data to write to the Excel sheet'
      }
    },
    required: ['fileAbsolutePath', 'sheetName', 'range', 'data']
  },
  handler: async ({ fileAbsolutePath, sheetName, range, data }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      
      try {
        await workbook.xlsx.readFile(fileAbsolutePath);
      } catch (err) {
        // File doesn't exist yet, create new workbook
        workbook.addWorksheet(sheetName);
      }
      
      let sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        sheet = workbook.addWorksheet(sheetName);
      }

      // Parse range
      const rangeMatch = range.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/);
      if (!rangeMatch) {
        throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const startRowNum = parseInt(startRow);

      // Write data to the specified range
      data.forEach((rowData, rowIndex) => {
        const row = sheet.getRow(startRowNum + rowIndex);
        rowData.forEach((cellValue, colIndex) => {
          const col = startColNum + colIndex;
          row.getCell(col).value = cellValue;
        });
        row.commit();
      });

      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);

      return {
        status: 'success',
        message: `Data written to ${sheetName} in range ${range}`,
        sheetName,
        writtenRange: range
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to write Excel data: ${error.message}`);
      }
      throw error;
    }
  }
};
