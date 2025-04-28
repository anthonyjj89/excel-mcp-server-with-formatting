import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';

export const formatCells: Tool = {
  name: 'format_cells',
  description: 'Format cells in Excel sheet with colors, fonts, borders, etc.',
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
        description: 'Range of cells to format (e.g., "A1:C10")'
      },
      format: {
        type: 'object',
        description: 'Formatting options',
        properties: {
          bold: { type: 'boolean', description: 'Make text bold' },
          italic: { type: 'boolean', description: 'Make text italic' },
          underline: { type: 'boolean', description: 'Underline text' },
          fontSize: { type: 'number', description: 'Font size' },
          fontName: { type: 'string', description: 'Font name' },
          fontColor: { type: 'string', description: 'Font color (hex code e.g., "#FF0000")' },
          backgroundColor: { type: 'string', description: 'Background color (hex code e.g., "#FFFF00")' },
          horizontalAlignment: { 
            type: 'string', 
            enum: ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'],
            description: 'Horizontal alignment' 
          },
          verticalAlignment: { 
            type: 'string', 
            enum: ['top', 'middle', 'bottom', 'distributed', 'justify'],
            description: 'Vertical alignment' 
          },
          wrapText: { type: 'boolean', description: 'Wrap text' },
          numberFormat: { type: 'string', description: 'Number format (e.g., "0.00", "0%", "m/d/yy")' },
        }
      }
    },
    required: ['fileAbsolutePath', 'sheetName', 'range', 'format']
  },
  handler: async ({ fileAbsolutePath, sheetName, range, format }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`Worksheet "${sheetName}" not found`);
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

      // Apply formatting to the specified range
      for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
        const row = sheet.getRow(rowNum);
        for (let colNum = startColNum; colNum <= endColNum; colNum++) {
          const cell = row.getCell(colNum);
          
          // Apply font formatting
          if (format.bold !== undefined || 
              format.italic !== undefined || 
              format.underline !== undefined || 
              format.fontSize !== undefined ||
              format.fontName !== undefined ||
              format.fontColor !== undefined) {
            
            const fontOptions: Partial<ExcelJS.Font> = {};
            
            if (format.bold !== undefined) fontOptions.bold = format.bold;
            if (format.italic !== undefined) fontOptions.italic = format.italic;
            if (format.underline !== undefined) fontOptions.underline = format.underline;
            if (format.fontSize !== undefined) fontOptions.size = format.fontSize;
            if (format.fontName !== undefined) fontOptions.name = format.fontName;
            if (format.fontColor !== undefined) fontOptions.color = { argb: format.fontColor.replace('#', '') };
            
            cell.font = fontOptions;
          }
          
          // Apply background color
          if (format.backgroundColor !== undefined) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: format.backgroundColor.replace('#', '') }
            };
          }
          
          // Apply alignment
          if (format.horizontalAlignment !== undefined || 
              format.verticalAlignment !== undefined || 
              format.wrapText !== undefined) {
            
            const alignmentOptions: Partial<ExcelJS.Alignment> = {};
            
            if (format.horizontalAlignment !== undefined) alignmentOptions.horizontal = format.horizontalAlignment;
            if (format.verticalAlignment !== undefined) alignmentOptions.vertical = format.verticalAlignment;
            if (format.wrapText !== undefined) alignmentOptions.wrapText = format.wrapText;
            
            cell.alignment = alignmentOptions;
          }
          
          // Apply number format
          if (format.numberFormat !== undefined) {
            cell.numFmt = format.numberFormat;
          }
        }
      }

      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);

      return {
        status: 'success',
        message: `Formatting applied to ${sheetName} in range ${range}`,
        sheetName,
        formattedRange: range
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to format Excel cells: ${error.message}`);
      }
      throw error;
    }
  }
};
