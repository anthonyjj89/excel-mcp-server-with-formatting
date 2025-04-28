import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';
import path from 'path';

export const addWorksheet: Tool = {
  name: 'add_worksheet',
  description: 'Add a new worksheet to an Excel workbook',
  parameters: {
    type: 'object',
    properties: {
      fileAbsolutePath: {
        type: 'string',
        description: 'Absolute path to the Excel file'
      },
      sheetName: {
        type: 'string',
        description: 'Name for the new worksheet'
      },
      tabColor: {
        type: 'string',
        description: 'Tab color (hex code e.g., "#FF0000")'
      },
      columns: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            header: {
              type: 'string',
              description: 'Column header text'
            },
            key: {
              type: 'string',
              description: 'Column key (for referencing in data)'
            },
            width: {
              type: 'number',
              description: 'Column width'
            }
          }
        },
        description: 'Column definitions'
      },
      headerRowFormat: {
        type: 'object',
        description: 'Formatting for the header row (if columns provided)',
        properties: {
          bold: { type: 'boolean' },
          fontSize: { type: 'number' },
          fontColor: { type: 'string' },
          backgroundColor: { type: 'string' }
        }
      },
      position: {
        type: 'number',
        description: 'Position of the new sheet (0-based index)'
      }
    },
    required: ['fileAbsolutePath', 'sheetName']
  },
  handler: async ({ fileAbsolutePath, sheetName, tabColor, columns, headerRowFormat, position }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      // Check if sheet already exists
      if (workbook.getWorksheet(sheetName)) {
        throw new Error(`Worksheet "${sheetName}" already exists`);
      }
      
      // Add worksheet with options for tab color
      const worksheet = workbook.addWorksheet(sheetName, {
        properties: {
          tabColor: tabColor ? { argb: tabColor.replace('#', '') } : undefined
        }
      });
      // Note: Setting sheet position is not directly supported. Sheet will be added at the end.
      
      // Add columns if specified
      if (columns && columns.length > 0) {
        worksheet.columns = columns.map(col => ({
          header: col.header,
          key: col.key,
          width: col.width
        }));
        
        // Apply header row formatting
        if (headerRowFormat) {
          const headerRow = worksheet.getRow(1);
          
          if (headerRowFormat.bold) {
            headerRow.font = { bold: true };
          }
          
          if (headerRowFormat.fontSize) {
            headerRow.font = { ...headerRow.font, size: headerRowFormat.fontSize };
          }
          
          if (headerRowFormat.fontColor) {
            headerRow.font = { 
              ...headerRow.font, 
              color: { argb: headerRowFormat.fontColor.replace('#', '') } 
            };
          }
          
          if (headerRowFormat.backgroundColor) {
            headerRow.eachCell(cell => {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: headerRowFormat.backgroundColor.replace('#', '') }
              };
            });
          }
          
          headerRow.commit();
        }
      }
      
      // Save workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return {
        status: 'success',
        message: `Worksheet "${sheetName}" added to ${fileAbsolutePath}`,
        sheetName,
        workbook: path.basename(fileAbsolutePath)
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to add worksheet: ${error.message}`);
      }
      throw error;
    }
  }
};
