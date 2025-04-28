import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';

export const addBorders: Tool = {
  name: 'add_borders',
  description: 'Add borders to cells in Excel sheet',
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
        description: 'Range of cells to add borders to (e.g., "A1:C10")'
      },
      borderStyle: {
        type: 'object',
        description: 'Border style options',
        properties: {
          top: {
            type: 'object',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double'],
                description: 'Border style'
              },
              color: {
                type: 'string',
                description: 'Border color (hex code e.g., "#000000")'
              }
            }
          },
          bottom: {
            type: 'object',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double'],
                description: 'Border style'
              },
              color: {
                type: 'string',
                description: 'Border color (hex code e.g., "#000000")'
              }
            }
          },
          left: {
            type: 'object',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double'],
                description: 'Border style'
              },
              color: {
                type: 'string',
                description: 'Border color (hex code e.g., "#000000")'
              }
            }
          },
          right: {
            type: 'object',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double'],
                description: 'Border style'
              },
              color: {
                type: 'string',
                description: 'Border color (hex code e.g., "#000000")'
              }
            }
          },
          outline: {
            type: 'boolean',
            description: 'Add outline border to the entire range'
          },
          all: {
            type: 'object',
            description: 'Apply to all borders (top, bottom, left, right)',
            properties: {
              style: {
                type: 'string',
                enum: ['thin', 'medium', 'thick', 'dotted', 'dashed', 'double'],
                description: 'Border style'
              },
              color: {
                type: 'string',
                description: 'Border color (hex code e.g., "#000000")'
              }
            }
          }
        }
      }
    },
    required: ['fileAbsolutePath', 'sheetName', 'range', 'borderStyle']
  },
  handler: async ({ fileAbsolutePath, sheetName, range, borderStyle }) => {
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

      // Convert border styles to ExcelJS format
      const convertBorderStyle = (style: any) => {
        if (!style) return undefined;
        
        return {
          style: style.style,
          color: style.color ? { argb: style.color.replace('#', '') } : undefined
        };
      };

      // Apply "all" border style to all sides if specified
      const allBorderStyle = borderStyle.all ? convertBorderStyle(borderStyle.all) : undefined;

      // Apply borders to cells
      for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
        const row = sheet.getRow(rowNum);
        for (let colNum = startColNum; colNum <= endColNum; colNum++) {
          const cell = row.getCell(colNum);
          
          const cellBorder: any = {};
          
          // Apply specific border sides
          if (borderStyle.top || allBorderStyle) {
            cellBorder.top = borderStyle.top ? convertBorderStyle(borderStyle.top) : allBorderStyle;
          }
          
          if (borderStyle.bottom || allBorderStyle) {
            cellBorder.bottom = borderStyle.bottom ? convertBorderStyle(borderStyle.bottom) : allBorderStyle;
          }
          
          if (borderStyle.left || allBorderStyle) {
            cellBorder.left = borderStyle.left ? convertBorderStyle(borderStyle.left) : allBorderStyle;
          }
          
          if (borderStyle.right || allBorderStyle) {
            cellBorder.right = borderStyle.right ? convertBorderStyle(borderStyle.right) : allBorderStyle;
          }
          
          // Apply outline borders if specified
          if (borderStyle.outline) {
            // Top row
            if (rowNum === startRowNum) {
              cellBorder.top = allBorderStyle || convertBorderStyle({ 
                style: 'medium', 
                color: borderStyle.top?.color || '#000000' 
              });
            }
            
            // Bottom row
            if (rowNum === endRowNum) {
              cellBorder.bottom = allBorderStyle || convertBorderStyle({ 
                style: 'medium', 
                color: borderStyle.bottom?.color || '#000000' 
              });
            }
            
            // Left column
            if (colNum === startColNum) {
              cellBorder.left = allBorderStyle || convertBorderStyle({ 
                style: 'medium', 
                color: borderStyle.left?.color || '#000000' 
              });
            }
            
            // Right column
            if (colNum === endColNum) {
              cellBorder.right = allBorderStyle || convertBorderStyle({ 
                style: 'medium', 
                color: borderStyle.right?.color || '#000000' 
              });
            }
          }
          
          cell.border = cellBorder;
        }
      }

      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);

      return {
        status: 'success',
        message: `Borders added to ${sheetName} in range ${range}`,
        sheetName,
        borderedRange: range
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to add borders: ${error.message}`);
      }
      throw error;
    }
  }
};
