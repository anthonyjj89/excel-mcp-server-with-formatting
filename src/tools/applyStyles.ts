import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';
import path from 'path';

export const applyStyles: Tool = {
  name: 'apply_styles',
  description: 'Apply cell styles including conditional formatting to Excel',
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
      actions: {
        type: 'array',
        description: 'List of styling actions to apply',
        items: {
          type: 'object',
          properties: {
            type: {
              type: 'string',
              enum: [
                'alternating_rows', 
                'table_style', 
                'header_row', 
                'total_row',
                'banded_columns',
                'highlight_negative',
                'highlight_positive',
                'highlight_max',
                'highlight_min',
                'data_bars',
                'gradient_scale',
                'icon_set'
              ],
              description: 'Type of styling to apply'
            },
            range: {
              type: 'string',
              description: 'Range to apply the styling to (e.g., "A1:C10")'
            },
            properties: {
              type: 'object',
              description: 'Style properties specific to the selected style type'
            }
          },
          required: ['type', 'range']
        }
      }
    },
    required: ['fileAbsolutePath', 'sheetName', 'actions']
  },
  handler: async ({ fileAbsolutePath, sheetName, actions }) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        throw new Error(`Worksheet "${sheetName}" not found`);
      }

      const results = [];

      // Process each styling action
      for (const action of actions) {
        switch (action.type) {
          case 'alternating_rows':
            results.push(applyAlternatingRows(sheet, action.range, action.properties));
            break;
          case 'table_style':
            results.push(applyTableStyle(sheet, action.range, action.properties));
            break;
          case 'header_row':
            results.push(applyHeaderRow(sheet, action.range, action.properties));
            break;
          case 'total_row':
            results.push(applyTotalRow(sheet, action.range, action.properties));
            break;
          case 'highlight_negative':
            results.push(applyHighlightNegative(sheet, action.range, action.properties));
            break;
          case 'highlight_positive':
            results.push(applyHighlightPositive(sheet, action.range, action.properties));
            break;
          case 'highlight_max':
            results.push(applyHighlightMax(sheet, action.range, action.properties));
            break;
          case 'highlight_min':
            results.push(applyHighlightMin(sheet, action.range, action.properties));
            break;
          case 'data_bars':
            results.push(applyDataBars(sheet, action.range, action.properties));
            break;
          default:
            results.push({
              type: action.type,
              status: 'error',
              message: `Unsupported style type: ${action.type}`
            });
        }
      }

      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);

      return {
        status: 'success',
        message: `Styles applied to ${sheetName}`,
        results
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to apply styles: ${error.message}`);
      }
      throw error;
    }
  }
};

// Helper functions for different style types

function applyAlternatingRows(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    const evenColor = properties?.evenColor || '#F8F8F8';
    const oddColor = properties?.oddColor || '#FFFFFF';

    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      const isEvenRow = (rowNum - startRowNum) % 2 === 0;
      const bgcolor = isEvenRow ? evenColor : oddColor;

      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgcolor.replace('#', '') }
        };
      }
    }

    return {
      type: 'alternating_rows',
      status: 'success',
      range
    };
  } catch (error) {
    return {
      type: 'alternating_rows',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyTableStyle(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    // Header row
    if (properties?.headerRow !== false) {
      const headerRow = sheet.getRow(startRowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = headerRow.getCell(colNum);
        cell.font = { bold: true };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: (properties?.headerColor || '#D0D0D0').replace('#', '') }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }

    // Data rows
    for (let rowNum = startRowNum + (properties?.headerRow !== false ? 1 : 0); rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      const isEvenRow = (rowNum - startRowNum) % 2 === 0;
      const bgcolor = properties?.stripedRows && isEvenRow ? 
        (properties?.evenColor || '#F8F8F8') : 
        (properties?.oddColor || '#FFFFFF');

      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        
        // Apply fill
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgcolor.replace('#', '') }
        };
        
        // Apply borders
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }

    return {
      type: 'table_style',
      status: 'success',
      range
    };
  } catch (error) {
    return {
      type: 'table_style',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyHeaderRow(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const rowNum = parseInt(startRow);

    const row = sheet.getRow(rowNum);
    const color = properties?.backgroundColor || '#D0D0D0';
    const textColor = properties?.textColor || '#000000';
    const fontSize = properties?.fontSize || 12;
    const bold = properties?.bold !== false;

    for (let colNum = startColNum; colNum <= endColNum; colNum++) {
      const cell = row.getCell(colNum);
      
      // Apply font
      cell.font = { 
        bold, 
        color: { argb: textColor.replace('#', '') },
        size: fontSize
      };
      
      // Apply fill
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color.replace('#', '') }
      };
      
      // Apply alignment
      cell.alignment = {
        horizontal: properties?.alignment || 'center',
        vertical: 'middle'
      };
      
      // Apply borders
      if (properties?.borders !== false) {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    }

    return {
      type: 'header_row',
      status: 'success',
      range
    };
  } catch (error) {
    return {
      type: 'header_row',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyTotalRow(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, , endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const rowNum = parseInt(endRow);

    const row = sheet.getRow(rowNum);
    const color = properties?.backgroundColor || '#E0E0E0';
    const textColor = properties?.textColor || '#000000';
    const fontSize = properties?.fontSize || 12;
    const bold = properties?.bold !== false;

    for (let colNum = startColNum; colNum <= endColNum; colNum++) {
      const cell = row.getCell(colNum);
      
      // Apply font
      cell.font = { 
        bold, 
        color: { argb: textColor.replace('#', '') },
        size: fontSize
      };
      
      // Apply fill
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color.replace('#', '') }
      };
      
      // Apply borders
      if (properties?.borders !== false) {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
      
      // Apply number format if it's a number
      if (typeof cell.value === 'number' && properties?.numberFormat) {
        cell.numFmt = properties.numberFormat;
      }
    }

    return {
      type: 'total_row',
      status: 'success',
      range
    };
  } catch (error) {
    return {
      type: 'total_row',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyHighlightNegative(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    const color = properties?.color || '#FFC7CE';
    const textColor = properties?.textColor || '#9C0006';

    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        // Apply conditional formatting for negative numbers
        if (typeof value === 'number' && value < 0) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color.replace('#', '') }
          };
          
          cell.font = {
            ...cell.font,
            color: { argb: textColor.replace('#', '') }
          };
        }
      }
    }

    return {
      type: 'highlight_negative',
      status: 'success',
      range
    };
  } catch (error) {
    return {
      type: 'highlight_negative',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyHighlightPositive(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    const color = properties?.color || '#C6EFCE';
    const textColor = properties?.textColor || '#006100';

    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        // Apply conditional formatting for positive numbers
        if (typeof value === 'number' && value > 0) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color.replace('#', '') }
          };
          
          cell.font = {
            ...cell.font,
            color: { argb: textColor.replace('#', '') }
          };
        }
      }
    }

    return {
      type: 'highlight_positive',
      status: 'success',
      range
    };
  } catch (error) {
    return {
      type: 'highlight_positive',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyHighlightMax(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    const color = properties?.color || '#FFEB9C';
    const textColor = properties?.textColor || '#9C6500';

    // Find the maximum value in the range
    let maxValue = -Infinity;
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        if (typeof value === 'number' && value > maxValue) {
          maxValue = value;
        }
      }
    }
    
    // Highlight the maximum value(s)
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        if (typeof value === 'number' && value === maxValue) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color.replace('#', '') }
          };
          
          cell.font = {
            ...cell.font,
            color: { argb: textColor.replace('#', '') }
          };
        }
      }
    }

    return {
      type: 'highlight_max',
      status: 'success',
      range,
      maxValue
    };
  } catch (error) {
    return {
      type: 'highlight_max',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyHighlightMin(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    const color = properties?.color || '#FFD9D9';
    const textColor = properties?.textColor || '#9C0006';

    // Find the minimum value in the range
    let minValue = Infinity;
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        if (typeof value === 'number' && value < minValue) {
          minValue = value;
        }
      }
    }
    
    // Highlight the minimum value(s)
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        if (typeof value === 'number' && value === minValue) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color.replace('#', '') }
          };
          
          cell.font = {
            ...cell.font,
            color: { argb: textColor.replace('#', '') }
          };
        }
      }
    }

    return {
      type: 'highlight_min',
      status: 'success',
      range,
      minValue
    };
  } catch (error) {
    return {
      type: 'highlight_min',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}

function applyDataBars(sheet: ExcelJS.Worksheet, range: string, properties: any) {
  try {
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const endColNum = sheet.getColumn(endCol).number;     // Use getColumn().number
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);

    // Get the color for the data bars
    const gradientColor = properties?.color || '#638EC6';
    
    // Find min and max values in the range to calculate the scale
    let minValue = Infinity;
    let maxValue = -Infinity;
    
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        if (typeof value === 'number') {
          minValue = Math.min(minValue, value);
          maxValue = Math.max(maxValue, value);
        }
      }
    }
    
    // Avoid division by zero
    if (maxValue === minValue) {
      maxValue = minValue + 1;
    }
    
    // Apply data bars
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        const value = cell.value;
        
        if (typeof value === 'number') {
          // Calculate width based on min/max scale (0-100%)
          const normalizedValue = (value - minValue) / (maxValue - minValue);
          const widthPercent = Math.max(5, Math.min(100, normalizedValue * 100));
          
          // We'll simulate data bars with background colors and spacing
          // Could be enhanced with custom renderer in a real implementation
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: gradientColor.replace('#', '') },
            bgColor: { argb: 'FFFFFF' } // White background
          };
          
          // Apply number format if requested
          if (properties?.showValues !== false) {
            cell.numFmt = properties?.numberFormat || '0.00';
          } else {
            cell.numFmt = ';;;'; // Hide values
          }
        }
      }
    }

    return {
      type: 'data_bars',
      status: 'success',
      range,
      minValue,
      maxValue
    };
  } catch (error) {
    return {
      type: 'data_bars',
      status: 'error',
      message: error instanceof Error ? error.message : String(error),
      range
    };
  }
}
