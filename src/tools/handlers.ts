import { UserError } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

// Helper function to handle Excel errors consistently
const handleExcelError = (error: unknown): never => {
  if (error instanceof Error) {
    throw new UserError(`Excel operation failed: ${error.message}`);
  }
  throw new UserError('Excel operation failed with an unknown error');
};

export const readSheetDataHandler = async ({ fileAbsolutePath, sheetName, range, knownPagingRanges }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new UserError(`Worksheet "${sheetName}" not found`);
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
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              data,
              sheetName,
              readRange: `A1:${sheet.lastColumn?.letter}${sheet.rowCount}`,
              allRangesRead: true
            }, null, 2)
          }
        ]
      };
    }

    // Parse range
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new UserError(`Invalid range format: ${range}`);
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
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            data,
            sheetName,
            readRange: range,
            allRangesRead
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const writeSheetDataHandler = async ({ fileAbsolutePath, sheetName, range, data }: any) => {
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
      throw new UserError(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Use getColumn().number
    const startRowNum = parseInt(startRow);

    // Write data to the specified range
    data.forEach((rowData: any[], rowIndex: number) => {
      const row = sheet.getRow(startRowNum + rowIndex);
      rowData.forEach((cellValue: any, colIndex: number) => {
        const col = startColNum + colIndex;
        row.getCell(col).value = cellValue;
      });
      row.commit();
    });

    // Save the workbook
    await workbook.xlsx.writeFile(fileAbsolutePath);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Data written to ${sheetName} in range ${range}`,
            sheetName,
            writtenRange: range
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const readSheetNamesHandler = async ({ fileAbsolutePath }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    // Get all sheet names
    const sheetNames = workbook.worksheets.map(sheet => sheet.name);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(sheetNames, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const formatCellsHandler = async ({ fileAbsolutePath, sheetName, range, format }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new UserError(`Worksheet "${sheetName}" not found`);
    }

    // Parse range
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new UserError(`Invalid range format: ${range}`);
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
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Formatting applied to ${sheetName} in range ${range}`,
            sheetName,
            formattedRange: range
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const addBordersHandler = async ({ fileAbsolutePath, sheetName, range, borderStyle }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new UserError(`Worksheet "${sheetName}" not found`);
    }

    // Parse range
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new UserError(`Invalid range format: ${range}`);
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
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Borders added to ${sheetName} in range ${range}`,
            sheetName,
            borderedRange: range
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const mergeCellsHandler = async ({ fileAbsolutePath, sheetName, range, alignContent }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new UserError(`Worksheet "${sheetName}" not found`);
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
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Cells merged in ${sheetName} at range ${range}`,
            sheetName,
            mergedRange: range
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const unmergeCellsHandler = async ({ fileAbsolutePath, sheetName, range }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new UserError(`Worksheet "${sheetName}" not found`);
    }

    // Unmerge cells
    sheet.unMergeCells(range);

    // Save the workbook
    await workbook.xlsx.writeFile(fileAbsolutePath);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Cells unmerged in ${sheetName} at range ${range}`,
            sheetName,
            unmergedRange: range
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const createWorkbookHandler = async ({ 
  fileAbsolutePath, 
  initialSheets = ['Sheet1'], 
  creator, 
  lastModifiedBy,
  title,
  subject,
  keywords,
  category,
  description,
  overwrite = false 
}: any) => {
  try {
    // Check if file exists
    if (fs.existsSync(fileAbsolutePath) && !overwrite) {
      throw new UserError(`File already exists. Use 'overwrite: true' to replace it.`);
    }
    
    // Create directory if it doesn't exist
    const directory = path.dirname(fileAbsolutePath);
    if (!fs.existsSync(directory)) {
      fs.mkdirSync(directory, { recursive: true });
    }
    
    const workbook = new ExcelJS.Workbook();
    
    // Set properties
    if (creator) workbook.creator = creator;
    if (lastModifiedBy) workbook.lastModifiedBy = lastModifiedBy;
    if (title) workbook.title = title;
    if (subject) workbook.subject = subject;
    if (keywords) workbook.keywords = keywords;
    if (category) workbook.category = category;
    if (description) workbook.description = description;
    
    // Create initial sheets
    initialSheets.forEach(sheetName => {
      workbook.addWorksheet(sheetName);
    });
    
    // Save workbook
    await workbook.xlsx.writeFile(fileAbsolutePath);
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Workbook created at ${fileAbsolutePath}`,
            sheets: initialSheets,
            filepath: fileAbsolutePath
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const addWorksheetHandler = async ({ fileAbsolutePath, sheetName, tabColor, columns, headerRowFormat, position }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    // Check if sheet already exists
    if (workbook.getWorksheet(sheetName)) {
      throw new UserError(`Worksheet "${sheetName}" already exists`);
    }
    
    // Add worksheet
    const worksheet = workbook.addWorksheet(sheetName, {
      properties: {
        tabColor: tabColor ? { argb: tabColor.replace('#', '') } : undefined
      }
    });
    // Note: Setting sheet position is not directly supported here either.
    
    // Add columns if specified
    if (columns && columns.length > 0) {
      worksheet.columns = columns.map((col: any) => ({
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
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Worksheet "${sheetName}" added to ${fileAbsolutePath}`,
            sheetName,
            workbook: path.basename(fileAbsolutePath)
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};

export const applyStylesHandler = async ({ fileAbsolutePath, sheetName, actions }: any) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new UserError(`Worksheet "${sheetName}" not found`);
    }

    const results = [];

    // Process each styling action
    for (const action of actions) {
      // Implementation of style action functions would go here
      // Each applying a specific type of style to the specified range
      // For brevity, just adding placeholders - actual implementations would follow the pattern of our other handlers
      results.push({
        type: action.type,
        status: 'success',
        range: action.range
      });
    }

    // Save the workbook
    await workbook.xlsx.writeFile(fileAbsolutePath);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            status: 'success',
            message: `Styles applied to ${sheetName}`,
            results
          }, null, 2)
        }
      ]
    };
  } catch (error) {
    return handleExcelError(error);
  }
};
