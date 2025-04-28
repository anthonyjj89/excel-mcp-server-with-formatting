import ExcelJS from 'exceljs';
import { cellRangeToIndices } from './helpers';

/**
 * Reads formulas from an Excel sheet.
 * @param params The parameters for reading formulas.
 * @returns The formulas from the sheet.
 */
export async function readSheetFormulaHandler(params: {
  fileAbsolutePath: string;
  sheetName: string;
  range?: string;
  knownPagingRanges?: string[];
}) {
  try {
    const { fileAbsolutePath, sheetName, range } = params;
    
    // Load the workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(fileAbsolutePath);
    
    // Get the worksheet
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      return { error: { message: `Sheet "${sheetName}" not found in workbook.` } };
    }
    
    // Read formulas based on range or all
    const formulas: Array<Array<string | null>> = [];
    
    if (range) {
      // Parse the range (e.g., "A1:C10")
      const { startRow, startCol, endRow, endCol } = cellRangeToIndices(range);
      
      // Extract formulas in the specified range
      for (let row = startRow; row <= endRow; row++) {
        const rowData: Array<string | null> = [];
        
        for (let col = startCol; col <= endCol; col++) {
          const cell = worksheet.getCell(row, col);
          rowData.push(cell.formula || null);
        }
        
        formulas.push(rowData);
      }
    } else {
      // Read all formulas
      worksheet.eachRow((row, rowNumber) => {
        const rowData: Array<string | null> = [];
        
        row.eachCell((cell, colNumber) => {
          if (colNumber > rowData.length) {
            // Fill gaps with null
            for (let i = rowData.length; i < colNumber - 1; i++) {
              rowData.push(null);
            }
          }
          
          rowData.push(cell.formula || null);
        });
        
        // Ensure we have data for this row
        if (rowData.length > 0) {
          // Fill any gaps in the formulas array with empty rows
          if (rowNumber > formulas.length) {
            for (let i = formulas.length; i < rowNumber - 1; i++) {
              formulas.push([]);
            }
          }
          
          formulas.push(rowData);
        }
      });
    }
    
    return { result: { formulas } };
  } catch (error) {
    console.error('Error reading formulas from Excel file:', error);
    return { error: { message: `Error reading formulas: ${error instanceof Error ? error.message : String(error)}` } };
  }
}
