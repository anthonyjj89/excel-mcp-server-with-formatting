const ExcelJS = require('exceljs');
const path = require('path');

async function testAddBorders() {
  try {
    // Load the existing test file with the new sheet
    const filePath = path.join(__dirname, 'excel_files', 'test-data-with-new-sheet.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Get the Test Sheet worksheet
    const sheet = workbook.getWorksheet('Test Sheet');
    if (!sheet) {
      throw new Error('Test Sheet not found');
    }
    
    console.log('Adding borders to Test Sheet...');
    
    // Define the range
    const range = 'A1:C3';
    console.log(`Range: ${range}`);
    
    // Parse range
    const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!rangeMatch) {
      throw new Error(`Invalid range format: ${range}`);
    }

    const [_, startCol, startRow, endCol, endRow] = rangeMatch;
    const startColNum = sheet.getColumn(startCol).number; // Using our fixed method
    const endColNum = sheet.getColumn(endCol).number;     // Using our fixed method
    const startRowNum = parseInt(startRow);
    const endRowNum = parseInt(endRow);
    
    console.log(`Column numbers: ${startColNum}-${endColNum}, Row numbers: ${startRowNum}-${endRowNum}`);
    
    // Apply borders to cells
    for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
      const row = sheet.getRow(rowNum);
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = row.getCell(colNum);
        
        // Apply borders
        cell.border = {
          top: { style: 'thin', color: { argb: '000000' } },
          left: { style: 'thin', color: { argb: '000000' } },
          bottom: { style: 'thin', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: '000000' } }
        };
        
        // Apply thick borders to the outline
        if (rowNum === startRowNum) {
          cell.border.top = { style: 'medium', color: { argb: '000000' } };
        }
        if (rowNum === endRowNum) {
          cell.border.bottom = { style: 'medium', color: { argb: '000000' } };
        }
        if (colNum === startColNum) {
          cell.border.left = { style: 'medium', color: { argb: '000000' } };
        }
        if (colNum === endColNum) {
          cell.border.right = { style: 'medium', color: { argb: '000000' } };
        }
      }
    }

    // Save the file with a different name
    const newFilePath = path.join(__dirname, 'excel_files', 'test-data-with-borders.xlsx');
    await workbook.xlsx.writeFile(newFilePath);
    
    console.log('\nTest completed successfully!');
    console.log(`New file created at: ${newFilePath}`);
  } catch (error) {
    console.error('Error testing add borders:', error);
  }
}

testAddBorders();
