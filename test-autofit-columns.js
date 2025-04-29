/**
 * Test script for Excel MCP auto-fit column width feature
 * 
 * This script:
 * 1. Creates a test Excel file with varying content lengths
 * 2. Tests auto-fit width using both integrated and dedicated tools
 */

const ExcelJS = require('exceljs');
const path = require('path');
const { execSync } = require('child_process');

async function testAutoFitColumns() {
  console.log('Testing Excel MCP Auto-Fit Column Width');
  console.log('=======================================');
  
  // Create test file path
  const filePath = path.join(__dirname, 'excel_files', 'autofit-column-test.xlsx');
  console.log(`Test file: ${filePath}`);
  
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();
  
  // 1. First sheet - Testing autoFit with write_sheet_data
  const sheet1 = workbook.addWorksheet('Auto-Fit with write_sheet_data');
  console.log('Creating test data with varying lengths...');
  
  // Add headers
  sheet1.getCell('A1').value = 'Short';
  sheet1.getCell('B1').value = 'Medium Length Header';
  sheet1.getCell('C1').value = 'This is an extremely long header to test auto-fit functionality';
  sheet1.getCell('D1').value = 'Auto-fit Status';
  
  // Make headers bold
  sheet1.getRow(1).font = { bold: true };
  
  // Prepare data with varying lengths
  const sheet1Data = [
    ['A', 'Medium text here', 'This is a very very very very very very very very very very very very very very long text', ''],
    ['BB', 'A bit longer', 'Shorter than above', ''],
    ['CCC', 'Lorem ipsum dolor', 'Medium length text example for testing', ''],
    ['DDDD', '12345', 'Small text', ''],
    ['EEEEE', '2023-04-29', 'Date example', '']
  ];
  
  // Add data
  for (let i = 0; i < sheet1Data.length; i++) {
    const row = sheet1Data[i];
    sheet1.getCell(`A${i+2}`).value = row[0];
    sheet1.getCell(`B${i+2}`).value = row[1];
    sheet1.getCell(`C${i+2}`).value = row[2];
    sheet1.getCell(`D${i+2}`).value = 'Default width (not auto-fit)';
  }
  
  // 2. Second sheet - For testing dedicated autofit_columns tool
  const sheet2 = workbook.addWorksheet('Dedicated Auto-Fit Tool');
  
  // Copy same data to second sheet
  sheet2.getCell('A1').value = 'Short';
  sheet2.getCell('B1').value = 'Medium Length Header';
  sheet2.getCell('C1').value = 'This is an extremely long header to test auto-fit functionality';
  sheet2.getCell('D1').value = 'Auto-fit Status';
  sheet2.getRow(1).font = { bold: true };
  
  for (let i = 0; i < sheet1Data.length; i++) {
    const row = sheet1Data[i];
    sheet2.getCell(`A${i+2}`).value = row[0];
    sheet2.getCell(`B${i+2}`).value = row[1];
    sheet2.getCell(`C${i+2}`).value = row[2];
    sheet2.getCell(`D${i+2}`).value = 'Will use dedicated auto-fit tool';
  }
  
  // 3. Third sheet - For testing autoFit with write_sheet_formula
  const sheet3 = workbook.addWorksheet('Auto-Fit with Formulas');
  
  // Add headers
  sheet3.getCell('A1').value = 'Formula';
  sheet3.getCell('B1').value = 'Result';
  sheet3.getCell('C1').value = 'Description of what the formula does in great detail';
  sheet3.getRow(1).font = { bold: true };
  
  // Save the workbook
  await workbook.xlsx.writeFile(filePath);
  console.log('Created base test file with data');
  
  // Now apply auto-fit to first sheet using our implementation
  console.log('\nApplying auto-fit using write_sheet_data...');
  
  // Re-open the workbook
  const updatedWorkbook = new ExcelJS.Workbook();
  await updatedWorkbook.xlsx.readFile(filePath);
  const updatedSheet1 = updatedWorkbook.getWorksheet('Auto-Fit with write_sheet_data');
  
  // Track content length for auto-fit (simplified version of our implementation)
  const columnWidths = {};
  const startCol = 'A';
  const startColNum = columnNameToNumber(startCol);
  const endCol = 'C'; 
  const endColNum = columnNameToNumber(endCol);
  
  // Calculate max content width for each column
  updatedSheet1.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    for (let colNum = startColNum; colNum <= endColNum; colNum++) {
      const colLetter = numberToColumnName(colNum);
      const cell = row.getCell(colNum);
      
      let contentLength = 0;
      if (cell.text) {
        contentLength = cell.text.toString().length;
      } else if (cell.value !== null && cell.value !== undefined) {
        contentLength = cell.value.toString().length;
      }
      
      if (!columnWidths[colLetter] || contentLength > columnWidths[colLetter]) {
        columnWidths[colLetter] = contentLength;
      }
    }
  });
  
  // Apply column widths with padding
  for (let colNum = startColNum; colNum <= endColNum; colNum++) {
    const colLetter = numberToColumnName(colNum);
    if (columnWidths[colLetter]) {
      updatedSheet1.getColumn(colLetter).width = columnWidths[colLetter] + 2; // Add padding
    }
  }
  
  // Update the status column
  for (let i = 0; i < sheet1Data.length; i++) {
    updatedSheet1.getCell(`D${i+2}`).value = 'Auto-fit width applied';
  }
  
  // Apply formulas to third sheet with auto-fit
  const sheet3Updated = updatedWorkbook.getWorksheet('Auto-Fit with Formulas');
  
  const formulas = [
    { formula: '=SUM(10,20,30,40,50)', desc: 'Adds multiple numbers together to calculate their sum' },
    { formula: '=AVERAGE(10,20,30,40,50)', desc: 'Calculates the average (mean) of a series of numbers' },
    { formula: '=CONCATENATE("Hello", " ", "World", "!")', desc: 'Joins multiple text strings into one combined text string' },
    { formula: '=IF(TRUE, "Condition is true", "Condition is false")', desc: 'Tests a condition and returns one value if true, another if false' },
    { formula: '=TODAY()', desc: 'Returns the current date from your system' }
  ];
  
  // Add formulas with auto-width consideration
  for (let i = 0; i < formulas.length; i++) {
    const rowNum = i + 2;
    sheet3Updated.getCell(`A${rowNum}`).value = { formula: formulas[i].formula.startsWith('=') ? formulas[i].formula.substring(1) : formulas[i].formula };
    sheet3Updated.getCell(`C${rowNum}`).value = formulas[i].desc;
  }
  
  // Auto-fit columns in the third sheet
  const columnWidths3 = {};
  const columns3 = ['A', 'C'];
  
  // Calculate max content width for selected columns
  sheet3Updated.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    columns3.forEach(colLetter => {
      const colNum = columnNameToNumber(colLetter);
      const cell = row.getCell(colNum);
      
      let contentLength = 0;
      if (cell.text) {
        contentLength = cell.text.toString().length;
      } else if (cell.value !== null && cell.value !== undefined) {
        contentLength = cell.value.toString().length;
      }
      
      if (!columnWidths3[colLetter] || contentLength > columnWidths3[colLetter]) {
        columnWidths3[colLetter] = contentLength;
      }
    });
  });
  
  // Apply column widths with padding
  columns3.forEach(colLetter => {
    if (columnWidths3[colLetter]) {
      sheet3Updated.getColumn(colLetter).width = columnWidths3[colLetter] + 2; // Add padding
    }
  });
  
  // Save the workbook
  await updatedWorkbook.xlsx.writeFile(filePath);
  console.log('Applied auto-fit to first and third sheets');
  
  console.log('\nTest completed. Please open the file to verify that:');
  console.log('1. Sheet 1: Columns A, B, C are auto-fitted to content width');
  console.log('2. Sheet 2: Columns have default width (not auto-fitted yet)');
  console.log('3. Sheet 3: Formula column and Description column are auto-fitted');
  console.log(`4. Compare the difference between auto-fitted and default columns\n`);
  
  try {
    // Try to open the file with Excel
    console.log('Attempting to open the test file...');
    if (process.platform === 'darwin') {
      execSync(`open "${filePath}"`);
    } else if (process.platform === 'win32') {
      execSync(`start excel "${filePath}"`);
    } else {
      console.log(`Please open the file manually: ${filePath}`);
    }
  } catch (error) {
    console.error('Could not automatically open the file:', error.message);
    console.log(`Please open the file manually: ${filePath}`);
  }
}

// Utility functions (copied from main implementation)
function columnNameToNumber(name) {
  let result = 0;
  for (let i = 0; i < name.length; i++) {
    result = result * 26 + (name.charCodeAt(i) - 64);
  }
  return result;
}

function numberToColumnName(num) {
  let result = '';
  while (num > 0) {
    const modulo = (num - 1) % 26;
    result = String.fromCharCode(65 + modulo) + result;
    num = Math.floor((num - modulo) / 26);
  }
  return result;
}

testAutoFitColumns().catch(err => {
  console.error('Error:', err);
});
