/**
 * Test script for Excel MCP formula calculation fix
 * 
 * This script:
 * 1. Creates a test Excel file with formulas
 * 2. Writes the formulas using the updated tool
 * 3. Opens the file to verify formulas are calculated correctly
 */

const ExcelJS = require('exceljs');
const path = require('path');
const { execSync } = require('child_process');

async function testFormulaFix() {
  console.log('Testing Excel MCP Formula Fix');
  console.log('============================');
  
  // Create test file path
  const filePath = path.join(__dirname, 'excel_files', 'formula-fix-test.xlsx');
  console.log(`Test file: ${filePath}`);
  
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Formula Test');
  
  // Add some test data
  worksheet.getCell('A1').value = 10;
  worksheet.getCell('A2').value = 20;
  worksheet.getCell('A3').value = 30;
  worksheet.getCell('A4').value = 40;
  worksheet.getCell('A5').value = 50;
  
  worksheet.getCell('B1').value = 5;
  worksheet.getCell('B2').value = 15;
  worksheet.getCell('B3').value = 25;
  worksheet.getCell('B4').value = 35;
  worksheet.getCell('B5').value = 45;
  
  // Add header row
  worksheet.getCell('A7').value = 'Formula';
  worksheet.getCell('B7').value = 'Expected Result';
  worksheet.getCell('C7').value = 'Actual Result';
  
  // Make header bold
  worksheet.getCell('A7').font = { bold: true };
  worksheet.getCell('B7').font = { bold: true };
  worksheet.getCell('C7').font = { bold: true };
  
  // Save the basic workbook
  await workbook.xlsx.writeFile(filePath);
  console.log('Created base test file with data');
  
  // Now add formulas using our updated tool logic
  console.log('Adding formulas with fix applied...');
  
  // Re-open the workbook
  const updatedWorkbook = new ExcelJS.Workbook();
  await updatedWorkbook.xlsx.readFile(filePath);
  
  // Enable formula calculation on load - FORMULA FIX
  updatedWorkbook.calcProperties = updatedWorkbook.calcProperties || {};
  updatedWorkbook.calcProperties.fullCalcOnLoad = true;
  
  const updatedWorksheet = updatedWorkbook.getWorksheet('Formula Test');
  
  // Add test formulas
  const formulas = [
    { cell: 'A8', formula: '=SUM(A1:A5)', expected: 150 },
    { cell: 'A9', formula: '=AVERAGE(A1:A5)', expected: 30 },
    { cell: 'A10', formula: '=A1+B1', expected: 15 },
    { cell: 'A11', formula: '=A2*B2', expected: 300 },
    { cell: 'A12', formula: '=A3-B3', expected: 5 },
    { cell: 'A13', formula: '=A4/B4', expected: 1.143 }
  ];
  
  // Add formulas and expected results
  formulas.forEach(({ cell, formula, expected }) => {
    const cellRef = updatedWorksheet.getCell(cell);
    cellRef.value = { formula: formula.startsWith('=') ? formula.substring(1) : formula };
    
    // Force Excel to recalculate this formula - FORMULA FIX
    if (cellRef.model) {
      cellRef.model.result = undefined;
    }
    
    // Add expected result in B column
    const rowNum = cellRef.row;
    updatedWorksheet.getCell(`B${rowNum}`).value = expected;
    
    // Add a formula to check if result matches expected
    updatedWorksheet.getCell(`C${rowNum}`).value = { 
      formula: `IF(ROUND(${cell},3)=ROUND(B${rowNum},3),"✓","✗")` 
    };
  });
  
  // Save the workbook
  await updatedWorkbook.xlsx.writeFile(filePath);
  console.log('Test file saved with formulas');
  
  console.log('\nTest completed. Please open the file to verify that:');
  console.log('1. All formulas are calculated correctly');
  console.log('2. The checkmarks in column C show formula results match expected values');
  console.log(`3. You don't need to manually enter formulas or press Enter to activate them\n`);
  
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

testFormulaFix().catch(err => {
  console.error('Error:', err);
});
