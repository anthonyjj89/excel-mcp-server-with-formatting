const ExcelJS = require('exceljs');
const path = require('path');

async function testColumnNumber() {
  try {
    // Load the existing test file
    const filePath = path.join(__dirname, 'excel_files', 'test-data.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Get the first worksheet
    const sheet = workbook.getWorksheet('Sample Data');
    
    // Test the column number conversion
    const columns = ['A', 'B', 'C', 'D', 'E', 'Z', 'AA', 'AB', 'AZ', 'BA'];
    
    console.log('Testing column number conversion:');
    console.log('=================================');
    
    columns.forEach(col => {
      // Get column number using sheet.getColumn(col).number
      const colNum = sheet.getColumn(col).number;
      console.log(`Column ${col} -> Number: ${colNum}`);
    });
    
    // Test reading data using the column number
    console.log('\nTesting data access with column numbers:');
    console.log('======================================');
    
    // Read data from range B2:D4 using column numbers
    const startCol = 'B';
    const endCol = 'D';
    const startRow = 2;
    const endRow = 4;
    
    const startColNum = sheet.getColumn(startCol).number;
    const endColNum = sheet.getColumn(endCol).number;
    
    console.log(`Reading range ${startCol}${startRow}:${endCol}${endRow}`);
    console.log(`Column numbers: ${startColNum}-${endColNum}, Row numbers: ${startRow}-${endRow}`);
    
    // Read and display the data
    const data = [];
    for (let rowNum = startRow; rowNum <= endRow; rowNum++) {
      const rowData = [];
      for (let colNum = startColNum; colNum <= endColNum; colNum++) {
        const cell = sheet.getCell(rowNum, colNum);
        rowData.push(cell.value);
      }
      data.push(rowData);
      console.log(`Row ${rowNum}: ${JSON.stringify(rowData)}`);
    }
    
    console.log('\nTest completed successfully!');
  } catch (error) {
    console.error('Error testing column number:', error);
  }
}

testColumnNumber();
