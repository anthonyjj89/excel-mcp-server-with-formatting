const ExcelJS = require('exceljs');
const path = require('path');

async function testAddWorksheet() {
  try {
    // Load the existing test file
    const filePath = path.join(__dirname, 'excel_files', 'test-data.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Get current worksheets
    console.log('Current worksheets:');
    workbook.worksheets.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.name}`);
    });
    
    // Add a new worksheet
    console.log('\nAdding new worksheet "Test Sheet"...');
    const newSheet = workbook.addWorksheet('Test Sheet', {
      properties: {
        tabColor: { argb: 'FF0000' }
      }
    });
    
    // Add some data to the new worksheet
    newSheet.columns = [
      { header: 'Column A', key: 'colA', width: 15 },
      { header: 'Column B', key: 'colB', width: 15 },
      { header: 'Column C', key: 'colC', width: 15 }
    ];
    
    newSheet.addRow({ colA: 'Value A1', colB: 'Value B1', colC: 'Value C1' });
    newSheet.addRow({ colA: 'Value A2', colB: 'Value B2', colC: 'Value C2' });
    
    // Format the header row
    const headerRow = newSheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF00' }
    };
    headerRow.commit();
    
    // Save the file with a different name to preserve the original
    const newFilePath = path.join(__dirname, 'excel_files', 'test-data-with-new-sheet.xlsx');
    await workbook.xlsx.writeFile(newFilePath);
    
    // Verify the new file
    const verifyWorkbook = new ExcelJS.Workbook();
    await verifyWorkbook.xlsx.readFile(newFilePath);
    
    console.log('\nWorksheets in the new file:');
    verifyWorkbook.worksheets.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.name}`);
    });
    
    console.log('\nTest completed successfully!');
    console.log(`New file created at: ${newFilePath}`);
  } catch (error) {
    console.error('Error testing add worksheet:', error);
  }
}

testAddWorksheet();
