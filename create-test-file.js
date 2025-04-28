const ExcelJS = require('exceljs');
const path = require('path');

async function createTestFile() {
  const workbook = new ExcelJS.Workbook();
  
  // Set workbook properties
  workbook.creator = 'Excel MCP Server';
  workbook.lastModifiedBy = 'Excel MCP Server';
  workbook.created = new Date();
  workbook.modified = new Date();
  
  // Create a worksheet
  const worksheet = workbook.addWorksheet('Sample Data');
  
  // Add column headers
  worksheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Age', key: 'age', width: 10 },
    { header: 'City', key: 'city', width: 20 },
    { header: 'Country', key: 'country', width: 20 }
  ];
  
  // Add rows
  worksheet.addRow({ name: 'John Doe', email: 'john@example.com', age: 35, city: 'New York', country: 'USA' });
  worksheet.addRow({ name: 'Jane Smith', email: 'jane@example.com', age: 28, city: 'London', country: 'UK' });
  worksheet.addRow({ name: 'Bob Johnson', email: 'bob@example.com', age: 42, city: 'Sydney', country: 'Australia' });
  worksheet.addRow({ name: 'Alice Brown', email: 'alice@example.com', age: 31, city: 'Toronto', country: 'Canada' });
  worksheet.addRow({ name: 'Charlie Lee', email: 'charlie@example.com', age: 39, city: 'Berlin', country: 'Germany' });
  
  // Add a second worksheet with formulas
  const formulaSheet = workbook.addWorksheet('Formulas');
  
  // Add column headers
  formulaSheet.columns = [
    { header: 'Value 1', key: 'val1', width: 15 },
    { header: 'Value 2', key: 'val2', width: 15 },
    { header: 'Sum', key: 'sum', width: 15 },
    { header: 'Product', key: 'product', width: 15 },
    { header: 'Average', key: 'average', width: 15 }
  ];
  
  // Add data rows
  formulaSheet.addRow({ val1: 10, val2: 20 });
  formulaSheet.addRow({ val1: 30, val2: 40 });
  formulaSheet.addRow({ val1: 50, val2: 60 });
  
  // Add formulas
  formulaSheet.getCell('C2').value = { formula: 'A2+B2' };
  formulaSheet.getCell('D2').value = { formula: 'A2*B2' };
  formulaSheet.getCell('E2').value = { formula: '(A2+B2)/2' };
  
  formulaSheet.getCell('C3').value = { formula: 'A3+B3' };
  formulaSheet.getCell('D3').value = { formula: 'A3*B3' };
  formulaSheet.getCell('E3').value = { formula: '(A3+B3)/2' };
  
  formulaSheet.getCell('C4').value = { formula: 'A4+B4' };
  formulaSheet.getCell('D4').value = { formula: 'A4*B4' };
  formulaSheet.getCell('E4').value = { formula: '(A4+B4)/2' };
  
  // Save the workbook
  const filePath = path.join(__dirname, 'excel_files', 'sample-data.xlsx');
  await workbook.xlsx.writeFile(filePath);
  
  console.log(`Test Excel file created at: ${filePath}`);
}

// Run the function
createTestFile().catch(error => {
  console.error('Error creating test file:', error);
});
