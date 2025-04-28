const ExcelJS = require('exceljs');
const path = require('path');

const workbook = new ExcelJS.Workbook();
const sheet1 = workbook.addWorksheet('Sample Data');
const sheet2 = workbook.addWorksheet('Formulas');

// Add data to Sheet1
sheet1.columns = [
  { header: 'ID', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 20 },
  { header: 'Email', key: 'email', width: 30 },
  { header: 'Phone', key: 'phone', width: 15 },
  { header: 'Amount', key: 'amount', width: 15 }
];

// Add some sample data
for (let i = 1; i <= 10; i++) {
  sheet1.addRow({
    id: i,
    name: `Person ${i}`,
    email: `person${i}@example.com`,
    phone: `${Math.floor(1000000000 + Math.random() * 9000000000)}`,
    amount: Math.floor(100 + Math.random() * 900) / 10
  });
}

// Add formulas to Sheet2
sheet2.columns = [
  { header: 'Value A', key: 'a', width: 15 },
  { header: 'Value B', key: 'b', width: 15 },
  { header: 'Sum', key: 'sum', width: 15 },
  { header: 'Product', key: 'product', width: 15 },
  { header: 'Average', key: 'average', width: 15 }
];

// Add data and formulas
for (let i = 1; i <= 5; i++) {
  const a = Math.floor(Math.random() * 100);
  const b = Math.floor(Math.random() * 100);
  
  const row = sheet2.addRow({
    a: a,
    b: b
  });
  
  const rowIndex = row.number;
  
  // Add formulas
  sheet2.getCell(`C${rowIndex}`).value = { formula: `A${rowIndex}+B${rowIndex}` };
  sheet2.getCell(`D${rowIndex}`).value = { formula: `A${rowIndex}*B${rowIndex}` };
  sheet2.getCell(`E${rowIndex}`).value = { formula: `(A${rowIndex}+B${rowIndex})/2` };
}

// Add a summary row with formulas
const summaryRow = sheet2.addRow({});
const summaryRowIndex = summaryRow.number;
sheet2.getCell(`A${summaryRowIndex}`).value = 'Totals:';
sheet2.getCell(`C${summaryRowIndex}`).value = { formula: `SUM(C2:C${summaryRowIndex-1})` };
sheet2.getCell(`D${summaryRowIndex}`).value = { formula: `SUM(D2:D${summaryRowIndex-1})` };
sheet2.getCell(`E${summaryRowIndex}`).value = { formula: `AVERAGE(E2:E${summaryRowIndex-1})` };

// Save the file
const filePath = path.resolve(__dirname, 'excel_files', 'sample-data.xlsx');
workbook.xlsx.writeFile(filePath)
  .then(() => {
    console.log(`Sample Excel file created at: ${filePath}`);
  })
  .catch(err => {
    console.error('Error creating Excel file:', err);
  });
