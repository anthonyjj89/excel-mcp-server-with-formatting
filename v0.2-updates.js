// Updated write_sheet_data tool with formula fix and auto-fit feature
server.tool(
  'write_sheet_data',
  'Write data to the Excel sheet',
  {
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
    data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
      .describe('Data to write to the Excel sheet'),
    autoFit: z.boolean().optional().describe('Automatically adjust column widths to fit content')
  },
  async ({ fileAbsolutePath, sheetName, range, data, autoFit = false }) => {
    try {
      console.error(`Writing data to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
      const workbook = new ExcelJS.Workbook();
      
      // Try to read existing file, create new if doesn't exist
      try {
        await workbook.xlsx.readFile(fileAbsolutePath);
      } catch (e) {
        console.error(`File ${fileAbsolutePath} doesn't exist. Creating a new workbook.`);
      }
      
      // Enable formula calculation on load
      workbook.calcProperties = workbook.calcProperties || {};
      workbook.calcProperties.fullCalcOnLoad = true;
      
      // Get or create worksheet
      let worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
      }
      
      // Parse range and write data
      const { startCol, startRow } = parseRange(range);
      const startColNum = columnNameToNumber(startCol);
      
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        for (let j = 0; j < row.length; j++) {
          const cellAddress = `${numberToColumnName(startColNum + j)}${startRow + i}`;
          const cell = worksheet.getCell(cellAddress);
          
          // Check if the value is a formula string
          if (typeof row[j] === 'string' && row[j].startsWith('=')) {
            cell.value = { formula: row[j].substring(1) };
            // Force Excel to recalculate this formula when opened
            if (cell.model) {
              cell.model.result = undefined;
            }
          } else {
            cell.value = row[j];
          }
        }
      }
      
      // Handle auto-fit if requested
      if (autoFit) {
        // Get the columns that were modified
        const { startCol, endCol } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Track max width for each column
        const columnWidths = {};
        
        // Analyze content width for affected columns
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          for (let colNum = startColNum; colNum <= endColNum; colNum++) {
            const colLetter = numberToColumnName(colNum);
            const cell = row.getCell(colNum);
            
            let contentLength = 0;
            if (cell.text) {
              contentLength = cell.text.toString().length;
            } else if (cell.value !== null && cell.value !== undefined) {
              contentLength = cell.value.toString().length;
            } else if (cell.formula) {
              contentLength = cell.formula.toString().length + 1; // Add space for =
            }
            
            // Update max width if needed
            if (!columnWidths[colLetter] || contentLength > columnWidths[colLetter]) {
              columnWidths[colLetter] = contentLength;
            }
          }
        });
        
        // Apply column widths with padding
        for (let colNum = startColNum; colNum <= endColNum; colNum++) {
          const colLetter = numberToColumnName(colNum);
          const width = columnWidths[colLetter] || 0;
          worksheet.getColumn(colLetter).width = width + 2; // Add padding
        }
        
        console.error(`Auto-fitted columns ${startCol}-${endCol}`);
      }
      
      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return { 
        content: [{ 
          type: "text", 
          text: `Successfully wrote data to ${sheetName} in range ${range}${autoFit ? ' and auto-fitted columns' : ''}` 
        }]
      };
    } catch (error) {
      console.error(`Error writing sheet data: ${error.message}`);
      return {
        content: [{ 
          type: "text", 
          text: `Failed to write sheet data: ${error.message}` 
        }],
        isError: true
      };
    }
  }
);

// Updated write_sheet_formula tool with formula fix and auto-fit feature
server.tool(
  'write_sheet_formula',
  'Write formulas to the Excel sheet',
  {
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
    formulas: z.array(z.array(z.string())).describe('Formulas to write to the Excel sheet (e.g., "=A1+B1")'),
    autoFit: z.boolean().optional().describe('Automatically adjust column widths to fit content')
  },
  async ({ fileAbsolutePath, sheetName, range, formulas, autoFit = false }) => {
    try {
      console.error(`Writing formulas to ${fileAbsolutePath}, sheet: ${sheetName}, range: ${range}`);
      const workbook = new ExcelJS.Workbook();
      
      // Try to read existing file, create new if doesn't exist
      try {
        await workbook.xlsx.readFile(fileAbsolutePath);
      } catch (e) {
        console.error(`File ${fileAbsolutePath} doesn't exist. Creating a new workbook.`);
      }
      
      // Enable formula calculation on load
      workbook.calcProperties = workbook.calcProperties || {};
      workbook.calcProperties.fullCalcOnLoad = true;
      
      // Get or create worksheet
      let worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
      }
      
      // Parse range and write formulas
      const { startCol, startRow } = parseRange(range);
      const startColNum = columnNameToNumber(startCol);
      
      for (let i = 0; i < formulas.length; i++) {
        const row = formulas[i];
        for (let j = 0; j < row.length; j++) {
          const cellAddress = `${numberToColumnName(startColNum + j)}${startRow + i}`;
          const cell = worksheet.getCell(cellAddress);
          
          if (row[j]) {
            if (row[j].startsWith('=')) {
              cell.value = { formula: row[j].substring(1) };
            } else {
              cell.value = { formula: row[j] };
            }
            
            // Force Excel to recalculate this formula when the file is opened
            if (cell.model) {
              cell.model.result = undefined;
            }
          }
        }
      }
      
      // Handle auto-fit if requested
      if (autoFit) {
        // Get the columns that were modified
        const { startCol, endCol } = parseRange(range);
        const startColNum = columnNameToNumber(startCol);
        const endColNum = columnNameToNumber(endCol);
        
        // Track max width for each column
        const columnWidths = {};
        
        // Analyze content width for affected columns
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          for (let colNum = startColNum; colNum <= endColNum; colNum++) {
            const colLetter = numberToColumnName(colNum);
            const cell = row.getCell(colNum);
            
            let contentLength = 0;
            if (cell.text) {
              contentLength = cell.text.toString().length;
            } else if (cell.value !== null && cell.value !== undefined) {
              contentLength = cell.value.toString().length;
            } else if (cell.formula) {
              contentLength = cell.formula.toString().length + 1; // Add space for =
            }
            
            // Update max width if needed
            if (!columnWidths[colLetter] || contentLength > columnWidths[colLetter]) {
              columnWidths[colLetter] = contentLength;
            }
          }
        });
        
        // Apply column widths with padding
        for (let colNum = startColNum; colNum <= endColNum; colNum++) {
          const colLetter = numberToColumnName(colNum);
          const width = columnWidths[colLetter] || 0;
          worksheet.getColumn(colLetter).width = width + 2; // Add padding
        }
        
        console.error(`Auto-fitted columns ${startCol}-${endCol}`);
      }
      
      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return { 
        content: [{ 
          type: "text", 
          text: `Successfully wrote formulas to ${sheetName} in range ${range}${autoFit ? ' and auto-fitted columns' : ''}` 
        }]
      };
    } catch (error) {
      console.error(`Error writing sheet formulas: ${error.message}`);
      return {
        content: [{ 
          type: "text", 
          text: `Failed to write sheet formulas: ${error.message}` 
        }],
        isError: true
      };
    }
  }
);

// Also add a dedicated column auto-fit tool for more control
server.tool(
  'autofit_columns',
  'Automatically adjust column width based on content',
  {
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    columns: z.array(z.string()).optional().describe('Columns to auto-fit (e.g., ["A", "B", "C"]). Default: all columns'),
    padding: z.number().optional().describe('Additional padding to add (in characters). Default: 2'),
    minWidth: z.number().optional().describe('Minimum column width to set'),
    maxWidth: z.number().optional().describe('Maximum column width to set')
  },
  async ({ fileAbsolutePath, sheetName, columns, padding = 2, minWidth, maxWidth }) => {
    try {
      console.error(`Auto-fitting columns in ${fileAbsolutePath}, sheet: ${sheetName}`);
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        throw new Error(`Sheet "${sheetName}" not found`);
      }
      
      // Store max length for each column
      const columnWidths = {};
      
      // Process all rows to find the max content width for each column
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const column = worksheet.getColumn(colNumber);
          const columnLetter = column.letter;
          
          // Skip if we're only processing specific columns and this isn't one of them
          if (columns && !columns.includes(columnLetter)) {
            return;
          }
          
          let contentLength = 0;
          
          if (cell.text) {
            contentLength = cell.text.toString().length;
          } else if (cell.value !== null && cell.value !== undefined) {
            contentLength = cell.value.toString().length;
          } else if (cell.formula) {
            contentLength = cell.formula.toString().length + 1; // Add space for =
          }
          
          // Account for header/column name length too
          const headerLength = column.header ? column.header.toString().length : 0;
          
          // Update max width if needed
          if (!columnWidths[columnLetter] || contentLength > columnWidths[columnLetter]) {
            columnWidths[columnLetter] = Math.max(contentLength, headerLength);
          }
        });
      });
      
      // Set column widths based on content length plus padding
      const processed = [];
      Object.keys(columnWidths).forEach(columnLetter => {
        const column = worksheet.getColumn(columnLetter);
        let width = columnWidths[columnLetter] + padding;
        
        // Apply min/max constraints if specified
        if (minWidth !== undefined && width < minWidth) {
          width = minWidth;
        }
        if (maxWidth !== undefined && width > maxWidth) {
          width = maxWidth;
        }
        
        column.width = width;
        processed.push(columnLetter);
      });
      
      // Save the workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return {
        content: [{
          type: "text",
          text: `Successfully auto-fit columns in ${sheetName}. Columns adjusted: ${processed.join(', ')}`
        }]
      };
    } catch (error) {
      console.error(`Error auto-fitting columns: ${error.message}`);
      return {
        content: [{
          type: "text",
          text: `Failed to auto-fit columns: ${error.message}`
        }],
        isError: true
      };
    }
  }
);
