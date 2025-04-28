const { FastMCP } = require('@modelcontextprotocol/sdk');
const ExcelJS = require('exceljs');

// Create the MCP server instance
const server = new FastMCP({
  name: "Excel MCP with Formatting",
  version: "1.0.0",
  description: "An MCP server for Excel with rich formatting capabilities"
});

// Read Sheet Names Tool
server.addTool({
  name: "read_sheet_names",
  description: "List all sheet names in an Excel file",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      }
    },
    required: ["fileAbsolutePath"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath } = params;
      console.error(`Reading sheet names from: ${fileAbsolutePath}`);
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const sheetNames = workbook.worksheets.map(sheet => sheet.name);
      
      return { 
        result: { sheetNames }
      };
    } catch (error) {
      console.error("Error reading sheet names:", error);
      return { error: { message: `Error reading sheet names: ${error.message}` } };
    }
  }
});

// Read Sheet Data Tool
server.addTool({
  name: "read_sheet_data",
  description: "Read data from Excel sheet with pagination.",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      },
      sheetName: {
        type: "string",
        description: "Sheet name in the Excel file"
      },
      range: {
        type: "string",
        description: "Range of cells to read in the Excel sheet (e.g., \"A1:C10\"). [default: first paging range]"
      },
      knownPagingRanges: {
        type: "array",
        items: {
          type: "string"
        },
        description: "List of already read paging ranges"
      }
    },
    required: ["fileAbsolutePath", "sheetName"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath, sheetName, range } = params;
      console.error(`Reading data from: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range || 'default'}`);
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return { error: { message: `Sheet "${sheetName}" not found in workbook.` } };
      }
      
      let data = [];
      
      if (range) {
        // Parse range
        const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return { error: { message: `Invalid range format: ${range}` } };
        }
        
        const [_, startCol, startRow, endCol, endRow] = rangeMatch;
        
        // Convert column letters to numbers
        const getColumnNumber = (col) => {
          let num = 0;
          for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
          }
          return num;
        };
        
        const startColNum = getColumnNumber(startCol);
        const endColNum = getColumnNumber(endCol);
        const startRowNum = parseInt(startRow, 10);
        const endRowNum = parseInt(endRow, 10);
        
        // Read data from range
        for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
          const rowData = [];
          for (let colNum = startColNum; colNum <= endColNum; colNum++) {
            // Convert column number back to letter
            const getColumnLetter = (num) => {
              let letter = '';
              while (num > 0) {
                const remainder = (num - 1) % 26;
                letter = String.fromCharCode(65 + remainder) + letter;
                num = Math.floor((num - 1) / 26);
              }
              return letter;
            };
            
            const colLetter = getColumnLetter(colNum);
            const cell = worksheet.getCell(`${colLetter}${rowNum}`);
            rowData.push(cell.value);
          }
          data.push(rowData);
        }
      } else {
        // Read all data
        worksheet.eachRow((row, rowNumber) => {
          const rowData = [];
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
          });
          data[rowNumber - 1] = rowData;
        });
      }
      
      return { 
        result: { data }
      };
    } catch (error) {
      console.error("Error reading sheet data:", error);
      return { error: { message: `Error reading sheet data: ${error.message}` } };
    }
  }
});

// Read Sheet Formula Tool
server.addTool({
  name: "read_sheet_formula",
  description: "Read formulas from Excel sheet with pagination.",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      },
      sheetName: {
        type: "string",
        description: "Sheet name in the Excel file"
      },
      range: {
        type: "string",
        description: "Range of cells to read in the Excel sheet (e.g., \"A1:C10\"). [default: first paging range]"
      },
      knownPagingRanges: {
        type: "array",
        items: {
          type: "string"
        },
        description: "List of already read paging ranges"
      }
    },
    required: ["fileAbsolutePath", "sheetName"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath, sheetName, range } = params;
      console.error(`Reading formulas from: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range || 'default'}`);
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return { error: { message: `Sheet "${sheetName}" not found in workbook.` } };
      }
      
      let formulas = [];
      
      if (range) {
        // Parse range
        const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (!rangeMatch) {
          return { error: { message: `Invalid range format: ${range}` } };
        }
        
        const [_, startCol, startRow, endCol, endRow] = rangeMatch;
        
        // Convert column letters to numbers
        const getColumnNumber = (col) => {
          let num = 0;
          for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
          }
          return num;
        };
        
        const startColNum = getColumnNumber(startCol);
        const endColNum = getColumnNumber(endCol);
        const startRowNum = parseInt(startRow, 10);
        const endRowNum = parseInt(endRow, 10);
        
        // Read formulas from range
        for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
          const rowData = [];
          for (let colNum = startColNum; colNum <= endColNum; colNum++) {
            // Convert column number back to letter
            const getColumnLetter = (num) => {
              let letter = '';
              while (num > 0) {
                const remainder = (num - 1) % 26;
                letter = String.fromCharCode(65 + remainder) + letter;
                num = Math.floor((num - 1) / 26);
              }
              return letter;
            };
            
            const colLetter = getColumnLetter(colNum);
            const cell = worksheet.getCell(`${colLetter}${rowNum}`);
            rowData.push(cell.formula || null);
          }
          formulas.push(rowData);
        }
      } else {
        // Read all formulas
        worksheet.eachRow((row, rowNumber) => {
          const rowData = [];
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            rowData[colNumber - 1] = cell.formula || null;
          });
          formulas[rowNumber - 1] = rowData;
        });
      }
      
      return { 
        result: { formulas }
      };
    } catch (error) {
      console.error("Error reading sheet formulas:", error);
      return { error: { message: `Error reading sheet formulas: ${error.message}` } };
    }
  }
});

// Write Sheet Data Tool
server.addTool({
  name: "write_sheet_data",
  description: "Write data to the Excel sheet",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      },
      sheetName: {
        type: "string",
        description: "Sheet name in the Excel file"
      },
      range: {
        type: "string",
        description: "Range of cells in the Excel sheet (e.g., \"A1:C10\")"
      },
      data: {
        type: "array",
        items: {
          type: "array",
          items: {
            type: ["string", "number", "boolean", "null"]
          }
        },
        description: "Data to write to the Excel sheet"
      }
    },
    required: ["fileAbsolutePath", "sheetName", "range", "data"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath, sheetName, range, data } = params;
      console.error(`Writing data to: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`);
      
      let workbook = new ExcelJS.Workbook();
      
      try {
        await workbook.xlsx.readFile(fileAbsolutePath);
      } catch (error) {
        // File doesn't exist or can't be read
        console.error(`Creating new workbook: ${error.message}`);
      }
      
      let worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
      }
      
      // Parse range - we only need the starting cell for writing
      const rangeMatch = range.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/);
      if (!rangeMatch) {
        return { error: { message: `Invalid range format: ${range}` } };
      }
      
      const [_, startCol, startRow] = rangeMatch;
      
      // Convert column letters to numbers
      const getColumnNumber = (col) => {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
          num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
      };
      
      const startColNum = getColumnNumber(startCol);
      const startRowNum = parseInt(startRow, 10);
      
      // Write data
      data.forEach((rowData, rowIndex) => {
        rowData.forEach((cellValue, colIndex) => {
          // Convert column number back to letter
          const getColumnLetter = (num) => {
            let letter = '';
            while (num > 0) {
              const remainder = (num - 1) % 26;
              letter = String.fromCharCode(65 + remainder) + letter;
              num = Math.floor((num - 1) / 26);
            }
            return letter;
          };
          
          const colLetter = getColumnLetter(startColNum + colIndex);
          const rowNum = startRowNum + rowIndex;
          worksheet.getCell(`${colLetter}${rowNum}`).value = cellValue;
        });
      });
      
      // Save workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return { 
        result: { 
          message: `Successfully wrote data to ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`
        }
      };
    } catch (error) {
      console.error("Error writing sheet data:", error);
      return { error: { message: `Error writing sheet data: ${error.message}` } };
    }
  }
});

// Write Sheet Formula Tool
server.addTool({
  name: "write_sheet_formula",
  description: "Write formulas to the Excel sheet",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      },
      sheetName: {
        type: "string",
        description: "Sheet name in the Excel file"
      },
      range: {
        type: "string",
        description: "Range of cells in the Excel sheet (e.g., \"A1:C10\")"
      },
      formulas: {
        type: "array",
        items: {
          type: "array",
          items: {
            type: "string"
          }
        },
        description: "Formulas to write to the Excel sheet (e.g., \"=A1+B1\")"
      }
    },
    required: ["fileAbsolutePath", "sheetName", "range", "formulas"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath, sheetName, range, formulas } = params;
      console.error(`Writing formulas to: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`);
      
      let workbook = new ExcelJS.Workbook();
      
      try {
        await workbook.xlsx.readFile(fileAbsolutePath);
      } catch (error) {
        // File doesn't exist or can't be read
        console.error(`Creating new workbook: ${error.message}`);
      }
      
      let worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
      }
      
      // Parse range - we only need the starting cell for writing
      const rangeMatch = range.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/);
      if (!rangeMatch) {
        return { error: { message: `Invalid range format: ${range}` } };
      }
      
      const [_, startCol, startRow] = rangeMatch;
      
      // Convert column letters to numbers
      const getColumnNumber = (col) => {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
          num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
      };
      
      const startColNum = getColumnNumber(startCol);
      const startRowNum = parseInt(startRow, 10);
      
      // Write formulas
      formulas.forEach((rowData, rowIndex) => {
        rowData.forEach((formula, colIndex) => {
          if (formula) {
            // Convert column number back to letter
            const getColumnLetter = (num) => {
              let letter = '';
              while (num > 0) {
                const remainder = (num - 1) % 26;
                letter = String.fromCharCode(65 + remainder) + letter;
                num = Math.floor((num - 1) / 26);
              }
              return letter;
            };
            
            const colLetter = getColumnLetter(startColNum + colIndex);
            const rowNum = startRowNum + rowIndex;
            worksheet.getCell(`${colLetter}${rowNum}`).value = { 
              formula: formula.startsWith('=') ? formula.substring(1) : formula 
            };
          }
        });
      });
      
      // Save workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return { 
        result: { 
          message: `Successfully wrote formulas to ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`
        }
      };
    } catch (error) {
      console.error("Error writing sheet formulas:", error);
      return { error: { message: `Error writing sheet formulas: ${error.message}` } };
    }
  }
});

// Format Cells Tool
server.addTool({
  name: "format_cells",
  description: "Format cells in Excel sheet with colors, fonts, borders, etc.",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      },
      sheetName: {
        type: "string",
        description: "Sheet name in the Excel file"
      },
      range: {
        type: "string",
        description: "Range of cells to format (e.g., \"A1:C10\")"
      },
      format: {
        type: "object",
        properties: {
          bold: {
            type: "boolean",
            description: "Make text bold"
          },
          italic: {
            type: "boolean",
            description: "Make text italic"
          },
          underline: {
            type: "boolean",
            description: "Underline text"
          },
          fontSize: {
            type: "number",
            description: "Font size"
          },
          fontName: {
            type: "string",
            description: "Font name"
          },
          fontColor: {
            type: "string",
            description: "Font color (hex code e.g., \"#FF0000\")"
          },
          backgroundColor: {
            type: "string",
            description: "Background color (hex code e.g., \"#FFFF00\")"
          },
          horizontalAlignment: {
            type: "string",
            enum: ["left", "center", "right", "fill", "justify", "centerContinuous", "distributed"],
            description: "Horizontal alignment"
          },
          verticalAlignment: {
            type: "string",
            enum: ["top", "middle", "bottom", "distributed", "justify"],
            description: "Vertical alignment"
          },
          wrapText: {
            type: "boolean",
            description: "Wrap text"
          },
          numberFormat: {
            type: "string",
            description: "Number format (e.g., \"0.00\", \"0%\", \"m/d/yy\")"
          }
        },
        description: "Formatting options"
      }
    },
    required: ["fileAbsolutePath", "sheetName", "range", "format"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath, sheetName, range, format } = params;
      console.error(`Formatting cells in: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`);
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return { error: { message: `Sheet "${sheetName}" not found in workbook.` } };
      }
      
      // Parse range
      const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        return { error: { message: `Invalid range format: ${range}` } };
      }
      
      const [_, startCol, startRow, endCol, endRow] = rangeMatch;
      
      // Convert column letters to numbers
      const getColumnNumber = (col) => {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
          num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
      };
      
      const startColNum = getColumnNumber(startCol);
      const endColNum = getColumnNumber(endCol);
      const startRowNum = parseInt(startRow, 10);
      const endRowNum = parseInt(endRow, 10);
      
      // Apply formatting to cells in range
      for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
        for (let colNum = startColNum; colNum <= endColNum; colNum++) {
          // Convert column number back to letter
          const getColumnLetter = (num) => {
            let letter = '';
            while (num > 0) {
              const remainder = (num - 1) % 26;
              letter = String.fromCharCode(65 + remainder) + letter;
              num = Math.floor((num - 1) / 26);
            }
            return letter;
          };
          
          const colLetter = getColumnLetter(colNum);
          const cell = worksheet.getCell(`${colLetter}${rowNum}`);
          
          // Apply font formatting
          if (format.bold !== undefined || 
              format.italic !== undefined || 
              format.underline !== undefined || 
              format.fontSize !== undefined ||
              format.fontName !== undefined ||
              format.fontColor !== undefined) {
            
            const fontOptions = {};
            
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
            
            const alignmentOptions = {};
            
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
      
      // Save workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return { 
        result: { 
          message: `Successfully formatted cells in ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`
        }
      };
    } catch (error) {
      console.error("Error formatting cells:", error);
      return { error: { message: `Error formatting cells: ${error.message}` } };
    }
  }
});

// Add Borders Tool
server.addTool({
  name: "add_borders",
  description: "Add borders to cells in Excel sheet",
  parameters: {
    type: "object",
    properties: {
      fileAbsolutePath: {
        type: "string",
        description: "Absolute path to the Excel file"
      },
      sheetName: {
        type: "string",
        description: "Sheet name in the Excel file"
      },
      range: {
        type: "string",
        description: "Range of cells to add borders to (e.g., \"A1:C10\")"
      },
      borderStyle: {
        type: "object",
        properties: {
          top: {
            type: "object",
            properties: {
              style: {
                type: "string",
                enum: ["thin", "medium", "thick", "dotted", "dashed", "double"],
                description: "Border style"
              },
              color: {
                type: "string",
                description: "Border color (hex code)"
              }
            },
            description: "Top border style"
          },
          bottom: {
            type: "object",
            properties: {
              style: {
                type: "string",
                enum: ["thin", "medium", "thick", "dotted", "dashed", "double"],
                description: "Border style"
              },
              color: {
                type: "string",
                description: "Border color (hex code)"
              }
            },
            description: "Bottom border style"
          },
          left: {
            type: "object",
            properties: {
              style: {
                type: "string",
                enum: ["thin", "medium", "thick", "dotted", "dashed", "double"],
                description: "Border style"
              },
              color: {
                type: "string",
                description: "Border color (hex code)"
              }
            },
            description: "Left border style"
          },
          right: {
            type: "object",
            properties: {
              style: {
                type: "string",
                enum: ["thin", "medium", "thick", "dotted", "dashed", "double"],
                description: "Border style"
              },
              color: {
                type: "string",
                description: "Border color (hex code)"
              }
            },
            description: "Right border style"
          },
          outline: {
            type: "boolean",
            description: "Add outline borders to the range"
          },
          all: {
            type: "object",
            properties: {
              style: {
                type: "string",
                enum: ["thin", "medium", "thick", "dotted", "dashed", "double"],
                description: "Border style"
              },
              color: {
                type: "string",
                description: "Border color (hex code)"
              }
            },
            description: "Apply the same border style to all sides"
          }
        },
        description: "Border style options"
      }
    },
    required: ["fileAbsolutePath", "sheetName", "range", "borderStyle"]
  },
  async execute(params) {
    try {
      const { fileAbsolutePath, sheetName, range, borderStyle } = params;
      console.error(`Adding borders to cells in: ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`);
      
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(fileAbsolutePath);
      
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return { error: { message: `Sheet "${sheetName}" not found in workbook.` } };
      }
      
      // Parse range
      const rangeMatch = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (!rangeMatch) {
        return { error: { message: `Invalid range format: ${range}` } };
      }
      
      const [_, startCol, startRow, endCol, endRow] = rangeMatch;
      
      // Convert column letters to numbers
      const getColumnNumber = (col) => {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
          num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
      };
      
      const startColNum = getColumnNumber(startCol);
      const endColNum = getColumnNumber(endCol);
      const startRowNum = parseInt(startRow, 10);
      const endRowNum = parseInt(endRow, 10);
      
      // Helper function to convert border style
      const convertBorderStyle = (style) => {
        if (!style) return undefined;
        
        return {
          style: style.style,
          color: style.color ? { argb: style.color.replace('#', '') } : undefined
        };
      };
      
      // Apply "all" border style to all sides if specified
      const allBorderStyle = borderStyle.all ? convertBorderStyle(borderStyle.all) : undefined;
      
      // Apply borders to cells in range
      for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
        for (let colNum = startColNum; colNum <= endColNum; colNum++) {
          // Convert column number back to letter
          const getColumnLetter = (num) => {
            let letter = '';
            while (num > 0) {
              const remainder = (num - 1) % 26;
              letter = String.fromCharCode(65 + remainder) + letter;
              num = Math.floor((num - 1) / 26);
            }
            return letter;
          };
          
          const colLetter = getColumnLetter(colNum);
          const cell = worksheet.getCell(`${colLetter}${rowNum}`);
          
          const cellBorder = {};
          
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
      
      // Save workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return { 
        result: { 
          message: `Successfully added borders to cells in ${fileAbsolutePath}, Sheet: ${sheetName}, Range: ${range}`
        }
      };
    } catch (error) {
      console.error("Error adding borders to cells:", error);
      return { error: { message: `Error adding borders to cells: ${error.message}` } };
    }
  }
});

// Start the server
server.run().then(() => {
  console.error("Excel MCP Server started in production mode (stdin/stdout)");
}).catch((error) => {
  console.error("Failed to start Excel MCP Server:", error);
});
