import { FastMCP, UserError } from '@modelcontextprotocol/sdk';
import { z } from 'zod'; // Using Zod for schema validation
import * as tools from './tools';

// Create the MCP server instance
const server = new FastMCP({
  name: "Excel MCP with Formatting",
  version: "1.0.0",
  description: "An MCP server for Excel with rich formatting capabilities"
});

// Register our tools with proper parameter validation using Zod
server.addTool({
  name: "read_sheet_data",
  description: "Read data from Excel sheet with pagination.",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().optional().describe('Range of cells to read in the Excel sheet (e.g., "A1:C10")'),
    knownPagingRanges: z.array(z.string()).optional().describe('List of already read paging ranges')
  }),
  execute: tools.readSheetDataHandler
});

server.addTool({
  name: "read_sheet_formula",
  description: "Read formulas from Excel sheet with pagination.",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().optional().describe('Range of cells to read in the Excel sheet (e.g., "A1:C10")'),
    knownPagingRanges: z.array(z.string()).optional().describe('List of already read paging ranges')
  }),
  execute: tools.readSheetFormulaHandler
});

server.addTool({
  name: "write_sheet_data",
  description: "Write data to the Excel sheet",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells in the Excel sheet (e.g., "A1:C10")'),
    data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
      .describe('Data to write to the Excel sheet')
  }),
  execute: tools.writeSheetDataHandler
});

server.addTool({
  name: "read_sheet_names",
  description: "List all sheet names in an Excel file",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file')
  }),
  execute: tools.readSheetNamesHandler
});

server.addTool({
  name: "format_cells",
  description: "Format cells in Excel sheet with colors, fonts, borders, etc.",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells to format (e.g., "A1:C10")'),
    format: z.object({
      bold: z.boolean().optional().describe('Make text bold'),
      italic: z.boolean().optional().describe('Make text italic'),
      underline: z.boolean().optional().describe('Underline text'),
      fontSize: z.number().optional().describe('Font size'),
      fontName: z.string().optional().describe('Font name'),
      fontColor: z.string().optional().describe('Font color (hex code e.g., "#FF0000")'),
      backgroundColor: z.string().optional().describe('Background color (hex code e.g., "#FFFF00")'),
      horizontalAlignment: z.enum(['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'])
        .optional().describe('Horizontal alignment'),
      verticalAlignment: z.enum(['top', 'middle', 'bottom', 'distributed', 'justify'])
        .optional().describe('Vertical alignment'),
      wrapText: z.boolean().optional().describe('Wrap text'),
      numberFormat: z.string().optional().describe('Number format (e.g., "0.00", "0%", "m/d/yy")')
    }).describe('Formatting options')
  }),
  execute: tools.formatCellsHandler
});

server.addTool({
  name: "add_borders",
  description: "Add borders to cells in Excel sheet",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells to add borders to (e.g., "A1:C10")'),
    borderStyle: z.object({
      top: z.object({
        style: z.enum(['thin', 'medium', 'thick', 'dotted', 'dashed', 'double']).optional(),
        color: z.string().optional()
      }).optional(),
      bottom: z.object({
        style: z.enum(['thin', 'medium', 'thick', 'dotted', 'dashed', 'double']).optional(),
        color: z.string().optional()
      }).optional(),
      left: z.object({
        style: z.enum(['thin', 'medium', 'thick', 'dotted', 'dashed', 'double']).optional(),
        color: z.string().optional()
      }).optional(),
      right: z.object({
        style: z.enum(['thin', 'medium', 'thick', 'dotted', 'dashed', 'double']).optional(),
        color: z.string().optional()
      }).optional(),
      outline: z.boolean().optional(),
      all: z.object({
        style: z.enum(['thin', 'medium', 'thick', 'dotted', 'dashed', 'double']),
        color: z.string().optional()
      }).optional()
    }).describe('Border style options')
  }),
  execute: tools.addBordersHandler
});

server.addTool({
  name: "merge_cells",
  description: "Merge cells in Excel sheet",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells to merge (e.g., "A1:C10")'),
    alignContent: z.object({
      horizontal: z.enum(['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']).optional(),
      vertical: z.enum(['top', 'middle', 'bottom', 'distributed', 'justify']).optional()
    }).optional().describe('Alignment options for merged content')
  }),
  execute: tools.mergeCellsHandler
});

server.addTool({
  name: "unmerge_cells",
  description: "Unmerge previously merged cells in Excel sheet",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    range: z.string().describe('Range of cells to unmerge (e.g., "A1:C10")')
  }),
  execute: tools.unmergeCellsHandler
});

server.addTool({
  name: "create_workbook",
  description: "Create a new Excel workbook",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to create the Excel file (including filename.xlsx)'),
    initialSheets: z.array(z.string()).optional().describe('Names of initial sheets to create (default: ["Sheet1"])'),
    creator: z.string().optional().describe('Creator of the workbook (metadata)'),
    lastModifiedBy: z.string().optional().describe('Last modified by (metadata)'),
    title: z.string().optional().describe('Workbook title (metadata)'),
    subject: z.string().optional().describe('Workbook subject (metadata)'),
    keywords: z.string().optional().describe('Workbook keywords (metadata)'),
    category: z.string().optional().describe('Workbook category (metadata)'),
    description: z.string().optional().describe('Workbook description (metadata)'),
    overwrite: z.boolean().optional().describe('Overwrite if file exists (default: false)')
  }),
  execute: tools.createWorkbookHandler
});

server.addTool({
  name: "add_worksheet",
  description: "Add a new worksheet to an Excel workbook",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Name for the new worksheet'),
    tabColor: z.string().optional().describe('Tab color (hex code e.g., "#FF0000")'),
    columns: z.array(
      z.object({
        header: z.string().describe('Column header text'),
        key: z.string().describe('Column key (for referencing in data)'),
        width: z.number().optional().describe('Column width')
      })
    ).optional().describe('Column definitions'),
    headerRowFormat: z.object({
      bold: z.boolean().optional(),
      fontSize: z.number().optional(),
      fontColor: z.string().optional(),
      backgroundColor: z.string().optional()
    }).optional().describe('Formatting for the header row (if columns provided)'),
    position: z.number().optional().describe('Position of the new sheet (0-based index)')
  }),
  execute: tools.addWorksheetHandler
});

server.addTool({
  name: "apply_styles",
  description: "Apply cell styles including conditional formatting to Excel",
  parameters: z.object({
    fileAbsolutePath: z.string().describe('Absolute path to the Excel file'),
    sheetName: z.string().describe('Sheet name in the Excel file'),
    actions: z.array(
      z.object({
        type: z.enum([
          'alternating_rows', 
          'table_style', 
          'header_row', 
          'total_row',
          'banded_columns',
          'highlight_negative',
          'highlight_positive',
          'highlight_max',
          'highlight_min',
          'data_bars',
          'gradient_scale',
          'icon_set'
        ]).describe('Type of styling to apply'),
        range: z.string().describe('Range to apply the styling to (e.g., "A1:C10")'),
        properties: z.record(z.any()).optional().describe('Style properties specific to the selected style type')
      })
    ).describe('List of styling actions to apply')
  }),
  execute: tools.applyStylesHandler
});

// Set up error handler
server.setErrorHandler((error) => {
  if (error instanceof UserError) {
    return { error: error.message };
  }
  
  // Log internal errors but return a generic message
  console.error('Internal error:', error);
  return { error: 'An internal error occurred. Please check your input parameters and try again.' };
});

// Start the server
const port = process.env.FASTMCP_PORT ? parseInt(process.env.FASTMCP_PORT) : 8000;

// Use the appropriate server transport
if (process.env.NODE_ENV === 'development') {
  // For development, use SSE server
  server.run_sse_async(port).then(() => {
    console.log(`Excel MCP Server with Formatting is running on port ${port}`);
  }).catch(error => {
    console.error('Error starting MCP server:', error);
  });
} else {
  // For production, use stdin/stdout
  server.run().then(() => {
    console.log(`Excel MCP Server with Formatting is running with stdin/stdout transport`);
  }).catch(error => {
    console.error('Error starting MCP server:', error);
  });
}
