import { Tool } from '@modelcontextprotocol/sdk';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

export const createWorkbook: Tool = {
  name: 'create_workbook',
  description: 'Create a new Excel workbook',
  parameters: {
    type: 'object',
    properties: {
      fileAbsolutePath: {
        type: 'string',
        description: 'Absolute path to create the Excel file (including filename.xlsx)'
      },
      initialSheets: {
        type: 'array',
        items: {
          type: 'string'
        },
        description: 'Names of initial sheets to create (default: ["Sheet1"])'
      },
      creator: {
        type: 'string',
        description: 'Creator of the workbook (metadata)'
      },
      lastModifiedBy: {
        type: 'string',
        description: 'Last modified by (metadata)'
      },
      title: {
        type: 'string',
        description: 'Workbook title (metadata)'
      },
      subject: {
        type: 'string',
        description: 'Workbook subject (metadata)'
      },
      keywords: {
        type: 'string',
        description: 'Workbook keywords (metadata)'
      },
      category: {
        type: 'string',
        description: 'Workbook category (metadata)'
      },
      description: {
        type: 'string',
        description: 'Workbook description (metadata)'
      },
      overwrite: {
        type: 'boolean',
        description: 'Overwrite if file exists (default: false)'
      }
    },
    required: ['fileAbsolutePath']
  },
  handler: async ({ 
    fileAbsolutePath, 
    initialSheets = ['Sheet1'], 
    creator, 
    lastModifiedBy,
    title,
    subject,
    keywords,
    category,
    description,
    overwrite = false 
  }) => {
    try {
      // Check if file exists
      if (fs.existsSync(fileAbsolutePath) && !overwrite) {
        throw new Error(`File already exists. Use 'overwrite: true' to replace it.`);
      }
      
      // Create directory if it doesn't exist
      const directory = path.dirname(fileAbsolutePath);
      if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory, { recursive: true });
      }
      
      const workbook = new ExcelJS.Workbook();
      
      // Set properties
      if (creator) workbook.creator = creator;
      if (lastModifiedBy) workbook.lastModifiedBy = lastModifiedBy;
      if (title) workbook.title = title;
      if (subject) workbook.subject = subject;
      if (keywords) workbook.keywords = keywords;
      if (category) workbook.category = category;
      if (description) workbook.description = description;
      
      // Create initial sheets
      initialSheets.forEach(sheetName => {
        workbook.addWorksheet(sheetName);
      });
      
      // Save workbook
      await workbook.xlsx.writeFile(fileAbsolutePath);
      
      return {
        status: 'success',
        message: `Workbook created at ${fileAbsolutePath}`,
        sheets: initialSheets,
        filepath: fileAbsolutePath
      };
    } catch (error) {
      if (error instanceof Error) {
        throw new Error(`Failed to create workbook: ${error.message}`);
      }
      throw error;
    }
  }
};
