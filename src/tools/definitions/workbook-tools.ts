/**
 * Workbook Tool Definitions
 * Single Responsibility: Workbook lifecycle operations
 */

import { ToolDefinition } from '../../types/index.js';

export const WORKBOOK_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_open_workbook',
    description: 'Open an Excel workbook file for reading and writing operations',
    parameters: [
      {
        name: 'filePath',
        type: 'string',
        description: 'Absolute path to the Excel file (.xlsx, .xls)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_create_workbook',
    description: 'Create a new Excel workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name for the new workbook (without extension)',
        required: true,
      },
      {
        name: 'outputPath',
        type: 'string',
        description: 'Directory path where the workbook will be saved',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_save_workbook',
    description: 'Save the workbook to disk',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'outputPath',
        type: 'string',
        description: 'Optional: Different path to save the file',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_close_workbook',
    description: 'Close an opened workbook and free resources',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
];