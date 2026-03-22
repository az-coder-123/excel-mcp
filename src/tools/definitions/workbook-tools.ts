/**
 * Workbook Tool Definitions
 * Single Responsibility: Workbook lifecycle operations
 */

import { ToolDefinition } from '../../types/index.js';

export const WORKBOOK_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_open_workbook',
    description: 'Open Excel file for operations',
    parameters: [
      {
        name: 'filePath',
        type: 'string',
        description: 'Absolute path to Excel file (.xlsx, .xls)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_create_workbook',
    description: 'Create new workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name (without extension)',
        required: true,
      },
      {
        name: 'outputPath',
        type: 'string',
        description: 'Directory path for saving',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_save_workbook',
    description: 'Save workbook to disk',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'outputPath',
        type: 'string',
        description: 'Optional: Different save path',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_close_workbook',
    description: 'Close workbook, free resources',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_export_worksheet_to_new_file',
    description: 'Export worksheet to new Excel file',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name to export',
        required: true,
      },
      {
        name: 'newFilePath',
        type: 'string',
        description: 'Full path for new file (e.g., /path/to/newfile.xlsx)',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_get_workbook_context',
    description: 'Get workbook context and state',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name to check',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
];
