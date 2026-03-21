/**
 * Tool Definitions - Defines all available MCP tools
 * Single Responsibility: Tool schema definitions only
 */

import { ToolDefinition, Permission } from '../types/index.js';

export const TOOL_DEFINITIONS: ToolDefinition[] = [
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
    name: 'excel_read_cell',
    description: 'Read the value of a specific cell in an Excel workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_read_range',
    description: 'Read values from a range of cells',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell of the range (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of the range (e.g., C10)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_write_cell',
    description: 'Write a value to a specific cell',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1, B2)',
        required: true,
      },
      {
        name: 'value',
        type: 'string',
        description: 'Value to write (can be string, number, or formula starting with =)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_list_worksheets',
    description: 'List all worksheets in an opened workbook',
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
  {
    name: 'excel_add_worksheet',
    description: 'Add a new worksheet to an opened workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheetName',
        type: 'string',
        description: 'Name for the new worksheet',
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
  {
    name: 'excel_get_cell_info',
    description: 'Get detailed information about a cell including type and formula',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
];

/**
 * Get tool definition by name
 */
export function getToolDefinition(name: string): ToolDefinition | undefined {
  return TOOL_DEFINITIONS.find((tool) => tool.name === name);
}

/**
 * Get all tool names
 */
export function getAllToolNames(): string[] {
  return TOOL_DEFINITIONS.map((tool) => tool.name);
}

/**
 * Get tools by required permission
 */
export function getToolsByPermission(permission: Permission): ToolDefinition[] {
  return TOOL_DEFINITIONS.filter((tool) => 
    tool.requiredPermissions.includes(permission)
  );
}