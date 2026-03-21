/**
 * Worksheet Tool Definitions
 * Single Responsibility: Worksheet and structure operations
 */

import { ToolDefinition } from '../../types/index.js';

export const WORKSHEET_TOOLS: ToolDefinition[] = [
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
    name: 'excel_delete_worksheet',
    description: 'Delete a worksheet from the workbook',
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
        description: 'Name of the worksheet to delete',
        required: true,
      },
    ],
    requiredPermissions: ['delete'],
  },
  {
    name: 'excel_rename_worksheet',
    description: 'Rename a worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'oldName',
        type: 'string',
        description: 'Current worksheet name',
        required: true,
      },
      {
        name: 'newName',
        type: 'string',
        description: 'New worksheet name',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_copy_worksheet',
    description: 'Copy a worksheet within the workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'sourceWorksheet',
        type: 'string',
        description: 'Source worksheet name',
        required: true,
      },
      {
        name: 'targetName',
        type: 'string',
        description: 'Name for the copied worksheet',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_insert_rows',
    description: 'Insert rows at a specific position',
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
        name: 'startRow',
        type: 'number',
        description: 'Row number to start insertion (1-based)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of rows to insert',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_insert_columns',
    description: 'Insert columns at a specific position',
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
        name: 'startColumn',
        type: 'string',
        description: 'Column letter to start insertion (A, B, C...)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of columns to insert',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_rows',
    description: 'Delete rows from a specific position',
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
        name: 'startRow',
        type: 'number',
        description: 'Row number to start deletion (1-based)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of rows to delete',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_columns',
    description: 'Delete columns from a specific position',
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
        name: 'startColumn',
        type: 'string',
        description: 'Column letter to start deletion (A, B, C...)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of columns to delete',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_merge_cells',
    description: 'Merge a range of cells',
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
        description: 'Start cell of range (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of range (e.g., C3)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_unmerge_cells',
    description: 'Unmerge a merged cell range',
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
        description: 'Any cell within the merged range',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_table',
    description: 'Create an Excel table with headers',
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
        description: 'Start cell of table (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of table (e.g., D10)',
        required: true,
      },
      {
        name: 'tableName',
        type: 'string',
        description: 'Name for the table',
        required: true,
      },
      {
        name: 'style',
        type: 'string',
        description: 'Table style name',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_filter',
    description: 'Add auto-filter to a range',
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
        description: 'Start cell of range (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of range (e.g., D100)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_filter',
    description: 'Remove auto-filter from worksheet',
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
    ],
    requiredPermissions: ['write'],
  },
];