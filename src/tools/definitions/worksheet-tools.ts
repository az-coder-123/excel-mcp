/**
 * Worksheet Tool Definitions
 * Single Responsibility: Worksheet and structure operations
 */

import { ToolDefinition } from '../../types/index.js';

export const WORKSHEET_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_list_worksheets',
    description: 'List all worksheets',
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
    name: 'excel_add_worksheet',
    description: 'Add new worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheetName',
        type: 'string',
        description: 'New worksheet name',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_worksheet',
    description: 'Delete worksheet',
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
        description: 'Worksheet to delete',
        required: true,
      },
    ],
    requiredPermissions: ['delete'],
  },
  {
    name: 'excel_rename_worksheet',
    description: 'Rename worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'oldName',
        type: 'string',
        description: 'Current name',
        required: true,
      },
      {
        name: 'newName',
        type: 'string',
        description: 'New name',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_copy_worksheet',
    description: 'Copy worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'sourceWorksheet',
        type: 'string',
        description: 'Source worksheet',
        required: true,
      },
      {
        name: 'targetName',
        type: 'string',
        description: 'Copy name',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_insert_rows',
    description: 'Insert rows',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startRow',
        type: 'number',
        description: 'Start row (1-based)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of rows',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_insert_columns',
    description: 'Insert columns',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startColumn',
        type: 'string',
        description: 'Start column (A, B, C...)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of columns',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_rows',
    description: 'Delete rows',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startRow',
        type: 'number',
        description: 'Start row (1-based)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of rows',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_columns',
    description: 'Delete columns',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startColumn',
        type: 'string',
        description: 'Start column (A, B, C...)',
        required: true,
      },
      {
        name: 'count',
        type: 'number',
        description: 'Number of columns',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_merge_cells',
    description: 'Merge cell range',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Range start (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end (e.g., C3)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_unmerge_cells',
    description: 'Unmerge cells',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell in merged range',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_table',
    description: 'Create Excel table',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Table start (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Table end (e.g., D10)',
        required: true,
      },
      {
        name: 'tableName',
        type: 'string',
        description: 'Table name',
        required: true,
      },
      {
        name: 'style',
        type: 'string',
        description: 'Table style',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_filter',
    description: 'Add auto-filter',
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
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Range start (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end (e.g., D100)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_filter',
    description: 'Remove auto-filter',
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
        description: 'Worksheet name',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
];