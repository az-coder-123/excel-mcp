/**
 * Advanced Data Tool Definitions
 * Single Responsibility: Advanced data operations like pivot, import/export, and duplicate detection
 */

import { ToolDefinition } from '../../types/index.js';

export const ADVANCED_DATA_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_find_duplicates',
    description: 'Find duplicates in range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
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
        description: 'Range start',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_count_unique_values',
    description: 'Count unique/duplicate values',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
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
        description: 'Range start',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_highlight_duplicates',
    description: 'Highlight duplicates with color',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
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
        description: 'Range start',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end',
        required: true,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Color hex',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_get_duplicate_info',
    description: 'Get duplicate details for column',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'column',
        type: 'string',
        description: 'Column letter',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_import_csv',
    description: 'Import CSV to worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'csvContent',
        type: 'string',
        description: 'CSV content',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Import start cell',
        required: true,
      },
      {
        name: 'delimiter',
        type: 'string',
        description: 'Delimiter',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_export_csv',
    description: 'Export range to CSV',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
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
        description: 'Range start',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end',
        required: true,
      },
      {
        name: 'delimiter',
        type: 'string',
        description: 'Delimiter',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_create_pivot_table',
    description: 'Create pivot table',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'sourceWorksheet',
        type: 'string',
        description: 'Source worksheet',
        required: true,
      },
      {
        name: 'sourceStartCell',
        type: 'string',
        description: 'Source start',
        required: true,
      },
      {
        name: 'sourceEndCell',
        type: 'string',
        description: 'Source end',
        required: true,
      },
      {
        name: 'targetWorksheet',
        type: 'string',
        description: 'Target worksheet',
        required: true,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Placement cell',
        required: true,
      },
      {
        name: 'rowFields',
        type: 'array',
        description: 'Row fields',
        required: false,
        items: {
          name: 'fieldName',
          type: 'string',
          description: 'Field',
          required: true
        },
      },
      {
        name: 'columnFields',
        type: 'array',
        description: 'Column fields',
        required: false,
        items: {
          name: 'fieldName',
          type: 'string',
          description: 'Field',
          required: true
        },
      },
      {
        name: 'valueFields',
        type: 'array',
        description: 'Value fields',
        required: false,
        items: {
          name: 'fieldName',
          type: 'string',
          description: 'Field',
          required: true
        },
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_vlookup',
    description: 'VLOOKUP formula',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Target cell',
        required: true,
      },
      {
        name: 'lookupValue',
        type: 'string',
        description: 'Lookup value',
        required: true,
      },
      {
        name: 'tableArray',
        type: 'string',
        description: 'Table range',
        required: true,
      },
      {
        name: 'colIndex',
        type: 'number',
        description: 'Column index',
        required: true,
      },
      {
        name: 'rangeLookup',
        type: 'boolean',
        description: 'Approximate match',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_index_match',
    description: 'INDEX/MATCH formula',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Target cell',
        required: true,
      },
      {
        name: 'returnRange',
        type: 'string',
        description: 'Return range',
        required: true,
      },
      {
        name: 'lookupRange',
        type: 'string',
        description: 'Lookup range',
        required: true,
      },
      {
        name: 'lookupValue',
        type: 'string',
        description: 'Lookup value',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_duplicates',
    description: 'Remove duplicates',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
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
        description: 'Range start',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end',
        required: true,
      },
      {
        name: 'columns',
        type: 'array',
        description: 'Columns to check',
        required: false,
        items: {
          name: 'columnIndex',
          type: 'number',
          description: 'Index',
          required: true
        },
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_text_to_columns',
    description: 'Split text to columns',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'sourceCell',
        type: 'string',
        description: 'Source cell',
        required: true,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Target start',
        required: true,
      },
      {
        name: 'delimiter',
        type: 'string',
        description: 'Delimiter',
        required: true,
      },
      {
        name: 'numberOfColumns',
        type: 'number',
        description: 'Columns count',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_flash_fill',
    description: 'Apply flash fill',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'sourceRange',
        type: 'string',
        description: 'Source range',
        required: true,
      },
      {
        name: 'targetRange',
        type: 'string',
        description: 'Target range',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_split_data_to_worksheets',
    description: 'Split data to worksheets',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'sourceWorksheet',
        type: 'string',
        description: 'Source worksheet',
        required: true,
      },
      {
        name: 'columnLetter',
        type: 'string',
        description: 'Filter column',
        required: true,
      },
      {
        name: 'sheetNames',
        type: 'object',
        description: 'Value-to-sheet mapping',
        required: false,
      },
      {
        name: 'includeHeader',
        type: 'boolean',
        description: 'Include header',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_get_unique_values',
    description: 'Get unique values',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'columnLetter',
        type: 'string',
        description: 'Column letter',
        required: true,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Has header',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_filter_data',
    description: 'Filter data to new sheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'sourceWorksheet',
        type: 'string',
        description: 'Source worksheet',
        required: true,
      },
      {
        name: 'targetWorksheet',
        type: 'string',
        description: 'Target worksheet',
        required: true,
      },
      {
        name: 'filters',
        type: 'array',
        description: 'Filter conditions',
        required: true,
        items: {
          name: 'column',
          type: 'string',
          description: 'Column',
          required: true
        },
      },
      {
        name: 'sourceRange',
        type: 'string',
        description: 'Source range',
        required: false,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Target cell',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
];