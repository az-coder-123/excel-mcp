/**
 * Advanced Data Tool Definitions
 * Single Responsibility: Advanced data operations like pivot, import/export, and duplicate detection
 */

import { ToolDefinition } from '../../types/index.js';

export const ADVANCED_DATA_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_find_duplicates',
    description: 'Find duplicate values in a range',
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
        description: 'End cell of range (e.g., D10)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_count_unique_values',
    description: 'Count unique and duplicate values in a range',
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
        description: 'End cell of range (e.g., D10)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_highlight_duplicates',
    description: 'Highlight duplicate values in a range with color',
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
        description: 'End cell of range (e.g., D10)',
        required: true,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Highlight color in hex format (e.g., FFFF0000 for red)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_get_duplicate_info',
    description: 'Get detailed duplicate information for a specific column',
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
        name: 'column',
        type: 'string',
        description: 'Column letter (e.g., A, B, C)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_import_csv',
    description: 'Import CSV data into worksheet',
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
        name: 'csvContent',
        type: 'string',
        description: 'CSV content to import',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell for import (e.g., A1)',
        required: true,
      },
      {
        name: 'delimiter',
        type: 'string',
        description: 'CSV delimiter',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_export_csv',
    description: 'Export worksheet data to CSV format',
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
        description: 'Start cell of range',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of range',
        required: true,
      },
      {
        name: 'delimiter',
        type: 'string',
        description: 'CSV delimiter',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_create_pivot_table',
    description: 'Create a pivot table from data range',
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
        description: 'Source worksheet with data',
        required: true,
      },
      {
        name: 'sourceStartCell',
        type: 'string',
        description: 'Start cell of source data',
        required: true,
      },
      {
        name: 'sourceEndCell',
        type: 'string',
        description: 'End cell of source data',
        required: true,
      },
      {
        name: 'targetWorksheet',
        type: 'string',
        description: 'Target worksheet for pivot table',
        required: true,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Cell for pivot table placement',
        required: true,
      },
      {
        name: 'rowFields',
        type: 'array',
        description: 'Array of field names for rows',
        required: false,
        items: {
          name: 'fieldName',
          type: 'string',
          description: 'Field name',
          required: true
        },
      },
      {
        name: 'columnFields',
        type: 'array',
        description: 'Array of field names for columns',
        required: false,
        items: {
          name: 'fieldName',
          type: 'string',
          description: 'Field name',
          required: true
        },
      },
      {
        name: 'valueFields',
        type: 'array',
        description: 'Array of field names for values',
        required: false,
        items: {
          name: 'fieldName',
          type: 'string',
          description: 'Field name',
          required: true
        },
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_vlookup',
    description: 'Perform VLOOKUP operation',
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
        name: 'targetCell',
        type: 'string',
        description: 'Cell to write the formula',
        required: true,
      },
      {
        name: 'lookupValue',
        type: 'string',
        description: 'Value to lookup',
        required: true,
      },
      {
        name: 'tableArray',
        type: 'string',
        description: 'Table array range (e.g., A1:D10)',
        required: true,
      },
      {
        name: 'colIndex',
        type: 'number',
        description: 'Column index to return',
        required: true,
      },
      {
        name: 'rangeLookup',
        type: 'boolean',
        description: 'Approximate (true) or exact (false) match',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_index_match',
    description: 'Perform INDEX/MATCH operation',
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
        name: 'targetCell',
        type: 'string',
        description: 'Cell to write the formula',
        required: true,
      },
      {
        name: 'returnRange',
        type: 'string',
        description: 'Range to return values from',
        required: true,
      },
      {
        name: 'lookupRange',
        type: 'string',
        description: 'Range to lookup in',
        required: true,
      },
      {
        name: 'lookupValue',
        type: 'string',
        description: 'Value to lookup',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_duplicates',
    description: 'Remove duplicate rows from range',
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
        description: 'Start cell of range',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of range',
        required: true,
      },
      {
        name: 'columns',
        type: 'array',
        description: 'Column indices to check for duplicates',
        required: false,
        items: {
          name: 'columnIndex',
          type: 'number',
          description: 'Column index',
          required: true
        },
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_text_to_columns',
    description: 'Split text into multiple columns',
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
        name: 'sourceCell',
        type: 'string',
        description: 'Source cell with text',
        required: true,
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Target start cell',
        required: true,
      },
      {
        name: 'delimiter',
        type: 'string',
        description: 'Delimiter to split by',
        required: true,
      },
      {
        name: 'numberOfColumns',
        type: 'number',
        description: 'Number of columns to create',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_flash_fill',
    description: 'Apply flash fill pattern',
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
        name: 'sourceRange',
        type: 'string',
        description: 'Source range with examples',
        required: true,
      },
      {
        name: 'targetRange',
        type: 'string',
        description: 'Target range to fill',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
];