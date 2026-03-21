/**
 * Cell Tool Definitions
 * Single Responsibility: Cell and range data operations
 */

import { ToolDefinition } from '../../types/index.js';

export const CELL_TOOLS: ToolDefinition[] = [
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
    name: 'excel_write_batch',
    description: 'Write multiple values to cells in batch for better performance',
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
        name: 'data',
        type: 'array',
        description: 'Array of objects with cellAddress and value properties',
        required: true,
        items: {
          name: 'item',
          type: 'object',
          description: 'Cell data item',
          required: true,
          properties: {
            cellAddress: { type: 'string', description: 'Cell address' },
            value: { type: 'string', description: 'Cell value' }
          }
        }
      },
    ],
    requiredPermissions: ['write'],
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
  {
    name: 'excel_copy_range',
    description: 'Copy a range of cells to another location',
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
        name: 'sourceStart',
        type: 'string',
        description: 'Source range start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'sourceEnd',
        type: 'string',
        description: 'Source range end cell (e.g., C10)',
        required: true,
      },
      {
        name: 'targetStart',
        type: 'string',
        description: 'Target start cell (e.g., E1)',
        required: true,
      },
      {
        name: 'targetWorksheet',
        type: 'string',
        description: 'Target worksheet (same worksheet if not specified)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_find_replace',
    description: 'Find and replace text in worksheet',
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
        name: 'findText',
        type: 'string',
        description: 'Text to find',
        required: true,
      },
      {
        name: 'replaceText',
        type: 'string',
        description: 'Text to replace with',
        required: true,
      },
      {
        name: 'matchCase',
        type: 'boolean',
        description: 'Case sensitive search',
        required: false,
      },
      {
        name: 'matchEntireCell',
        type: 'boolean',
        description: 'Match entire cell content',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_sort_range',
    description: 'Sort a range of data',
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
      {
        name: 'sortColumn',
        type: 'number',
        description: 'Column number to sort by (1-based)',
        required: true,
      },
      {
        name: 'ascending',
        type: 'boolean',
        description: 'Sort ascending (true) or descending (false)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_get_named_ranges',
    description: 'Get list of named ranges in workbook',
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
    name: 'excel_add_named_range',
    description: 'Create a named range',
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
        name: 'name',
        type: 'string',
        description: 'Name for the range',
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
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_calculate_formula',
    description: 'Force recalculation of all formulas',
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