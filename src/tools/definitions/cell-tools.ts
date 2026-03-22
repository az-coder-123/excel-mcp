/**
 * Cell Tool Definitions
 * Single Responsibility: Cell and range data operations
 */

import { ToolDefinition } from '../../types/index.js';

export const CELL_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_read_cell',
    description: 'Read single cell value',
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
        description: 'Cell address (e.g., A1)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_read_range',
    description: 'Read cell range values',
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
        description: 'Range start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end cell (e.g., C10)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_write_cell',
    description: 'Write value to cell',
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
        description: 'Cell address (e.g., A1)',
        required: true,
      },
      {
        name: 'value',
        type: 'string',
        description: 'Value (string, number, or formula with =)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_write_batch',
    description: 'Write multiple cells at once',
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
        name: 'data',
        type: 'array',
        description: 'Array of {cellAddress, value}',
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
    description: 'Get cell details (type, formula)',
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
        description: 'Cell address (e.g., A1)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_copy_range',
    description: 'Copy cell range to location',
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
        description: 'Source worksheet name',
        required: true,
      },
      {
        name: 'sourceStart',
        type: 'string',
        description: 'Source range start (e.g., A1)',
        required: true,
      },
      {
        name: 'sourceEnd',
        type: 'string',
        description: 'Source range end (e.g., C10)',
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
        description: 'Target worksheet (optional)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_find_replace',
    description: 'Find and replace text',
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
        name: 'findText',
        type: 'string',
        description: 'Text to find',
        required: true,
      },
      {
        name: 'replaceText',
        type: 'string',
        description: 'Replacement text',
        required: true,
      },
      {
        name: 'matchCase',
        type: 'boolean',
        description: 'Case sensitive',
        required: false,
      },
      {
        name: 'matchEntireCell',
        type: 'boolean',
        description: 'Match entire cell',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_sort_range',
    description: 'Sort range by column',
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
      {
        name: 'sortColumn',
        type: 'number',
        description: 'Column number (1-based)',
        required: true,
      },
      {
        name: 'ascending',
        type: 'boolean',
        description: 'Ascending (true) or descending',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_get_named_ranges',
    description: 'List named ranges',
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
    name: 'excel_add_named_range',
    description: 'Create named range',
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
        name: 'name',
        type: 'string',
        description: 'Range name',
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
        description: 'Range end (e.g., D10)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_calculate_formula',
    description: 'Recalculate all formulas',
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
];