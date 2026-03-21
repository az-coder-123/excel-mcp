/**
 * Formatting Tool Definitions
 * Single Responsibility: Cell and range formatting operations
 */

import { ToolDefinition } from '../../types/index.js';

export const FORMATTING_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_set_cell_format',
    description: 'Apply formatting to a cell (font, color, borders, alignment)',
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
        description: 'Cell address (e.g., A1)',
        required: true,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Formatting options object',
        required: true,
        properties: {
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'boolean' },
          fontSize: { type: 'number' },
          fontColor: { type: 'string' },
          backgroundColor: { type: 'string' },
          borderColor: { type: 'string' },
          borderStyle: { type: 'string' },
          horizontalAlignment: { type: 'string' },
          verticalAlignment: { type: 'string' },
          wrapText: { type: 'boolean' },
          numberFormat: { type: 'string' }
        }
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_range_format',
    description: 'Apply formatting to a range of cells',
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
        description: 'End cell of range (e.g., C10)',
        required: true,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Formatting options object',
        required: true,
        properties: {
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          fontSize: { type: 'number' },
          fontColor: { type: 'string' },
          backgroundColor: { type: 'string' },
          numberFormat: { type: 'string' }
        }
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_auto_fit_columns',
    description: 'Auto-fit column widths to content',
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
        description: 'Start column (e.g., A)',
        required: false,
      },
      {
        name: 'endColumn',
        type: 'string',
        description: 'End column (e.g., Z)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_column_width',
    description: 'Set width for a specific column',
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
        description: 'Column letter (e.g., A)',
        required: true,
      },
      {
        name: 'width',
        type: 'number',
        description: 'Column width',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_row_height',
    description: 'Set height for a specific row',
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
        name: 'row',
        type: 'number',
        description: 'Row number (1-based)',
        required: true,
      },
      {
        name: 'height',
        type: 'number',
        description: 'Row height',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_freeze_panes',
    description: 'Freeze panes for easier navigation',
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
        description: 'Cell to freeze at (e.g., B2 freezes row 1 and column A)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_print_area',
    description: 'Set the print area for a worksheet',
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
        description: 'Start cell of print area (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of print area (e.g., G50)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_header_footer',
    description: 'Add header and footer to worksheet',
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
        name: 'header',
        type: 'string',
        description: 'Header text (use &[Page], &[Pages], &[Date], &[Time], &[File])',
        required: false,
      },
      {
        name: 'footer',
        type: 'string',
        description: 'Footer text',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
];