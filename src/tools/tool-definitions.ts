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
    name: 'excel_add_data_validation',
    description: 'Add data validation to a cell (dropdown, number range, etc.)',
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
        name: 'type',
        type: 'string',
        description: 'Validation type (list, whole, decimal, date, textLength)',
        required: true,
        enum: ['list', 'whole', 'decimal', 'date', 'textLength'],
      },
      {
        name: 'formula1',
        type: 'string',
        description: 'First formula/value (for list: comma-separated items)',
        required: true,
      },
      {
        name: 'formula2',
        type: 'string',
        description: 'Second formula (for between operators)',
        required: false,
      },
      {
        name: 'operator',
        type: 'string',
        description: 'Comparison operator',
        required: false,
        enum: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
      },
      {
        name: 'errorMessage',
        type: 'string',
        description: 'Error message for invalid input',
        required: false,
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
  {
    name: 'excel_add_hyperlink',
    description: 'Add hyperlink to a cell',
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
        name: 'url',
        type: 'string',
        description: 'URL or internal reference',
        required: true,
      },
      {
        name: 'display',
        type: 'string',
        description: 'Display text for the hyperlink',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_comment',
    description: 'Add comment/note to a cell',
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
        name: 'comment',
        type: 'string',
        description: 'Comment text',
        required: true,
      },
      {
        name: 'author',
        type: 'string',
        description: 'Comment author',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_comment',
    description: 'Remove comment from a cell',
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