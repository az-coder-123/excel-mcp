/**
 * Extended Formatting Tool Definitions
 * All-in-one formatting tools for efficiency
 */

import { ToolDefinition } from '../../types/index.js';

export const EXTENDED_FORMATTING_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_format_worksheet',
    description: 'Format entire worksheet with professional styling in one operation',
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
        name: 'titleStyle',
        type: 'object',
        description: 'Title cell styling (optional)',
        required: false,
        properties: {
          cellAddress: {
            type: 'string',
            description: 'Title cell address (e.g., A1)',
          },
          fontSize: {
            type: 'number',
            description: 'Font size (default: 16)',
          },
          color: {
            type: 'string',
            description: 'Font color hex (default: FF1F4E78)',
          },
          bold: {
            type: 'boolean',
            description: 'Bold text (default: true)',
          },
        },
      },
      {
        name: 'headerStyle',
        type: 'object',
        description: 'Header row styling (optional)',
        required: false,
        properties: {
          startCell: {
            type: 'string',
            description: 'Header range start (e.g., A2)',
          },
          endCell: {
            type: 'string',
            description: 'Header range end (e.g., F2)',
          },
          backgroundColor: {
            type: 'string',
            description: 'Background hex (default: FFD3D3D3)',
          },
          fontColor: {
            type: 'string',
            description: 'Font color hex (default: FF000000)',
          },
          bold: {
            type: 'boolean',
            description: 'Bold text (default: true)',
          },
        },
      },
      {
        name: 'dataStyle',
        type: 'object',
        description: 'Data range styling (optional)',
        required: false,
        properties: {
          startCell: {
            type: 'string',
            description: 'Data range start (e.g., A3)',
          },
          endCell: {
            type: 'string',
            description: 'Data range end (e.g., F100)',
          },
          borders: {
            type: 'boolean',
            description: 'Apply borders (default: true)',
          },
          borderStyle: {
            type: 'string',
            description: 'Border style: thin/medium/thick (default: thin)',
          },
          borderColor: {
            type: 'string',
            description: 'Border color hex (default: FF000000)',
          },
          alternateRows: {
            type: 'boolean',
            description: 'Alternate row colors (default: false)',
          },
          alternateRowColor: {
            type: 'string',
            description: 'Alternate row color hex (default: FFF9F9F9)',
          },
        },
      },
      {
        name: 'columnWidths',
        type: 'string',
        description: 'JSON string of column widths: [{"column":"A","width":20},{"column":"B","width":15}] (optional)',
        required: false,
      },
      {
        name: 'autoFitColumns',
        type: 'boolean',
        description: 'Auto-fit all columns (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_batch_format',
    description: 'Apply multiple formatting operations in one call',
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
        name: 'operations',
        type: 'string',
        description: 'JSON string array of operations (e.g., [{"type":"borders","startCell":"A3","endCell":"F100"}])',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
];