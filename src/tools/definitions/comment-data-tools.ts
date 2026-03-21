/**
 * Comment and Data Tool Definitions
 * Single Responsibility: Comments, hyperlinks, and data validation operations
 */

import { ToolDefinition } from '../../types/index.js';

export const COMMENT_DATA_TOOLS: ToolDefinition[] = [
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
];