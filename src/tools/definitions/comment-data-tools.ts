/**
 * Comment and Data Tool Definitions
 * Single Responsibility: Comments, hyperlinks, and data validation operations
 */

import { ToolDefinition } from '../../types/index.js';

export const COMMENT_DATA_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_add_comment',
    description: 'Add cell comment',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
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
        description: 'Author',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_comment',
    description: 'Remove cell comment',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_hyperlink',
    description: 'Add hyperlink to cell',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
      {
        name: 'url',
        type: 'string',
        description: 'URL or reference',
        required: true,
      },
      {
        name: 'display',
        type: 'string',
        description: 'Display text',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_data_validation',
    description: 'Add data validation',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
      {
        name: 'type',
        type: 'string',
        description: 'Type (list/whole/decimal/date/textLength)',
        required: true,
        enum: ['list', 'whole', 'decimal', 'date', 'textLength'],
      },
      {
        name: 'formula1',
        type: 'string',
        description: 'First value (list: comma-separated)',
        required: true,
      },
      {
        name: 'formula2',
        type: 'string',
        description: 'Second value (for between)',
        required: false,
      },
      {
        name: 'operator',
        type: 'string',
        description: 'Operator',
        required: false,
        enum: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
      },
      {
        name: 'errorMessage',
        type: 'string',
        description: 'Error message',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
];