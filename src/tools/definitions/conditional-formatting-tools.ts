/**
 * Conditional Formatting Tool Definitions
 * Single Responsibility: Conditional formatting rules
 */

import { ToolDefinition } from '../../types/index.js';

export const CONDITIONAL_FORMATTING_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_add_conditional_format',
    description: 'Add conditional formatting to a range',
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
        description: 'End cell of range (e.g., A10)',
        required: true,
      },
      {
        name: 'ruleType',
        type: 'string',
        description: 'Type of conditional formatting rule',
        required: true,
        enum: ['cellValue', 'colorScale', 'dataBar', 'iconSet', 'containsText', 'blanks', 'errors'],
      },
      {
        name: 'operator',
        type: 'string',
        description: 'Comparison operator for cellValue rules',
        required: false,
        enum: ['equal', 'notEqual', 'greaterThan', 'lessThan', 'between', 'notBetween', 'greaterThanOrEqual', 'lessThanOrEqual'],
      },
      {
        name: 'formula1',
        type: 'string',
        description: 'First formula or value for the rule',
        required: false,
      },
      {
        name: 'formula2',
        type: 'string',
        description: 'Second formula or value (for between/notBetween)',
        required: false,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Format to apply when condition is met',
        required: true,
        properties: {
          backgroundColor: { type: 'string', description: 'Background color (hex)' },
          fontColor: { type: 'string', description: 'Font color (hex)' },
          bold: { type: 'boolean', description: 'Bold text' },
          italic: { type: 'boolean', description: 'Italic text' },
        },
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_conditional_format',
    description: 'Remove conditional formatting from a range',
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
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_data_bar',
    description: 'Add data bar conditional formatting',
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
        name: 'color',
        type: 'string',
        description: 'Data bar color (hex)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_color_scale',
    description: 'Add color scale conditional formatting',
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
        name: 'minColor',
        type: 'string',
        description: 'Color for minimum value (hex)',
        required: false,
      },
      {
        name: 'midColor',
        type: 'string',
        description: 'Color for middle value (hex)',
        required: false,
      },
      {
        name: 'maxColor',
        type: 'string',
        description: 'Color for maximum value (hex)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_icon_set',
    description: 'Add icon set conditional formatting',
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
        name: 'iconSet',
        type: 'string',
        description: 'Icon set type',
        required: true,
        enum: ['3Arrows', '3TrafficLights', '4Arrows', '4TrafficLights', '5Arrows', '5Rating'],
      },
    ],
    requiredPermissions: ['write'],
  },
];