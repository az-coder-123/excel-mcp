/**
 * Conditional Formatting Tool Definitions
 * Single Responsibility: Conditional formatting rules
 */

import { ToolDefinition } from '../../types/index.js';

export const CONDITIONAL_FORMATTING_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_add_conditional_format',
    description: 'Add conditional formatting',
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
        name: 'ruleType',
        type: 'string',
        description: 'Rule type',
        required: true,
        enum: ['cellValue', 'colorScale', 'dataBar', 'iconSet', 'containsText', 'blanks', 'errors'],
      },
      {
        name: 'operator',
        type: 'string',
        description: 'Operator (cellValue only)',
        required: false,
        enum: ['equal', 'notEqual', 'greaterThan', 'lessThan', 'between', 'notBetween', 'greaterThanOrEqual', 'lessThanOrEqual'],
      },
      {
        name: 'formula1',
        type: 'string',
        description: 'First value/formula',
        required: false,
      },
      {
        name: 'formula2',
        type: 'string',
        description: 'Second value (between/notBetween)',
        required: false,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Format when met',
        required: true,
        properties: {
          backgroundColor: { type: 'string', description: 'BG color hex' },
          fontColor: { type: 'string', description: 'Font color hex' },
          bold: { type: 'boolean', description: 'Bold' },
          italic: { type: 'boolean', description: 'Italic' },
        },
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_remove_conditional_format',
    description: 'Remove conditional formatting',
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
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_data_bar',
    description: 'Add data bar formatting',
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
    name: 'excel_add_color_scale',
    description: 'Add color scale formatting',
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
        name: 'minColor',
        type: 'string',
        description: 'Min color hex',
        required: false,
      },
      {
        name: 'midColor',
        type: 'string',
        description: 'Mid color hex',
        required: false,
      },
      {
        name: 'maxColor',
        type: 'string',
        description: 'Max color hex',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_icon_set',
    description: 'Add icon set formatting',
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
        name: 'iconSet',
        type: 'string',
        description: 'Icon set',
        required: true,
        enum: ['3Arrows', '3TrafficLights', '4Arrows', '4TrafficLights', '5Arrows', '5Rating'],
      },
    ],
    requiredPermissions: ['write'],
  },
];