/**
 * Chart Tool Definitions
 * Single Responsibility: Chart creation and management
 */

import { ToolDefinition } from '../../types/index.js';

export const CHART_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_add_chart',
    description: 'Create chart from data',
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
        name: 'dataStartCell',
        type: 'string',
        description: 'Data start (e.g., A1)',
        required: true,
      },
      {
        name: 'dataEndCell',
        type: 'string',
        description: 'Data end (e.g., D10)',
        required: true,
      },
      {
        name: 'chartType',
        type: 'string',
        description: 'Chart type',
        required: true,
        enum: ['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'radar'],
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Placement cell (e.g., F1)',
        required: true,
      },
      {
        name: 'title',
        type: 'string',
        description: 'Chart title',
        required: false,
      },
      {
        name: 'width',
        type: 'number',
        description: 'Width in pixels',
        required: false,
      },
      {
        name: 'height',
        type: 'number',
        description: 'Height in pixels',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_update_chart',
    description: 'Update chart data',
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
        name: 'chartIndex',
        type: 'number',
        description: 'Chart index (0-based)',
        required: true,
      },
      {
        name: 'dataStartCell',
        type: 'string',
        description: 'New data start',
        required: true,
      },
      {
        name: 'dataEndCell',
        type: 'string',
        description: 'New data end',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_chart',
    description: 'Delete chart',
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
        name: 'chartIndex',
        type: 'number',
        description: 'Chart index (0-based)',
        required: true,
      },
    ],
    requiredPermissions: ['delete'],
  },
  {
    name: 'excel_list_charts',
    description: 'List worksheet charts',
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
    ],
    requiredPermissions: ['read'],
  },
];