/**
 * Chart Tool Definitions
 * Single Responsibility: Chart creation and management
 */

import { ToolDefinition } from '../../types/index.js';

export const CHART_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_add_chart',
    description: 'Create a chart from data range',
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
        name: 'dataStartCell',
        type: 'string',
        description: 'Start cell of data range (e.g., A1)',
        required: true,
      },
      {
        name: 'dataEndCell',
        type: 'string',
        description: 'End cell of data range (e.g., D10)',
        required: true,
      },
      {
        name: 'chartType',
        type: 'string',
        description: 'Type of chart',
        required: true,
        enum: ['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'radar'],
      },
      {
        name: 'targetCell',
        type: 'string',
        description: 'Cell where chart will be placed (e.g., F1)',
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
        description: 'Chart width in pixels',
        required: false,
      },
      {
        name: 'height',
        type: 'number',
        description: 'Chart height in pixels',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_update_chart',
    description: 'Update chart data range',
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
        name: 'chartIndex',
        type: 'number',
        description: 'Index of the chart (0-based)',
        required: true,
      },
      {
        name: 'dataStartCell',
        type: 'string',
        description: 'New start cell of data range',
        required: true,
      },
      {
        name: 'dataEndCell',
        type: 'string',
        description: 'New end cell of data range',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_delete_chart',
    description: 'Delete a chart from worksheet',
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
        name: 'chartIndex',
        type: 'number',
        description: 'Index of the chart to delete (0-based)',
        required: true,
      },
    ],
    requiredPermissions: ['delete'],
  },
  {
    name: 'excel_list_charts',
    description: 'List all charts in a worksheet',
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
    requiredPermissions: ['read'],
  },
];