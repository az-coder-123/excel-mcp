/**
 * Analysis Tool Definitions
 * Single Responsibility: Data analysis, statistics, filtering, and profiling operations
 */

import { ToolDefinition } from '../../types/index.js';

export const ANALYSIS_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_get_column_stats',
    description: 'Get column statistics',
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
        name: 'column',
        type: 'string',
        description: 'Column letter',
        required: true,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Has header',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_filter_data',
    description: 'Filter data by conditions',
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
        name: 'filters',
        type: 'array',
        description: 'Filter conditions',
        required: true,
        items: {
          name: 'filter',
          type: 'object',
          description: 'Filter',
          required: true,
          properties: {
            column: {
              type: 'string',
              description: 'Column',
            },
            operator: {
              type: 'string',
              description: 'Operator',
            },
            value: {
              type: 'string',
              description: 'Value',
            },
            matchCase: {
              type: 'boolean',
              description: 'Match case',
            },
          },
        },
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Has header',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_group_aggregate',
    description: 'Group and aggregate data',
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
        name: 'groupByColumn',
        type: 'string',
        description: 'Group column',
        required: true,
      },
      {
        name: 'aggregateColumn',
        type: 'string',
        description: 'Aggregate column',
        required: false,
      },
      {
        name: 'operation',
        type: 'string',
        description: 'Operation (count/sum/avg/max/min)',
        required: true,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Has header',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_profile_data',
    description: 'Profile data range',
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
        name: 'hasHeader',
        type: 'boolean',
        description: 'Has header',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_search',
    description: 'Advanced search',
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
        name: 'searchQuery',
        type: 'string',
        description: 'Search query',
        required: true,
      },
      {
        name: 'searchType',
        type: 'string',
        description: 'Type (exact/contains/startsWith/endsWith/regex)',
        required: false,
      },
      {
        name: 'columnRange',
        type: 'string',
        description: 'Column range',
        required: false,
      },
      {
        name: 'matchCase',
        type: 'boolean',
        description: 'Match case',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Max results',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_compare_ranges',
    description: 'Compare two ranges',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet1',
        type: 'string',
        description: 'First worksheet',
        required: true,
      },
      {
        name: 'range1Start',
        type: 'string',
        description: 'Range 1 start',
        required: true,
      },
      {
        name: 'range1End',
        type: 'string',
        description: 'Range 1 end',
        required: true,
      },
      {
        name: 'worksheet2',
        type: 'string',
        description: 'Second worksheet',
        required: false,
      },
      {
        name: 'range2Start',
        type: 'string',
        description: 'Range 2 start',
        required: true,
      },
      {
        name: 'range2End',
        type: 'string',
        description: 'Range 2 end',
        required: true,
      },
      {
        name: 'compareType',
        type: 'string',
        description: 'Type (values/formulas/formats)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
];