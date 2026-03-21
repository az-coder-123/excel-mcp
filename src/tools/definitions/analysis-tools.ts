/**
 * Analysis Tool Definitions
 * Single Responsibility: Data analysis, statistics, filtering, and profiling operations
 */

import { ToolDefinition } from '../../types/index.js';

export const ANALYSIS_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_get_column_stats',
    description: 'Get statistics for a column: unique values, counts, percentages, min/max, missing values',
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
        description: 'Column letter (e.g., A, B, C)',
        required: true,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Whether the first row contains header (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_filter_data',
    description: 'Filter data based on multiple conditions',
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
        description: 'Start cell of data range (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of data range (e.g., D100)',
        required: true,
      },
      {
        name: 'filters',
        type: 'array',
        description: 'Array of filter conditions',
        required: true,
        items: {
          name: 'filter',
          type: 'object',
          description: 'Filter condition object',
          required: true,
          properties: {
            column: {
              type: 'string',
              description: 'Column letter (e.g., A, B, C)',
            },
            operator: {
              type: 'string',
              description: 'Comparison operator (equals, notEquals, contains, notContains, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, isBlank, isNotBlank)',
            },
            value: {
              type: 'string',
              description: 'Value to compare with (not required for isBlank/isNotBlank)',
            },
            matchCase: {
              type: 'boolean',
              description: 'Case sensitive comparison (default: false)',
            },
          },
        },
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Whether the first row contains header (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_group_aggregate',
    description: 'Group data and perform aggregate operations (count, sum, avg, max, min)',
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
        description: 'Start cell of data range (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of data range (e.g., D100)',
        required: true,
      },
      {
        name: 'groupByColumn',
        type: 'string',
        description: 'Column letter to group by (e.g., A, B, C)',
        required: true,
      },
      {
        name: 'aggregateColumn',
        type: 'string',
        description: 'Column letter to aggregate (e.g., A, B, C). Leave empty for count operation',
        required: false,
      },
      {
        name: 'operation',
        type: 'string',
        description: 'Aggregate operation: count, sum, avg, max, min',
        required: true,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Whether the first row contains header (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_profile_data',
    description: 'Analyze data profile: data types, null percentage, unique values, patterns for a range',
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
        description: 'End cell of range (e.g., D10)',
        required: true,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Whether the first row contains header (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_search',
    description: 'Advanced search in worksheet with multiple search types',
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
        name: 'searchQuery',
        type: 'string',
        description: 'Search query text or pattern',
        required: true,
      },
      {
        name: 'searchType',
        type: 'string',
        description: 'Search type: exact, contains, startsWith, endsWith, regex',
        required: false,
      },
      {
        name: 'columnRange',
        type: 'string',
        description: 'Column range to search in (e.g., A:C). Search all columns if not specified',
        required: false,
      },
      {
        name: 'matchCase',
        type: 'boolean',
        description: 'Case sensitive search (default: false)',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum number of results to return (default: 100)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_compare_ranges',
    description: 'Compare two ranges and return differences',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet1',
        type: 'string',
        description: 'First worksheet name',
        required: true,
      },
      {
        name: 'range1Start',
        type: 'string',
        description: 'Start cell of first range (e.g., A1)',
        required: true,
      },
      {
        name: 'range1End',
        type: 'string',
        description: 'End cell of first range (e.g., D10)',
        required: true,
      },
      {
        name: 'worksheet2',
        type: 'string',
        description: 'Second worksheet name (leave empty to use same worksheet)',
        required: false,
      },
      {
        name: 'range2Start',
        type: 'string',
        description: 'Start cell of second range (e.g., A1)',
        required: true,
      },
      {
        name: 'range2End',
        type: 'string',
        description: 'End cell of second range (e.g., D10)',
        required: true,
      },
      {
        name: 'compareType',
        type: 'string',
        description: 'Comparison type: values (default), formulas, formats',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
];