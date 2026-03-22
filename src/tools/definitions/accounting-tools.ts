/**
 * Accounting Tool Definitions
 * Single Responsibility: Accounting-specific operations for financial data
 */

import { ToolDefinition } from '../../types/index.js';

export const ACCOUNTING_TOOLS: ToolDefinition[] = [
  // Financial Calculations
  {
    name: 'excel_financial_sum',
    description: 'Calculate sum of values in a range, optionally with criteria',
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
        name: 'rangeStart',
        type: 'string',
        description: 'Start cell of range (e.g., A1)',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'End cell of range (e.g., A10)',
        required: true,
      },
      {
        name: 'criteriaColumn',
        type: 'string',
        description: 'Column letter for criteria filtering (e.g., A, B)',
        required: false,
      },
      {
        name: 'criteriaValue',
        type: 'string',
        description: 'Value to filter by in criteria column',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_financial_average',
    description: 'Calculate average of values in a range',
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
        name: 'rangeStart',
        type: 'string',
        description: 'Start cell of range (e.g., A1)',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'End cell of range (e.g., A10)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_running_total',
    description: 'Calculate running total (cumulative sum) for a column',
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
        name: 'valueStartCell',
        type: 'string',
        description: 'Start cell of values (e.g., A2)',
        required: true,
      },
      {
        name: 'valueEndCell',
        type: 'string',
        description: 'End cell of values (e.g., A100)',
        required: true,
      },
      {
        name: 'outputStartCell',
        type: 'string',
        description: 'Start cell for running total (e.g., B2)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_percentage_of_total',
    description: 'Calculate percentage of total for each value in a range',
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
        name: 'valueStartCell',
        type: 'string',
        description: 'Start cell of values (e.g., A2)',
        required: true,
      },
      {
        name: 'valueEndCell',
        type: 'string',
        description: 'End cell of values (e.g., A100)',
        required: true,
      },
      {
        name: 'outputStartCell',
        type: 'string',
        description: 'Start cell for percentages (e.g., B2)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_year_to_date',
    description: 'Calculate year-to-date (YTD) cumulative values',
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
        name: 'dateColumn',
        type: 'string',
        description: 'Column letter containing dates (e.g., A)',
        required: true,
      },
      {
        name: 'valueColumn',
        type: 'string',
        description: 'Column letter containing values (e.g., B)',
        required: true,
      },
      {
        name: 'rangeStart',
        type: 'string',
        description: 'Start cell of data (e.g., A1)',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'End cell of data (e.g., D365)',
        required: true,
      },
      {
        name: 'outputColumn',
        type: 'string',
        description: 'Column letter for YTD output (e.g., C)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },

  // Accounting Formats
  {
    name: 'excel_accounting_format',
    description: 'Apply professional accounting number format',
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
        name: 'decimalPlaces',
        type: 'number',
        description: 'Number of decimal places (default: 2)',
        required: false,
      },
      {
        name: 'useSeparator',
        type: 'boolean',
        description: 'Use thousand separator (default: true)',
        required: false,
      },
      {
        name: 'showNegativeInRed',
        type: 'boolean',
        description: 'Show negative numbers in red (default: true)',
        required: false,
      },
      {
        name: 'showNegativeInParentheses',
        type: 'boolean',
        description: 'Show negative numbers in parentheses (default: true)',
        required: false,
      },
      {
        name: 'currencySymbol',
        type: 'string',
        description: 'Currency symbol (default: empty)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_vnd_currency_format',
    description: 'Apply Vietnamese Dong (VND) currency format',
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
        name: 'decimalPlaces',
        type: 'number',
        description: 'Number of decimal places (default: 0)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_negative_red_format',
    description: 'Apply negative numbers in red format',
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
        name: 'decimalPlaces',
        type: 'number',
        description: 'Number of decimal places (default: 2)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_show_zeros_instead_of_empty',
    description: 'Replace empty cells with zeros',
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

  // Financial Analysis
  {
    name: 'excel_period_comparison',
    description: 'Calculate period comparison (current vs previous period)',
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
        name: 'currentValueRange',
        type: 'string',
        description: 'Range of current period values (e.g., B2:B13)',
        required: true,
      },
      {
        name: 'previousValueRange',
        type: 'string',
        description: 'Range of previous period values (e.g., C2:C13)',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Range for comparison results (e.g., D2:D13)',
        required: true,
      },
      {
        name: 'showPercentage',
        type: 'boolean',
        description: 'Show percentage change instead of absolute (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_variance_analysis',
    description: 'Calculate variance between budget and actual values',
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
        name: 'budgetRange',
        type: 'string',
        description: 'Range of budget values (e.g., B2:B13)',
        required: true,
      },
      {
        name: 'actualRange',
        type: 'string',
        description: 'Range of actual values (e.g., C2:C13)',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Range for variance results (e.g., D2:D13)',
        required: true,
      },
      {
        name: 'showPercentage',
        type: 'boolean',
        description: 'Show percentage variance instead of absolute (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },

  // Validation & Checks
  {
    name: 'excel_check_balance',
    description: 'Check if debit total equals credit total',
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
        name: 'debitRange',
        type: 'string',
        description: 'Range of debit column (e.g., C2:C100)',
        required: true,
      },
      {
        name: 'creditRange',
        type: 'string',
        description: 'Range of credit column (e.g., D2:D100)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_find_anomalies',
    description: 'Find anomalous values (outliers) using standard deviation',
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
        name: 'rangeStart',
        type: 'string',
        description: 'Start cell of range (e.g., A1)',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'End cell of range (e.g., A100)',
        required: true,
      },
      {
        name: 'stdDevThreshold',
        type: 'number',
        description: 'Standard deviation threshold for anomalies (default: 2)',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
];