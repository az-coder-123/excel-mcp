/**
 * Accounting Tool Definitions
 * Single Responsibility: Accounting-specific operations for financial data
 */

import { ToolDefinition } from '../../types/index.js';

export const ACCOUNTING_TOOLS: ToolDefinition[] = [
  // Financial Calculations
  {
    name: 'excel_financial_sum',
    description: 'Sum values with criteria',
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
        name: 'rangeStart',
        type: 'string',
        description: 'Range start',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'Range end',
        required: true,
      },
      {
        name: 'criteriaColumn',
        type: 'string',
        description: 'Criteria column',
        required: false,
      },
      {
        name: 'criteriaValue',
        type: 'string',
        description: 'Filter value',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_financial_average',
    description: 'Average values',
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
        name: 'rangeStart',
        type: 'string',
        description: 'Range start',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'Range end',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_running_total',
    description: 'Calculate running total',
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
        name: 'valueStartCell',
        type: 'string',
        description: 'Values start',
        required: true,
      },
      {
        name: 'valueEndCell',
        type: 'string',
        description: 'Values end',
        required: true,
      },
      {
        name: 'outputStartCell',
        type: 'string',
        description: 'Output start',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_percentage_of_total',
    description: 'Calculate percentage of total',
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
        name: 'valueStartCell',
        type: 'string',
        description: 'Values start',
        required: true,
      },
      {
        name: 'valueEndCell',
        type: 'string',
        description: 'Values end',
        required: true,
      },
      {
        name: 'outputStartCell',
        type: 'string',
        description: 'Output start',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_year_to_date',
    description: 'Calculate YTD totals',
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
        name: 'dateColumn',
        type: 'string',
        description: 'Date column',
        required: true,
      },
      {
        name: 'valueColumn',
        type: 'string',
        description: 'Value column',
        required: true,
      },
      {
        name: 'rangeStart',
        type: 'string',
        description: 'Data start',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'Data end',
        required: true,
      },
      {
        name: 'outputColumn',
        type: 'string',
        description: 'Output column',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },

  // Accounting Formats
  {
    name: 'excel_accounting_format',
    description: 'Apply accounting format',
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
        name: 'decimalPlaces',
        type: 'number',
        description: 'Decimals',
        required: false,
      },
      {
        name: 'useSeparator',
        type: 'boolean',
        description: 'Use separator',
        required: false,
      },
      {
        name: 'showNegativeInRed',
        type: 'boolean',
        description: 'Negative in red',
        required: false,
      },
      {
        name: 'showNegativeInParentheses',
        type: 'boolean',
        description: 'Negative in parens',
        required: false,
      },
      {
        name: 'currencySymbol',
        type: 'string',
        description: 'Currency symbol',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_vnd_currency_format',
    description: 'VND currency format',
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
        name: 'decimalPlaces',
        type: 'number',
        description: 'Decimals',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_negative_red_format',
    description: 'Negative numbers in red',
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
        name: 'decimalPlaces',
        type: 'number',
        description: 'Decimals',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_show_zeros_instead_of_empty',
    description: 'Show zeros instead of empty',
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

  // Financial Analysis
  {
    name: 'excel_period_comparison',
    description: 'Compare periods',
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
        name: 'currentValueRange',
        type: 'string',
        description: 'Current range',
        required: true,
      },
      {
        name: 'previousValueRange',
        type: 'string',
        description: 'Previous range',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Output range',
        required: true,
      },
      {
        name: 'showPercentage',
        type: 'boolean',
        description: 'Show percentage',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_variance_analysis',
    description: 'Budget vs actual variance',
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
        name: 'budgetRange',
        type: 'string',
        description: 'Budget range',
        required: true,
      },
      {
        name: 'actualRange',
        type: 'string',
        description: 'Actual range',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Output range',
        required: true,
      },
      {
        name: 'showPercentage',
        type: 'boolean',
        description: 'Show percentage',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },

  // Validation & Checks
  {
    name: 'excel_check_balance',
    description: 'Check debit/credit balance',
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
        name: 'debitRange',
        type: 'string',
        description: 'Debit range',
        required: true,
      },
      {
        name: 'creditRange',
        type: 'string',
        description: 'Credit range',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_find_anomalies',
    description: 'Find outliers',
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
        name: 'rangeStart',
        type: 'string',
        description: 'Range start',
        required: true,
      },
      {
        name: 'rangeEnd',
        type: 'string',
        description: 'Range end',
        required: true,
      },
      {
        name: 'stdDevThreshold',
        type: 'number',
        description: 'Std dev threshold',
        required: false,
      },
    ],
    requiredPermissions: ['read'],
  },
];