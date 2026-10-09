/**
 * Advanced Accounting Tool Definitions
 * Financial calculations and analysis tools
 */

import { ToolDefinition } from '../../types/index.js';

export const ADVANCED_ACCOUNTING_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_calculate_npv',
    description: 'Calculate NPV for cash flows',
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
        name: 'rate',
        type: 'number',
        description: 'Discount rate (decimal)',
        required: true,
      },
      {
        name: 'valuesRange',
        type: 'string',
        description: 'Cash flow range',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_calculate_irr',
    description: 'Calculate IRR for cash flows',
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
        name: 'valuesRange',
        type: 'string',
        description: 'Cash flow range',
        required: true,
      },
      {
        name: 'guess',
        type: 'number',
        description: 'Initial guess',
        required: false,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_calculate_financial_ratio',
    description: 'Calculate financial ratios',
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
        name: 'ratioType',
        type: 'string',
        description: 'Ratio type',
        required: true,
        enum: ['current', 'quick', 'debt-to-equity', 'return-on-equity', 'profit-margin'],
      },
      {
        name: 'numeratorRange',
        type: 'string',
        description: 'Numerator range',
        required: true,
      },
      {
        name: 'denominatorRange',
        type: 'string',
        description: 'Denominator range',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_create_amortization_schedule',
    description: 'Create loan amortization schedule',
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
        description: 'Start cell',
        required: true,
      },
      {
        name: 'principal',
        type: 'number',
        description: 'Principal amount',
        required: true,
      },
      {
        name: 'annualRate',
        type: 'number',
        description: 'Annual rate (decimal)',
        required: true,
      },
      {
        name: 'numberOfPeriods',
        type: 'number',
        description: 'Periods (months)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_create_aging_report',
    description: 'Create aging report',
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
        name: 'invoiceDateColumn',
        type: 'string',
        description: 'Date column',
        required: true,
      },
      {
        name: 'amountColumn',
        type: 'string',
        description: 'Amount column',
        required: true,
      },
      {
        name: 'asOfDate',
        type: 'string',
        description: 'As-of date (ISO)',
        required: true,
      },
      {
        name: 'outputStartCell',
        type: 'string',
        description: 'Output start',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_calculate_tax',
    description: 'Calculate tax amounts',
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
        name: 'amountRange',
        type: 'string',
        description: 'Amount range',
        required: true,
      },
      {
        name: 'taxRate',
        type: 'number',
        description: 'Tax rate (%)',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Output range',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_convert_currency',
    description: 'Convert currency amounts',
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
        name: 'amountRange',
        type: 'string',
        description: 'Amount range',
        required: true,
      },
      {
        name: 'exchangeRate',
        type: 'number',
        description: 'Exchange rate',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Output range',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
];