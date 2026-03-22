/**
 * Advanced Accounting Tool Definitions
 * Financial calculations and analysis tools
 */

import { ToolDefinition } from '../../types/index.js';

export const ADVANCED_ACCOUNTING_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_calculate_npv',
    description: 'Calculate Net Present Value (NPV) for a series of cash flows. NPV is used to evaluate the profitability of an investment by discounting future cash flows to present value.',
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
        name: 'rate',
        type: 'number',
        description: 'Discount rate (as decimal, e.g., 0.1 for 10%)',
        required: true,
      },
      {
        name: 'valuesRange',
        type: 'string',
        description: 'Range of cash flow values (e.g., A1:A10)',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_calculate_irr',
    description: 'Calculate Internal Rate of Return (IRR) for a series of cash flows. IRR is the discount rate that makes NPV zero, representing the expected annual return of an investment.',
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
        name: 'valuesRange',
        type: 'string',
        description: 'Range of cash flow values (e.g., A1:A10)',
        required: true,
      },
      {
        name: 'guess',
        type: 'number',
        description: 'Initial guess for IRR (default: 0.1 or 10%)',
        required: false,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_calculate_financial_ratio',
    description: 'Calculate financial ratios for analysis. Supports: Current Ratio, Quick Ratio, Debt-to-Equity Ratio, Return on Equity, Profit Margin.',
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
        name: 'ratioType',
        type: 'string',
        description: 'Type of ratio to calculate',
        required: true,
        enum: ['current', 'quick', 'debt-to-equity', 'return-on-equity', 'profit-margin'],
      },
      {
        name: 'numeratorRange',
        type: 'string',
        description: 'Range for numerator values (e.g., A1:A10)',
        required: true,
      },
      {
        name: 'denominatorRange',
        type: 'string',
        description: 'Range for denominator values (e.g., B1:B10)',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_create_amortization_schedule',
    description: 'Create a loan amortization schedule showing payment breakdown into principal and interest for each period.',
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
        description: 'Starting cell for the schedule (e.g., A1)',
        required: true,
      },
      {
        name: 'principal',
        type: 'number',
        description: 'Loan principal amount',
        required: true,
      },
      {
        name: 'annualRate',
        type: 'number',
        description: 'Annual interest rate (as decimal, e.g., 0.05 for 5%)',
        required: true,
      },
      {
        name: 'numberOfPeriods',
        type: 'number',
        description: 'Number of payment periods (months)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_create_aging_report',
    description: 'Create an aging report for accounts receivable or payable, categorizing amounts by days past due: 0-30, 31-60, 61-90, 91-120, 120+ days.',
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
        name: 'invoiceDateColumn',
        type: 'string',
        description: 'Column letter containing invoice dates (e.g., A)',
        required: true,
      },
      {
        name: 'amountColumn',
        type: 'string',
        description: 'Column letter containing amounts (e.g., B)',
        required: true,
      },
      {
        name: 'asOfDate',
        type: 'string',
        description: 'As-of date for aging calculation (ISO format: YYYY-MM-DD)',
        required: true,
      },
      {
        name: 'outputStartCell',
        type: 'string',
        description: 'Starting cell for the aging report (e.g., D1)',
        required: true,
      },
    ],
    requiredPermissions: ['read', 'write'],
  },
  {
    name: 'excel_calculate_tax',
    description: 'Calculate tax for amounts at a specified tax rate.',
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
        name: 'amountRange',
        type: 'string',
        description: 'Range of amounts to calculate tax on (e.g., A1:A10)',
        required: true,
      },
      {
        name: 'taxRate',
        type: 'number',
        description: 'Tax rate percentage (e.g., 10 for 10%)',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Range to write tax amounts (e.g., B1:B10)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_convert_currency',
    description: 'Convert amounts from one currency to another using an exchange rate.',
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
        name: 'amountRange',
        type: 'string',
        description: 'Range of amounts to convert (e.g., A1:A10)',
        required: true,
      },
      {
        name: 'exchangeRate',
        type: 'number',
        description: 'Exchange rate (e.g., 24000 for VND per USD)',
        required: true,
      },
      {
        name: 'outputRange',
        type: 'string',
        description: 'Range to write converted amounts (e.g., B1:B10)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
];