/**
 * Formula Analysis Tool Definitions
 * Single Responsibility: Formula analysis and explanation tools
 */

import { ToolDefinition } from '../../types/index.js';

export const FORMULA_ANALYSIS_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_list_formulas',
    description: 'List all formulas in a worksheet with their values and error status',
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
  {
    name: 'excel_analyze_formula',
    description: 'Analyze a specific formula to extract functions, cell references, and complexity',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address containing the formula (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_get_dependencies',
    description: 'Get dependency information for a cell (precedents and dependents)',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_trace_precedents',
    description: 'Trace the data sources (precedents) for a formula cell',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_trace_dependents',
    description: 'Find all cells that use the value from a specific cell (dependents)',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_check_circular',
    description: 'Check for circular references in formulas across the worksheet',
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
  {
    name: 'excel_explain_formula',
    description: 'Explain a formula in Vietnamese with step-by-step breakdown',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address containing the formula (e.g., A1, B2)',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_audit_formulas',
    description: 'Audit all formulas in worksheet for errors, warnings, and statistics',
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