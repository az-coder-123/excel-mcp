/**
 * Formula Analysis Tool Definitions
 * Single Responsibility: Formula analysis and explanation tools
 */

import { ToolDefinition } from '../../types/index.js';

export const FORMULA_ANALYSIS_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_list_formulas',
    description: 'List worksheet formulas',
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
  {
    name: 'excel_analyze_formula',
    description: 'Analyze formula structure',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_get_dependencies',
    description: 'Get cell dependencies',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_trace_precedents',
    description: 'Trace formula precedents',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_trace_dependents',
    description: 'Trace formula dependents',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_check_circular',
    description: 'Check circular references',
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
  {
    name: 'excel_explain_formula',
    description: 'Explain formula (Vietnamese)',
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
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address',
        required: true,
      },
    ],
    requiredPermissions: ['read'],
  },
  {
    name: 'excel_audit_formulas',
    description: 'Audit worksheet formulas',
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