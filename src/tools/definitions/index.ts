/**
 * Tool Definitions Index
 * Single Responsibility: Export all tool definitions from categorized modules
 */

import { ToolDefinition } from '../../types/index.js';
import { ACCOUNTING_TOOLS } from './accounting-tools.js';
import { ADVANCED_ACCOUNTING_TOOLS } from './advanced-accounting-tools.js';
import { ADVANCED_DATA_TOOLS } from './advanced-data-tools.js';
import { ANALYSIS_TOOLS } from './analysis-tools.js';
import { CELL_TOOLS } from './cell-tools.js';
import { CHART_TOOLS } from './chart-tools.js';
import { COMMENT_DATA_TOOLS } from './comment-data-tools.js';
import { CONDITIONAL_FORMATTING_TOOLS } from './conditional-formatting-tools.js';
import { EXTENDED_FORMATTING_TOOLS } from './formatting-tools-extended.js';
import { FORMATTING_TOOLS } from './formatting-tools.js';
import { FORMULA_ANALYSIS_TOOLS } from './formula-analysis-tools.js';
import { PROTECTION_TOOLS } from './protection-tools.js';
import { SYSTEM_TOOLS } from './system-tools.js';
import { WORKBOOK_TOOLS } from './workbook-tools.js';
import { WORKSHEET_TOOLS } from './worksheet-tools.js';

/**
 * All Excel MCP tool definitions
 */
export const TOOL_DEFINITIONS: ToolDefinition[] = [
  ...WORKBOOK_TOOLS,
  ...CELL_TOOLS,
  ...WORKSHEET_TOOLS,
  ...FORMATTING_TOOLS,
  ...EXTENDED_FORMATTING_TOOLS,
  ...COMMENT_DATA_TOOLS,
  ...CHART_TOOLS,
  ...CONDITIONAL_FORMATTING_TOOLS,
  ...PROTECTION_TOOLS,
  ...ADVANCED_DATA_TOOLS,
  ...ANALYSIS_TOOLS,
  ...ACCOUNTING_TOOLS,
  ...ADVANCED_ACCOUNTING_TOOLS,
  ...FORMULA_ANALYSIS_TOOLS,
  ...SYSTEM_TOOLS,
];

// Re-export individual categories for flexibility
export { ACCOUNTING_TOOLS } from './accounting-tools.js';
export { ADVANCED_DATA_TOOLS } from './advanced-data-tools.js';
export { ANALYSIS_TOOLS } from './analysis-tools.js';
export { CELL_TOOLS } from './cell-tools.js';
export { CHART_TOOLS } from './chart-tools.js';
export { COMMENT_DATA_TOOLS } from './comment-data-tools.js';
export { CONDITIONAL_FORMATTING_TOOLS } from './conditional-formatting-tools.js';
export { FORMATTING_TOOLS } from './formatting-tools.js';
export { FORMULA_ANALYSIS_TOOLS } from './formula-analysis-tools.js';
export { PROTECTION_TOOLS } from './protection-tools.js';
export { SYSTEM_TOOLS } from './system-tools.js';
export { WORKBOOK_TOOLS } from './workbook-tools.js';
export { WORKSHEET_TOOLS } from './worksheet-tools.js';

