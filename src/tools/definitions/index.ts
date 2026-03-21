/**
 * Tool Definitions Index
 * Single Responsibility: Export all tool definitions from categorized modules
 */

import { ToolDefinition } from '../../types/index.js';
import { WORKBOOK_TOOLS } from './workbook-tools.js';
import { CELL_TOOLS } from './cell-tools.js';
import { WORKSHEET_TOOLS } from './worksheet-tools.js';
import { FORMATTING_TOOLS } from './formatting-tools.js';
import { COMMENT_DATA_TOOLS } from './comment-data-tools.js';

/**
 * All Excel MCP tool definitions
 */
export const TOOL_DEFINITIONS: ToolDefinition[] = [
  ...WORKBOOK_TOOLS,
  ...CELL_TOOLS,
  ...WORKSHEET_TOOLS,
  ...FORMATTING_TOOLS,
  ...COMMENT_DATA_TOOLS,
];

// Re-export individual categories for flexibility
export { WORKBOOK_TOOLS } from './workbook-tools.js';
export { CELL_TOOLS } from './cell-tools.js';
export { WORKSHEET_TOOLS } from './worksheet-tools.js';
export { FORMATTING_TOOLS } from './formatting-tools.js';
export { COMMENT_DATA_TOOLS } from './comment-data-tools.js';