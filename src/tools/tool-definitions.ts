/**
 * Tool Definitions - Main entry point for MCP tool definitions
 * Single Responsibility: Re-exports from categorized modules
 */

import { ToolDefinition, Permission } from '../types/index.js';

// Import all tools from categorized modules
import {
  TOOL_DEFINITIONS,
  WORKBOOK_TOOLS,
  CELL_TOOLS,
  WORKSHEET_TOOLS,
  FORMATTING_TOOLS,
  COMMENT_DATA_TOOLS
} from './definitions/index.js';

// Re-export for external use
export {
  TOOL_DEFINITIONS,
  WORKBOOK_TOOLS,
  CELL_TOOLS,
  WORKSHEET_TOOLS,
  FORMATTING_TOOLS,
  COMMENT_DATA_TOOLS
};

/**
 * Get tool definition by name
 */
export function getToolDefinition(name: string): ToolDefinition | undefined {
  return TOOL_DEFINITIONS.find((tool) => tool.name === name);
}

/**
 * Get all tool names
 */
export function getAllToolNames(): string[] {
  return TOOL_DEFINITIONS.map((tool) => tool.name);
}

/**
 * Get tools by required permission
 */
export function getToolsByPermission(permission: Permission): ToolDefinition[] {
  return TOOL_DEFINITIONS.filter((tool) => 
    tool.requiredPermissions.includes(permission)
  );
}