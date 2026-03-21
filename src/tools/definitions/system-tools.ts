/**
 * System Tools - Health check and diagnostics
 * Single Responsibility: System-level operations like health checks
 */

import { ToolDefinition } from '../../types/index.js';

/**
 * Health Check Tool
 * Verify server status, dependencies, and configuration
 */
export const SYSTEM_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_health_check',
    description: 'Check server health, dependencies, and configuration status',
    parameters: [],
    requiredPermissions: [], // No permissions needed for health check
  },
];