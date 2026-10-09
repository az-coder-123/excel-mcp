/**
 * Excel MCP Server - Main Entry Point
 * Clean Architecture: Dependency injection and application bootstrap
 */

import { ExcelMCPServer } from './server/excel-mcp-server.js';
import { ServerConfig, PermissionConfig } from './types/index.js';

// Configuration from environment variables
function loadConfig(): Partial<ServerConfig> {
  const config: Partial<ServerConfig> = {
    name: process.env.MCP_SERVER_NAME || 'excel-mcp',
    version: process.env.MCP_SERVER_VERSION || '1.0.0',
    logLevel: (process.env.MCP_LOG_LEVEL as ServerConfig['logLevel']) || 'info',
  };

  // Parse permissions from environment
  const permissionsConfig = process.env.MCP_PERMISSIONS_CONFIG;
  if (permissionsConfig) {
    try {
      config.permissions = JSON.parse(permissionsConfig) as PermissionConfig;
    } catch {
      console.error('Failed to parse MCP_PERMISSIONS_CONFIG, using defaults');
    }
  }

  // Parse allowed paths
  const allowedPaths = process.env.MCP_ALLOWED_PATHS;
  if (allowedPaths) {
    if (!config.permissions) {
      config.permissions = {} as PermissionConfig;
    }
    config.permissions.allowedPaths = allowedPaths.split(',').map((p) => p.trim());
  }

  // Parse denied paths
  const deniedPaths = process.env.MCP_DENIED_PATHS;
  if (deniedPaths) {
    if (!config.permissions) {
      config.permissions = {} as PermissionConfig;
    }
    config.permissions.deniedPaths = deniedPaths.split(',').map((p) => p.trim());
  }

  // Parse max file size
  const maxFileSize = process.env.MCP_MAX_FILE_SIZE;
  if (maxFileSize) {
    if (!config.permissions) {
      config.permissions = {} as PermissionConfig;
    }
    config.permissions.maxFileSize = parseInt(maxFileSize, 10);
  }

  return config;
}

// Main function
async function main(): Promise<void> {
  try {
    const config = loadConfig();
    const server = new ExcelMCPServer(config);

    // Handle graceful shutdown
    const shutdown = async () => {
      console.error('Shutting down Excel MCP Server...');
      await server.stop();
      process.exit(0);
    };

    process.on('SIGINT', shutdown);
    process.on('SIGTERM', shutdown);

    // Start server
    await server.start();
    console.error('Excel MCP Server is running. Press Ctrl+C to stop.');
  } catch (error) {
    console.error('Failed to start Excel MCP Server:', error);
    process.exit(1);
  }
}

// Run if this is the main module
main();