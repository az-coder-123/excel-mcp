/**
 * Excel MCP Server - Main server implementation
 * Single Responsibility: MCP protocol handling and server lifecycle
 */

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
  CallToolResult,
} from '@modelcontextprotocol/sdk/types.js';

import { Logger } from '../utils/logger.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { ExcelService } from '../services/excel-service.js';
import { ToolHandler } from '../tools/tool-handler.js';
import { TOOL_DEFINITIONS } from '../tools/tool-definitions.js';
import { ServerConfig, PermissionConfig } from '../types/index.js';

export class ExcelMCPServer {
  private server: Server;
  private logger: Logger;
  private permissionChecker: PermissionChecker;
  private excelService: ExcelService;
  private toolHandler: ToolHandler;
  private config: ServerConfig;

  constructor(config?: Partial<ServerConfig>) {
    // Initialize configuration
    this.config = this.mergeConfig(config);

    // Initialize logger
    this.logger = new Logger(this.config.logLevel);

    // Initialize permission checker
    this.permissionChecker = new PermissionChecker(this.config.permissions);

    // Initialize Excel service
    this.excelService = new ExcelService(this.permissionChecker, this.logger);

    // Initialize tool handler
    this.toolHandler = new ToolHandler(
      this.excelService,
      this.permissionChecker,
      this.logger
    );

    // Initialize MCP server
    this.server = new Server(
      {
        name: this.config.name,
        version: this.config.version,
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupHandlers();
    this.logger.info('Excel MCP Server initialized');
  }

  /**
   * Merge user config with defaults
   */
  private mergeConfig(userConfig?: Partial<ServerConfig>): ServerConfig {
    const defaultPermissions: PermissionConfig = {
      allowedPaths: [], // Empty means all paths allowed (except denied)
      deniedPaths: [
        '/etc/*',
        '/sys/*',
        '/proc/*',
        'C:\\Windows\\*',
        'C:\\Program Files\\*',
      ],
      maxFileSize: 50 * 1024 * 1024, // 50MB
      allowedExtensions: ['.xlsx', '.xls', '.xlsm', '.xlsb'],
      permissions: ['read', 'write', 'delete'],
    };

    const defaultConfig: ServerConfig = {
      name: 'excel-mcp',
      version: '1.0.0',
      permissions: defaultPermissions,
      logLevel: 'info',
      maxConcurrentOperations: 10,
      operationTimeout: 30000,
    };

    return {
      ...defaultConfig,
      ...userConfig,
      permissions: {
        ...defaultConfig.permissions,
        ...userConfig?.permissions,
      },
    };
  }

  /**
   * Set up MCP request handlers
   */
  private setupHandlers(): void {
    // Handle list tools request
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      this.logger.debug('Listing available tools');

      const tools: Tool[] = TOOL_DEFINITIONS.map((def) => ({
        name: def.name,
        description: def.description,
        inputSchema: this.buildInputSchema(def),
      }));

      return { tools };
    });

    // Handle call tool request
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      this.logger.info(`Tool call received: ${name}`);

      const result = await this.toolHandler.executeTool(name, args || {});

      return this.formatToolResult(result);
    });

    // Error handling
    this.server.onerror = (error) => {
      this.logger.error('Server error', { error: error.message });
    };
  }

  /**
   * Build JSON Schema for tool input
   */
  private buildInputSchema(toolDef: typeof TOOL_DEFINITIONS[0]): Tool['inputSchema'] {
    const properties: Record<string, object> = {};
    const required: string[] = [];

    for (const param of toolDef.parameters) {
      properties[param.name] = {
        type: param.type,
        description: param.description,
      };

      if (param.enum) {
        (properties[param.name] as Record<string, unknown>).enum = param.enum;
      }

      if (param.required) {
        required.push(param.name);
      }
    }

    return {
      type: 'object',
      properties,
      required,
    };
  }

  /**
   * Format tool result for MCP response
   */
  private formatToolResult(result: {
    success: boolean;
    data?: unknown;
    error?: string;
  }): CallToolResult {
    if (!result.success) {
      return {
        content: [
          {
            type: 'text',
            text: `Error: ${result.error}`,
          },
        ],
        isError: true,
      };
    }

    const text = typeof result.data === 'string' 
      ? result.data 
      : JSON.stringify(result.data, null, 2);

    return {
      content: [
        {
          type: 'text',
          text,
        },
      ],
    };
  }

  /**
   * Start the server
   */
  public async start(): Promise<void> {
    try {
      const transport = new StdioServerTransport();
      await this.server.connect(transport);
      this.logger.info('Excel MCP Server started and listening on stdio');
    } catch (error) {
      this.logger.error('Failed to start server', {
        error: error instanceof Error ? error.message : 'Unknown error',
      });
      throw error;
    }
  }

  /**
   * Stop the server
   */
  public async stop(): Promise<void> {
    try {
      await this.server.close();
      this.logger.info('Excel MCP Server stopped');
    } catch (error) {
      this.logger.error('Failed to stop server', {
        error: error instanceof Error ? error.message : 'Unknown error',
      });
      throw error;
    }
  }

  /**
   * Get server configuration
   */
  public getConfig(): ServerConfig {
    return { ...this.config };
  }

  /**
   * Update permission configuration
   */
  public updatePermissions(permissions: Partial<PermissionConfig>): void {
    this.permissionChecker.updateConfig(permissions);
    this.logger.info('Permissions updated');
  }
}