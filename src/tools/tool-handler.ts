/**
 * Tool Handler - Processes MCP tool calls
 * Single Responsibility: Handles tool execution and validation
 */

import { ExcelService } from '../services/excel-service.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult, CellRange } from '../types/index.js';
import { getToolDefinition } from './tool-definitions.js';

export class ToolHandler {
  private excelService: ExcelService;
  private permissionChecker: PermissionChecker;
  private logger: Logger;

  constructor(
    excelService: ExcelService,
    permissionChecker: PermissionChecker,
    logger: Logger
  ) {
    this.excelService = excelService;
    this.permissionChecker = permissionChecker;
    this.logger = logger;
  }

  /**
   * Execute a tool call
   */
  public async executeTool(
    toolName: string,
    args: Record<string, unknown>
  ): Promise<OperationResult> {
    this.logger.info(`Executing tool: ${toolName}`, { args });

    // Validate tool exists
    const toolDef = getToolDefinition(toolName);
    if (!toolDef) {
      return {
        success: false,
        error: `Unknown tool: ${toolName}`,
      };
    }

    // Validate permissions
    for (const permission of toolDef.requiredPermissions) {
      const permResult = this.permissionChecker.hasPermission(permission);
      if (!permResult.success) {
        return permResult;
      }
    }

    // Execute tool
    try {
      switch (toolName) {
        case 'excel_open_workbook':
          return await this.handleOpenWorkbook(args);
        case 'excel_create_workbook':
          return await this.handleCreateWorkbook(args);
        case 'excel_read_cell':
          return await this.handleReadCell(args);
        case 'excel_read_range':
          return await this.handleReadRange(args);
        case 'excel_write_cell':
          return await this.handleWriteCell(args);
        case 'excel_list_worksheets':
          return await this.handleListWorksheets(args);
        case 'excel_add_worksheet':
          return await this.handleAddWorksheet(args);
        case 'excel_save_workbook':
          return await this.handleSaveWorkbook(args);
        case 'excel_close_workbook':
          return await this.handleCloseWorkbook(args);
        case 'excel_get_cell_info':
          return await this.handleGetCellInfo(args);
        default:
          return {
            success: false,
            error: `Tool not implemented: ${toolName}`,
          };
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Tool execution failed: ${message}`);
      return { success: false, error: message };
    }
  }

  private async handleOpenWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filePath = this.getStringArg(args, 'filePath');
    if (!filePath) {
      return { success: false, error: 'Missing required parameter: filePath' };
    }
    return this.excelService.openWorkbook(filePath);
  }

  private async handleCreateWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const outputPath = this.getStringArg(args, 'outputPath');
    
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    if (!outputPath) {
      return { success: false, error: 'Missing required parameter: outputPath' };
    }

    const fullPath = `${outputPath.replace(/\/$/, '')}/${filename}.xlsx`;
    return this.excelService.createWorkbook(fullPath);
  }

  private async handleReadCell(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.readCell(filename, worksheet, cellAddress);
  }

  private async handleReadRange(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    const range = this.parseCellRange(startCell, endCell);
    if (!range) {
      return { success: false, error: 'Invalid cell range format' };
    }

    return this.excelService.readRange(filename, worksheet, range);
  }

  private async handleWriteCell(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const value = args['value'];

    if (!filename || !worksheet || !cellAddress || value === undefined) {
      return { success: false, error: 'Missing required parameters' };
    }

    // Try to parse as number if it looks like one
    let parsedValue: unknown = value;
    if (typeof value === 'string') {
      // Check if it's a formula
      if (value.startsWith('=')) {
        parsedValue = { formula: value };
      } else if (!isNaN(Number(value)) && value.trim() !== '') {
        parsedValue = Number(value);
      }
    }

    return this.excelService.writeCell(filename, worksheet, cellAddress, parsedValue);
  }

  private async handleListWorksheets(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.getWorksheets(filename);
  }

  private async handleAddWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheetName = this.getStringArg(args, 'worksheetName');

    if (!filename || !worksheetName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addWorksheet(filename, worksheetName);
  }

  private async handleSaveWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const outputPath = this.getStringArg(args, 'outputPath');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.saveWorkbook(filename, outputPath || undefined);
  }

  private async handleCloseWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.closeWorkbook(filename);
  }

  private async handleGetCellInfo(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.readCell(filename, worksheet, cellAddress);
  }

  // Helper methods
  private getStringArg(args: Record<string, unknown>, key: string): string | undefined {
    const value = args[key];
    return typeof value === 'string' ? value : undefined;
  }

  private parseCellRange(startCell: string, endCell: string): CellRange | null {
    const parseAddress = (addr: string): { row: number; column: number } | null => {
      const match = addr.match(/^([A-Z]+)(\d+)$/i);
      if (!match) return null;

      const colStr = match[1].toUpperCase();
      const row = parseInt(match[2], 10);

      let column = 0;
      for (let i = 0; i < colStr.length; i++) {
        column = column * 26 + (colStr.charCodeAt(i) - 64);
      }

      return { row, column };
    };

    const start = parseAddress(startCell);
    const end = parseAddress(endCell);

    if (!start || !end) return null;

    return { start, end };
  }
}