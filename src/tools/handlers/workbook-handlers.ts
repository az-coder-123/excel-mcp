/**
 * Workbook Handlers
 * Single Responsibility: Workbook lifecycle operations
 */

import { ExcelService } from '../../services/excel-service.js';
import { OperationResult } from '../../types/index.js';
import { BaseHandler } from './base-handler.js';

export class WorkbookHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleOpenWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filePath = this.getStringArg(args, 'filePath');
    if (!filePath) {
      return { success: false, error: 'Missing required parameter: filePath' };
    }
    return this.excelService.openWorkbook(filePath);
  }

  async handleCreateWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
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

  async handleSaveWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const outputPath = this.getStringArg(args, 'outputPath');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.saveWorkbook(filename, outputPath || undefined);
  }

  async handleCloseWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.closeWorkbook(filename);
  }

  async handleExportWorksheetToNewFile(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const newFilePath = this.getStringArg(args, 'newFilePath');
    
    if (!filename || !worksheet || !newFilePath) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.exportWorksheetToNewFile(filename, worksheet, newFilePath);
  }

  async handleGetWorkbookContext(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.getWorkbookContext(filename);
  }
}
