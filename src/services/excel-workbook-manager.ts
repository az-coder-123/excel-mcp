/**
 * Excel Workbook Manager - Handles workbook lifecycle operations
 * Single Responsibility: Workbook open, create, save, close operations
 */

import ExcelJS from 'exceljs';
import { WorkbookInfo, WorksheetInfo, OperationResult } from '../types/index.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';

export class ExcelWorkbookManager {
  private permissionChecker: PermissionChecker;
  private logger: Logger;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(
    permissionChecker: PermissionChecker,
    logger: Logger,
    activeWorkbooks: Map<string, ExcelJS.Workbook>
  ) {
    this.permissionChecker = permissionChecker;
    this.logger = logger;
    this.activeWorkbooks = activeWorkbooks;
  }

  /**
   * Open an Excel workbook
   */
  public async openWorkbook(filePath: string): Promise<OperationResult<WorkbookInfo>> {
    try {
      // Validate file access
      const validation = this.permissionChecker.validateFileAccess(filePath, 0, 'read');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);

      const filename = this.getFilename(filePath);
      this.activeWorkbooks.set(filename, workbook);

      const worksheets: WorksheetInfo[] = workbook.worksheets.map((ws, index) => ({
        name: ws.name,
        index: index,
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
        hidden: ws.state === 'hidden' || ws.state === 'veryHidden',
      }));

      this.logger.info(`Opened workbook: ${filename}`);

      return {
        success: true,
        data: {
          filename,
          worksheetCount: workbook.worksheets.length,
          worksheets,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to open workbook: ${message}`);
      return { success: false, error: message };
    }
  }

  /**
   * Create new workbook
   */
  public async createWorkbook(filename: string): Promise<OperationResult<WorkbookInfo>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Excel MCP';
      workbook.created = new Date();

      const worksheet = workbook.addWorksheet('Sheet1');

      this.activeWorkbooks.set(filename, workbook);

      this.logger.info(`Created new workbook: ${filename}`);

      return {
        success: true,
        data: {
          filename,
          worksheetCount: 1,
          worksheets: [{
            name: worksheet.name,
            index: 0,
            rowCount: 0,
            columnCount: 0,
            hidden: false,
          }],
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Save workbook
   */
  public async saveWorkbook(filename: string, outputPath?: string): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const savePath = outputPath || filename;
      
      // Validate output path
      const pathValidation = this.permissionChecker.isPathAllowed(savePath);
      if (!pathValidation.success) {
        return { success: false, error: pathValidation.error };
      }

      await workbook.xlsx.writeFile(savePath);

      this.logger.info(`Saved workbook: ${savePath}`);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Close workbook
   */
  public closeWorkbook(filename: string): OperationResult<void> {
    if (this.activeWorkbooks.has(filename)) {
      this.activeWorkbooks.delete(filename);
      this.logger.info(`Closed workbook: ${filename}`);
      return { success: true };
    }
    return { success: false, error: `Workbook "${filename}" not found` };
  }

  /**
   * Get workbook by filename
   */
  public getWorkbook(filename: string): ExcelJS.Workbook | undefined {
    return this.activeWorkbooks.get(filename);
  }

  private getFilename(filePath: string): string {
    const parts = filePath.replace(/\\/g, '/').split('/');
    return parts[parts.length - 1];
  }
}