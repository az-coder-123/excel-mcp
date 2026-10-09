/**
 * Excel Workbook Manager - Handles workbook lifecycle operations
 * Single Responsibility: Workbook open, create, save, close operations
 */

import ExcelJS from 'exceljs';
import { PermissionChecker } from '../security/permission-checker.js';
import { OperationResult, WorkbookInfo, WorksheetInfo } from '../types/index.js';
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

      // Auto-save the workbook to disk
      await workbook.xlsx.writeFile(filename);

      this.logger.info(`Created and saved new workbook: ${filename}`);

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
      this.logger.error(`Failed to create workbook: ${message}`);
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
   * Export worksheet to new file - read file directly and write
   */
  public async exportWorksheetToNewFile(
    filename: string,
    worksheetName: string,
    newFilePath: string
  ): Promise<OperationResult<void>> {
    try {
      // Get the original file path from filename
      // For now, we'll save the in-memory workbook

      const sourceWorkbook = this.activeWorkbooks.get(filename);
      if (!sourceWorkbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const sourceWorksheet = sourceWorkbook.getWorksheet(worksheetName);
      if (!sourceWorksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // Create a fresh workbook and copy data cell by cell
      const newWorkbook = new ExcelJS.Workbook();
      const newWorksheet = newWorkbook.addWorksheet(worksheetName);

      // Copy all data without any formatting
      sourceWorksheet.eachRow((row) => {
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, cell => {
          rowData.push(cell.value);
        });
        newWorksheet.addRow(rowData);
      });

      // Save the new workbook
      await newWorkbook.xlsx.writeFile(newFilePath);

      this.logger.info(`Exported worksheet "${worksheetName}" to: ${newFilePath}`);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to export worksheet: ${message}`);
      return { success: false, error: message };
    }
  }

  /**
   * Get workbook context
   */
  public getWorkbookContext(filename: string): OperationResult<{
    isOpen: boolean;
    filePath: string;
    worksheets: WorksheetInfo[];
    currentWorksheet: string | null;
  }> {
    const workbook = this.activeWorkbooks.get(filename);

    if (!workbook) {
      return {
        success: true,
        data: {
          isOpen: false,
          filePath: filename,
          worksheets: [],
          currentWorksheet: null,
        },
      };
    }

    const worksheets: WorksheetInfo[] = workbook.worksheets.map((ws, index) => ({
      name: ws.name,
      index: index,
      rowCount: ws.rowCount,
      columnCount: ws.columnCount,
      hidden: ws.state === 'hidden' || ws.state === 'veryHidden',
    }));

    return {
      success: true,
      data: {
        isOpen: true,
        filePath: filename,
        worksheets,
        currentWorksheet: worksheets[0]?.name || null,
      },
    };
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
