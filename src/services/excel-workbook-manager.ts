/**
 * Excel Workbook Manager - Handles workbook lifecycle operations
 * Single Responsibility: Workbook open, create, save, close operations
 */

import ExcelJS from 'exceljs';
import path from 'node:path';
import { PermissionChecker } from '../security/permission-checker.js';
import { OperationResult, WorkbookInfo, WorksheetInfo } from '../types/index.js';
import { Logger } from '../utils/logger.js';

export class ExcelWorkbookManager {
  /** Maximum in-memory workbooks before LRU eviction kicks in (audit Finding 1.2). */
  private static readonly MAX_ACTIVE_WORKBOOKS = 10;

  private permissionChecker: PermissionChecker;
  private logger: Logger;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;
  /** Tracks the on-disk origin of each open workbook so save-without-path returns there (Finding 1.3). */
  private sourcePaths: Map<string, string> = new Map();

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
      this.registerWorkbook(filename, workbook, path.resolve(filePath));

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
      // Security: enforce full validation (permission + path + extension) on write targets
      const validation = this.permissionChecker.validateFileAccess(filename, 0, 'write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Excel MCP';
      workbook.created = new Date();

      const worksheet = workbook.addWorksheet('Sheet1');

      this.registerWorkbook(filename, workbook, path.resolve(filename));

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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }
      this.touchWorkbook(filename);

      // Without an explicit output path, save back to the workbook's original
      // on-disk location instead of resolving against process.cwd() (Finding 1.3)
      const savePath = outputPath || this.sourcePaths.get(filename) || filename;

      // Security: enforce full validation (permission + path + extension) on the save target
      const validation = this.permissionChecker.validateFileAccess(savePath, 0, 'write');
      if (!validation.success) {
        return { success: false, error: validation.error };
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
      this.sourcePaths.delete(filename);
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
      // Security: enforce full validation (permission + path + extension) on the export target
      const validation = this.permissionChecker.validateFileAccess(newFilePath, 0, 'write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

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
      this.touchWorkbook(filename);

      // Preserve ALL worksheet fidelity (values, formulas, styles, merged cells,
      // column widths) by round-tripping the in-memory workbook through a buffer
      // and removing the other worksheets from the isolated copy (audit Finding 6.2).
      const buffer = await sourceWorkbook.xlsx.writeBuffer();
      const exportWorkbook = new ExcelJS.Workbook();
      await exportWorkbook.xlsx.load(buffer);
      for (const ws of [...exportWorkbook.worksheets]) {
        if (ws.name !== worksheetName) {
          exportWorkbook.removeWorksheet(ws.id);
        }
      }

      // Save the new workbook
      await exportWorkbook.xlsx.writeFile(newFilePath);

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
    const workbook = this.activeWorkbooks.get(filename);
    if (workbook) {
      this.touchWorkbook(filename);
    }
    return workbook;
  }

  /**
   * Register a workbook and enforce the LRU cap.
   * Insertion order of a Map doubles as the LRU order: recently used entries
   * are re-inserted by touchWorkbook(), so the first key is always the
   * least-recently-used candidate for eviction (audit Finding 1.2).
   */
  private registerWorkbook(key: string, workbook: ExcelJS.Workbook, sourcePath: string): void {
    // The same physical file may already be open under a different map key
    // (createWorkbook stores under the path-as-passed while openWorkbook uses
    // the basename). Reuse the existing entry instead of inserting a duplicate
    // so LRU accounting stays accurate (see audit Finding 1.3).
    for (const [existingKey, existingPath] of this.sourcePaths) {
      if (existingPath === sourcePath) {
        key = existingKey;
        break;
      }
    }

    if (
      this.activeWorkbooks.has(key) &&
      this.sourcePaths.get(key) !== sourcePath
    ) {
      // Same basename opened from a different directory overwrites the earlier
      // entry (known limitation tracked for the Phase 3 tool-layer rewrite).
      this.logger.warn(
        `Workbook key collision: "${key}" already open from "${this.sourcePaths.get(key)}", replacing with "${sourcePath}"`
      );
    }
    // Re-registration must refresh LRU order: Map.set() on an existing key
    // keeps its ORIGINAL insertion position, so delete first to move it last.
    if (this.activeWorkbooks.has(key)) {
      this.activeWorkbooks.delete(key);
      this.sourcePaths.delete(key);
    }
    this.activeWorkbooks.set(key, workbook);
    this.sourcePaths.set(key, sourcePath);

    while (this.activeWorkbooks.size > ExcelWorkbookManager.MAX_ACTIVE_WORKBOOKS) {
      const oldest = this.activeWorkbooks.keys().next().value;
      if (oldest === undefined) break;
      this.activeWorkbooks.delete(oldest);
      this.sourcePaths.delete(oldest);
      this.logger.warn(
        `Evicted least-recently-used workbook "${oldest}" (limit ${ExcelWorkbookManager.MAX_ACTIVE_WORKBOOKS})`
      );
    }
  }

  /** Move a workbook to the most-recently-used position. */
  private touchWorkbook(key: string): void {
    const workbook = this.activeWorkbooks.get(key);
    if (workbook) {
      this.activeWorkbooks.delete(key);
      this.activeWorkbooks.set(key, workbook);
    }
  }

  private getFilename(filePath: string): string {
    const parts = filePath.replace(/\\/g, '/').split('/');
    return parts[parts.length - 1];
  }
}
