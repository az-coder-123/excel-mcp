/**
 * Excel Service - Handles all Excel file operations
 * Single Responsibility: Excel file manipulation only
 */

import ExcelJS from 'exceljs';
import { 
  CellValue, 
  WorksheetInfo, 
  WorkbookInfo, 
  OperationResult,
  CellRange 
} from '../types/index.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';

export class ExcelService {
  private permissionChecker: PermissionChecker;
  private logger: Logger;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(permissionChecker: PermissionChecker, logger: Logger) {
    this.permissionChecker = permissionChecker;
    this.logger = logger;
    this.activeWorkbooks = new Map();
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
   * Read cell value
   */
  public async readCell(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<OperationResult<CellValue>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const cell = worksheet.getCell(cellAddress);
      const cellValue = this.extractCellValue(cell);

      return { success: true, data: cellValue };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Read a range of cells
   */
  public async readRange(
    filename: string,
    worksheetName: string,
    range: CellRange
  ): Promise<OperationResult<CellValue[][]>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const result: CellValue[][] = [];
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        const rowData: CellValue[] = [];
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          rowData.push(this.extractCellValue(cell));
        }
        result.push(rowData);
      }

      return { success: true, data: result };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Write value to cell
   */
  public async writeCell(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    value: unknown
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const cell = worksheet.getCell(cellAddress);
      cell.value = value as ExcelJS.CellValue;

      this.logger.info(`Wrote value to ${filename}!${worksheetName}!${cellAddress}`);

      return { success: true };
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
   * Add worksheet to workbook
   */
  public async addWorksheet(
    filename: string,
    worksheetName: string
  ): Promise<OperationResult<WorksheetInfo>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.addWorksheet(worksheetName);

      return {
        success: true,
        data: {
          name: worksheet.name,
          index: workbook.worksheets.length - 1,
          rowCount: 0,
          columnCount: 0,
          hidden: false,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Get list of worksheets
   */
  public getWorksheets(filename: string): OperationResult<WorksheetInfo[]> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheets: WorksheetInfo[] = workbook.worksheets.map((ws, index) => ({
        name: ws.name,
        index: index,
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
        hidden: ws.state === 'hidden' || ws.state === 'veryHidden',
      }));

      return { success: true, data: worksheets };
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

  // Private helper methods
  private extractCellValue(cell: ExcelJS.Cell): CellValue {
    const address = cell.address;
    let value: unknown = cell.value;
    let type: CellValue['type'] = 'string';
    let formula: string | undefined;

    if (cell.formula) {
      type = 'formula';
      formula = cell.formula;
      value = cell.result;
    } else if (cell.value instanceof Date) {
      type = 'date';
      value = cell.value.toISOString();
    } else if (typeof cell.value === 'number') {
      type = 'number';
    } else if (typeof cell.value === 'boolean') {
      type = 'boolean';
    } else if (cell.type === ExcelJS.ValueType.Error) {
      type = 'error';
    }

    return { address, value, type, formula };
  }

  private getFilename(filePath: string): string {
    const parts = filePath.replace(/\\/g, '/').split('/');
    return parts[parts.length - 1];
  }
}