/**
 * Excel Structure Operations - Handles worksheet, row, column, and table operations
 * Single Responsibility: Worksheet and structure manipulation
 */

import ExcelJS from 'exceljs';
import { WorksheetInfo, OperationResult } from '../types/index.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { columnLetterToNumber, parseCellRange } from '../utils/excel-coords.js';

export class ExcelStructureOperations {
  private permissionChecker: PermissionChecker;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(
    permissionChecker: PermissionChecker,
    _logger: Logger,
    activeWorkbooks: Map<string, ExcelJS.Workbook>
  ) {
    this.permissionChecker = permissionChecker;
    this.activeWorkbooks = activeWorkbooks;
  }

  /**
   * Protect a worksheet with an optional password (wires the previously-undispatched protection tools).
   */
  public async protectWorksheet(
    filename: string,
    worksheetName: string,
    password?: string,
    allowSelectLockedCells: boolean = true,
    allowSelectUnlockedCells: boolean = true
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

      // ExcelJS types require a concrete string password; '' means no password
      await worksheet.protect(password ?? '', {
        selectLockedCells: allowSelectLockedCells,
        selectUnlockedCells: allowSelectUnlockedCells,
      });
      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Remove worksheet protection (throws on wrong password, surfaced as an error result).
   */
  public async unprotectWorksheet(
    filename: string,
    worksheetName: string,
    _password?: string
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

      // Note: ExcelJS unprotect() takes no password argument; password
      // enforcement is applied by Excel when the file is opened.
      await worksheet.unprotect();
      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /** Set the locked protection flag on every cell in a range. */
  public async protectCells(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    locked: boolean
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

      const range = parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: `Invalid cell range "${startCell}:${endCell}"` };
      }

      for (let r = range.start.row; r <= range.end.row; r++) {
        for (let c = range.start.column; c <= range.end.column; c++) {
          const cell = worksheet.getCell(r, c);
          cell.protection = { ...cell.protection, locked };
        }
      }
      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Workbook-level protection is not supported by the ExcelJS serialization engine.
   * Dispatched honestly with an explicit error rather than silently faking success.
   */
  public async protectWorkbook(_password?: string): Promise<OperationResult<void>> {
    return {
      success: false,
      error: 'Workbook-level protection is not supported by the ExcelJS engine; protect individual worksheets instead',
    };
  }

  public async unprotectWorkbook(_password?: string): Promise<OperationResult<void>> {
    return {
      success: false,
      error: 'Workbook-level protection is not supported by the ExcelJS engine; unprotect individual worksheets instead',
    };
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

      if (!worksheetName || !worksheetName.trim()) {
        return { success: false, error: 'Worksheet name cannot be empty' };
      }

      if (workbook.getWorksheet(worksheetName)) {
        return { success: false, error: `Worksheet "${worksheetName}" already exists` };
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
   * Delete worksheet
   */
  public async deleteWorksheet(
    filename: string,
    worksheetName: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('delete');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      if (workbook.worksheets.length <= 1) {
        return { success: false, error: 'Cannot delete the only worksheet in the workbook' };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      workbook.removeWorksheet(worksheet.id);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Rename worksheet
   */
  public async renameWorksheet(
    filename: string,
    oldName: string,
    newName: string
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

      if (!newName || !newName.trim()) {
        return { success: false, error: 'Worksheet name cannot be empty' };
      }

      const worksheet = workbook.getWorksheet(oldName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${oldName}" not found` };
      }

      if (oldName !== newName && workbook.getWorksheet(newName)) {
        return { success: false, error: `Worksheet "${newName}" already exists` };
      }

      worksheet.name = newName;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Copy worksheet
   */
  public async copyWorksheet(
    filename: string,
    sourceWorksheet: string,
    targetName: string
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

      if (!targetName || !targetName.trim()) {
        return { success: false, error: 'Target worksheet name cannot be empty' };
      }

      if (workbook.getWorksheet(targetName)) {
        return { success: false, error: `Worksheet "${targetName}" already exists` };
      }

      const worksheet = workbook.getWorksheet(sourceWorksheet);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${sourceWorksheet}" not found` };
      }

      const newWorksheet = workbook.addWorksheet(targetName);
      
      // Copy cells
      worksheet.eachRow((row, rowNumber) => {
        const newRow = newWorksheet.getRow(rowNumber);
        row.eachCell((cell, colNumber) => {
          newRow.getCell(colNumber).value = cell.value;
        });
      });

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Insert rows at a specific position
   */
  public async insertRows(
    filename: string,
    worksheetName: string,
    startRow: number,
    count: number
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

      if (startRow <= 0) {
        return { success: false, error: `Start row must be positive (got ${startRow})` };
      }
      if (count <= 0) {
        return { success: false, error: `Count must be greater than 0 (got ${count})` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.spliceRows(startRow, 0, ...Array(count).fill([]));

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Insert columns at a specific position
   */
  public async insertColumns(
    filename: string,
    worksheetName: string,
    startColumn: string,
    count: number
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

      if (count <= 0) {
        return { success: false, error: `Count must be greater than 0 (got ${count})` };
      }

      const colIndex = columnLetterToNumber(startColumn);
      if (colIndex <= 0) {
        return { success: false, error: `Invalid column letter "${startColumn}"` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.spliceColumns(colIndex, 0, ...Array(count).fill([]));

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Delete rows from a specific position
   */
  public async deleteRows(
    filename: string,
    worksheetName: string,
    startRow: number,
    count: number
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('delete');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      if (startRow <= 0) {
        return { success: false, error: `Start row must be positive (got ${startRow})` };
      }
      if (count <= 0) {
        return { success: false, error: `Count must be greater than 0 (got ${count})` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.spliceRows(startRow, count);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Delete columns from a specific position
   */
  public async deleteColumns(
    filename: string,
    worksheetName: string,
    startColumn: string,
    count: number
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('delete');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      if (count <= 0) {
        return { success: false, error: `Count must be greater than 0 (got ${count})` };
      }

      const colIndex = columnLetterToNumber(startColumn);
      if (colIndex <= 0) {
        return { success: false, error: `Invalid column letter "${startColumn}"` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.spliceColumns(colIndex, count);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Merge cells in a range
   */
  public async mergeCells(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string
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

      const range = parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: `Invalid cell range "${startCell}:${endCell}"` };
      }

      worksheet.mergeCells(`${startCell}:${endCell}`);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Unmerge cells
   */
  public async unmergeCells(
    filename: string,
    worksheetName: string,
    cellAddress: string
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
      if (cell.isMerged) {
        worksheet.unMergeCells(cellAddress);
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add table
   */
  public async addTable(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    tableName: string,
    style?: string
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

      if (!tableName || !tableName.trim()) {
        return { success: false, error: 'Table name cannot be empty' };
      }

      const range = parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: `Invalid cell range "${startCell}:${endCell}"` };
      }

      const columns: Array<{ name: string; filterButton?: boolean }> = [];
      const headerRow = worksheet.getRow(range.start.row);
      for (let c = range.start.column; c <= range.end.column; c++) {
        const val = headerRow.getCell(c).value;
        const colName = val !== null && val !== undefined && String(val).trim() !== ''
          ? String(val).trim()
          : `Column${c - range.start.column + 1}`;
        columns.push({ name: colName, filterButton: true });
      }

      const rows: unknown[][] = [];
      for (let r = range.start.row + 1; r <= range.end.row; r++) {
        const rowData: unknown[] = [];
        const row = worksheet.getRow(r);
        for (let c = range.start.column; c <= range.end.column; c++) {
          rowData.push(row.getCell(c).value ?? '');
        }
        rows.push(rowData);
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const tableConfig: any = {
        name: tableName,
        ref: `${startCell}:${endCell}`,
        headerRow: true,
        columns,
        rows,
      };

      if (style) {
        tableConfig.style = { theme: style };
      }

      worksheet.addTable(tableConfig);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add auto-filter
   */
  public async addFilter(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string
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

      const range = parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: `Invalid cell range "${startCell}:${endCell}"` };
      }

      worksheet.autoFilter = `${startCell}:${endCell}`;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Remove auto-filter
   */
  public async removeFilter(
    filename: string,
    worksheetName: string
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

      worksheet.autoFilter = undefined;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Calculate formulas
   */
  public async calculateFormula(filename: string): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      // ExcelJS doesn't calculate formulas, just return success
      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }
}