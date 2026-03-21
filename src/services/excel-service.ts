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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const colIndex = this.columnLetterToNumber(startColumn);
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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const colIndex = this.columnLetterToNumber(startColumn);
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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
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
   * Copy a range of cells
   */
  public async copyRange(
    filename: string,
    worksheetName: string,
    sourceStart: string,
    sourceEnd: string,
    targetStart: string,
    targetWorksheet?: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const targetWs = targetWorksheet 
        ? workbook.getWorksheet(targetWorksheet) 
        : worksheet;
      
      if (!targetWs) {
        return { success: false, error: `Target worksheet not found` };
      }

      // Read source range
      const sourceRange = this.parseCellRange(sourceStart, sourceEnd);
      if (!sourceRange) {
        return { success: false, error: 'Invalid source range' };
      }

      // Get target start position
      const targetPos = this.parseCellAddress(targetStart);
      if (!targetPos) {
        return { success: false, error: 'Invalid target address' };
      }

      // Copy cells
      for (let row = sourceRange.start.row; row <= sourceRange.end.row; row++) {
        for (let col = sourceRange.start.column; col <= sourceRange.end.column; col++) {
          const sourceCell = worksheet.getCell(row, col);
          const targetRow = targetPos.row + (row - sourceRange.start.row);
          const targetCol = targetPos.column + (col - sourceRange.start.column);
          const targetCell = targetWs.getCell(targetRow, targetCol);
          
          targetCell.value = sourceCell.value;
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set cell formatting
   */
  public async setCellFormat(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    format: Record<string, unknown>
  ): Promise<OperationResult<void>> {
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
      this.applyFormat(cell, format);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set range formatting
   */
  public async setRangeFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    format: Record<string, unknown>
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const range = this.parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: 'Invalid cell range' };
      }

      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          this.applyFormat(cell, format);
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Auto-fit column widths
   */
  public async autoFitColumns(
    filename: string,
    worksheetName: string,
    startColumn?: string,
    endColumn?: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // ExcelJS doesn't have autoFit, so we set a reasonable default width
      const start = startColumn ? this.columnLetterToNumber(startColumn) : 1;
      const end = endColumn ? this.columnLetterToNumber(endColumn) : worksheet.columnCount;

      for (let col = start; col <= end; col++) {
        worksheet.getColumn(col).width = 15;
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set column width
   */
  public async setColumnWidth(
    filename: string,
    worksheetName: string,
    column: string,
    width: number
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const colIndex = this.columnLetterToNumber(column);
      worksheet.getColumn(colIndex).width = width;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set row height
   */
  public async setRowHeight(
    filename: string,
    worksheetName: string,
    row: number,
    height: number
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.getRow(row).height = height;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add data validation
   */
  public async addDataValidation(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    type: string,
    formula1: string,
    formula2?: string,
    operator?: string,
    errorMessage?: string
  ): Promise<OperationResult<void>> {
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
      
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const validation: any = {
        type: type,
        formulae: formula2 ? [formula1, formula2] : [formula1],
      };

      if (operator) {
        validation.operator = operator;
      }

      if (errorMessage) {
        validation.error = errorMessage;
      }

      cell.dataValidation = validation;

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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const tableConfig: any = {
        name: tableName,
        ref: `${startCell}:${endCell}`,
        headerRow: true,
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
   * Freeze panes
   */
  public async freezePanes(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.views = [{
        state: 'frozen',
        xSplit: this.columnLetterToNumber(cellAddress.replace(/[0-9]/g, '')) - 1,
        ySplit: parseInt(cellAddress.replace(/[A-Z]/gi, '')) - 1,
      }];

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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
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
   * Delete worksheet
   */
  public async deleteWorksheet(
    filename: string,
    worksheetName: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(oldName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${oldName}" not found` };
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
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
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
   * Set print area
   */
  public async setPrintArea(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.pageSetup.printArea = `${startCell}:${endCell}`;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add header and footer
   */
  public async addHeaderFooter(
    filename: string,
    worksheetName: string,
    header?: string,
    footer?: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      worksheet.headerFooter = {
        oddHeader: header,
        oddFooter: footer,
      };

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add hyperlink
   */
  public async addHyperlink(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    url: string,
    display?: string
  ): Promise<OperationResult<void>> {
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
      cell.value = {
        text: display || url,
        hyperlink: url,
      };

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add comment to cell
   */
  public async addComment(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    comment: string,
    author?: string
  ): Promise<OperationResult<void>> {
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const note: any = {
        texts: [{ text: comment }],
      };
      if (author) {
        note.author = author;
      }
      cell.note = note;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Remove comment from cell
   */
  public async removeComment(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<OperationResult<void>> {
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (cell as any).note = undefined;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Find and replace text
   */
  public async findReplace(
    filename: string,
    worksheetName: string,
    findText: string,
    replaceText: string,
    matchCase?: boolean,
    matchEntireCell?: boolean
  ): Promise<OperationResult<{ count: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      let count = 0;
      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          if (typeof cell.value === 'string') {
            let shouldReplace = false;
            
            if (matchEntireCell) {
              shouldReplace = matchCase 
                ? cell.value === findText
                : cell.value.toLowerCase() === findText.toLowerCase();
            } else {
              shouldReplace = matchCase
                ? cell.value.includes(findText)
                : cell.value.toLowerCase().includes(findText.toLowerCase());
            }

            if (shouldReplace) {
              cell.value = replaceText;
              count++;
            }
          }
        });
      });

      return { success: true, data: { count } };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Sort range
   */
  public async sortRange(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    sortColumn: number,
    ascending: boolean
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // Read all data
      const range = this.parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: 'Invalid cell range' };
      }

      const data: Array<{ row: number; values: unknown[] }> = [];
      for (let row = range.start.row; row <= range.end.row; row++) {
        const values: unknown[] = [];
        for (let col = range.start.column; col <= range.end.column; col++) {
          values.push(worksheet.getCell(row, col).value);
        }
        data.push({ row, values });
      }

      // Sort data
      data.sort((a, b) => {
        const aVal = a.values[sortColumn - 1];
        const bVal = b.values[sortColumn - 1];
        
        if (aVal === bVal) return 0;
        if (aVal === null || aVal === undefined) return 1;
        if (bVal === null || bVal === undefined) return -1;
        
        const comparison = aVal < bVal ? -1 : 1;
        return ascending ? comparison : -comparison;
      });

      // Write sorted data back
      for (let i = 0; i < data.length; i++) {
        const rowData = data[i];
        for (let col = range.start.column; col <= range.end.column; col++) {
          worksheet.getCell(range.start.row + i, col).value = 
            rowData.values[col - range.start.column] as ExcelJS.CellValue;
        }
      }

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

  /**
   * Get named ranges
   */
  public async getNamedRanges(filename: string): Promise<OperationResult<Array<{ name: string; ref: string }>>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      const names = workbook.definedNames?.model || [];
      const result = names.map((n: { name: string; address?: string }) => ({
        name: n.name,
        ref: n.address || '',
      }));

      return { success: true, data: result };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add named range
   */
  public async addNamedRange(
    filename: string,
    worksheetName: string,
    name: string,
    startCell: string,
    endCell: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not opened` };
      }

      workbook.definedNames?.add(name, `'${worksheetName}'!${startCell}:${endCell}`);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
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

  private columnLetterToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }

  private parseCellAddress(address: string): { row: number; column: number } | null {
    const match = address.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return null;

    const column = this.columnLetterToNumber(match[1].toUpperCase());
    const row = parseInt(match[2], 10);

    return { row, column };
  }

  private parseCellRange(startCell: string, endCell: string): CellRange | null {
    const start = this.parseCellAddress(startCell);
    const end = this.parseCellAddress(endCell);

    if (!start || !end) return null;

    return { start, end };
  }

  private applyFormat(cell: ExcelJS.Cell, format: Record<string, unknown>): void {
    if (format.bold !== undefined) {
      cell.font = { ...cell.font, bold: format.bold as boolean };
    }
    if (format.italic !== undefined) {
      cell.font = { ...cell.font, italic: format.italic as boolean };
    }
    if (format.underline !== undefined) {
      cell.font = { ...cell.font, underline: format.underline as boolean };
    }
    if (format.fontSize !== undefined) {
      cell.font = { ...cell.font, size: format.fontSize as number };
    }
    if (format.fontColor !== undefined) {
      cell.font = { ...cell.font, color: { argb: format.fontColor as string } };
    }
    if (format.backgroundColor !== undefined) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: format.backgroundColor as string },
      };
    }
    if (format.horizontalAlignment !== undefined) {
      cell.alignment = { ...cell.alignment, horizontal: format.horizontalAlignment as ExcelJS.Alignment['horizontal'] };
    }
    if (format.verticalAlignment !== undefined) {
      cell.alignment = { ...cell.alignment, vertical: format.verticalAlignment as ExcelJS.Alignment['vertical'] };
    }
    if (format.wrapText !== undefined) {
      cell.alignment = { ...cell.alignment, wrapText: format.wrapText as boolean };
    }
    if (format.numberFormat !== undefined) {
      cell.numFmt = format.numberFormat as string;
    }
    if (format.borderColor !== undefined) {
      cell.border = {
        top: { style: 'thin', color: { argb: format.borderColor as string } },
        left: { style: 'thin', color: { argb: format.borderColor as string } },
        bottom: { style: 'thin', color: { argb: format.borderColor as string } },
        right: { style: 'thin', color: { argb: format.borderColor as string } },
      };
    }
  }
}
