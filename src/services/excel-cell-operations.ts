/**
 * Excel Cell Operations - Handles cell read/write operations
 * Single Responsibility: Cell and range data operations
 */

import ExcelJS from 'exceljs';
import { CellValue, CellRange, OperationResult } from '../types/index.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';

export class ExcelCellOperations {
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
   * Write batch values to cells
   */
  public async writeBatch(
    filename: string,
    worksheetName: string,
    data: Array<{ cellAddress: string; value: unknown }>
  ): Promise<OperationResult<{ count: number }>> {
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

      for (const item of data) {
        const cell = worksheet.getCell(item.cellAddress);
        cell.value = item.value as ExcelJS.CellValue;
      }

      this.logger.info(`Wrote ${data.length} values to ${filename}!${worksheetName}`);

      return { success: true, data: { count: data.length } };
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
      const validation = this.permissionChecker.hasPermission('write');
      if (!validation.success) {
        return { success: false, error: validation.error };
      }

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

  // Helper methods
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

  private columnLetterToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }
}