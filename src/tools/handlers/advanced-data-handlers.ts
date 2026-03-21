/**
 * Advanced Data Handlers
 * Single Responsibility: Advanced data operations like duplicate detection and analysis
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class AdvancedDataHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  /**
   * Find duplicate values in a range
   */
  async handleFindDuplicates(args: Record<string, unknown>): Promise<OperationResult> {
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

    // Read the range data
    const readResult = await this.excelService.readRange(filename, worksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    // Extract values and find duplicates
    const values: Array<{ value: string; row: number; column: number }> = [];
    const data = readResult.data as Array<Array<{ value: string }>>;

    for (let row = 0; row < data.length; row++) {
      for (let col = 0; col < data[row].length; col++) {
        const cell = data[row][col];
        if (cell && cell.value !== undefined && cell.value !== null) {
          values.push({
            value: String(cell.value),
            row: range.start.row + row,
            column: range.start.column + col
          });
        }
      }
    }

    // Find duplicates
    const valueMap = new Map<string, Array<{ row: number; column: number }>>();
    const duplicates: Array<{
      value: string;
      occurrences: Array<{ row: number; column: number; cellAddress: string }>;
      count: number;
    }> = [];

    for (const item of values) {
      if (!valueMap.has(item.value)) {
        valueMap.set(item.value, []);
      }
      valueMap.get(item.value)!.push({ row: item.row, column: item.column });
    }

    for (const [value, occurrences] of valueMap.entries()) {
      if (occurrences.length > 1) {
        const cellAddresses = occurrences.map(occ => ({
          row: occ.row,
          column: occ.column,
          cellAddress: this.columnNumberToLetter(occ.column) + occ.row
        }));

        duplicates.push({
          value,
          occurrences: cellAddresses,
          count: occurrences.length
        });
      }
    }

    return {
      success: true,
      data: {
        duplicates,
        totalDuplicates: duplicates.length,
        totalValues: values.length,
        uniqueValues: valueMap.size
      }
    };
  }

  /**
   * Count unique values in a range
   */
  async handleCountUniqueValues(args: Record<string, unknown>): Promise<OperationResult> {
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

    // Read the range data
    const readResult = await this.excelService.readRange(filename, worksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    // Extract values
    const values: string[] = [];
    const data = readResult.data as Array<Array<{ value: string }>>;

    for (let row = 0; row < data.length; row++) {
      for (let col = 0; col < data[row].length; col++) {
        const cell = data[row][col];
        if (cell && cell.value !== undefined && cell.value !== null) {
          values.push(String(cell.value));
        }
      }
    }

    // Count unique values
    const uniqueSet = new Set(values);
    const duplicateCount = values.length - uniqueSet.size;

    return {
      success: true,
      data: {
        totalValues: values.length,
        uniqueValues: uniqueSet.size,
        duplicateValues: duplicateCount,
        duplicatePercentage: values.length > 0 ? ((duplicateCount / values.length) * 100).toFixed(2) + '%' : '0%'
      }
    };
  }

  /**
   * Highlight duplicate values in a range using conditional formatting
   */
  async handleHighlightDuplicates(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const color = this.getStringArg(args, 'color') || 'FFFF0000'; // Default red

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    const range = this.parseCellRange(startCell, endCell);
    if (!range) {
      return { success: false, error: 'Invalid cell range format' };
    }

    // Find duplicates first
    const findResult = await this.handleFindDuplicates(args);
    if (!findResult.success || !findResult.data) {
      return findResult;
    }

    const duplicateData = findResult.data as { duplicates: Array<{ value: string; occurrences: Array<{ cellAddress: string }> }> };
    
    // Highlight each duplicate cell
    const format = {
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      },
      font: { color: { argb: 'FFFFFFFF' } } // White text for contrast
    };

    for (const duplicate of duplicateData.duplicates) {
      for (const occ of duplicate.occurrences) {
        const formatResult = await this.excelService.setCellFormat(
          filename,
          worksheet,
          occ.cellAddress,
          format
        );
        if (!formatResult.success) {
          return formatResult;
        }
      }
    }

    return {
      success: true,
      data: {
        highlightedCells: duplicateData.duplicates.length,
        message: `Highlighted ${duplicateData.duplicates.length} duplicate values with color ${color}`
      }
    };
  }

  /**
   * Get detailed duplicate information for a specific column
   */
  async handleGetDuplicateInfo(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const column = this.getStringArg(args, 'column');

    if (!filename || !worksheet || !column) {
      return { success: false, error: 'Missing required parameters' };
    }

    // Get worksheet row count
    const listResult = this.excelService.getWorksheets(filename);
    if (!listResult.success || !listResult.data) {
      return listResult;
    }

    const sheets = listResult.data as Array<{ name: string; rowCount: number }>;
    const sheetInfo = sheets.find(s => s.name === worksheet);
    if (!sheetInfo) {
      return { success: false, error: 'Worksheet not found' };
    }

    // Parse column letter to number
    const columnNumber = this.columnLetterToNumber(column);
    if (!columnNumber) {
      return { success: false, error: 'Invalid column format' };
    }

    // Read the entire column
    const endCell = this.columnNumberToLetter(columnNumber) + sheetInfo.rowCount;
    const range = this.parseCellRange(`${column}1`, endCell);
    if (!range) {
      return { success: false, error: 'Invalid column range' };
    }

    // Read the column data
    const readResult = await this.excelService.readRange(filename, worksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    // Find duplicates in the column
    const valueMap = new Map<string, number[]>();
    const data = readResult.data as Array<Array<{ value: string }>>;

    for (let row = 0; row < data.length; row++) {
      const cell = data[row][0];
      if (cell && cell.value !== undefined && cell.value !== null) {
        const value = String(cell.value);
        if (!valueMap.has(value)) {
          valueMap.set(value, []);
        }
        valueMap.get(value)!.push(row + 1); // Row numbers are 1-based
      }
    }

    // Get only duplicates
    const duplicates: Array<{
      value: string;
      rows: number[];
      count: number;
      cellAddresses: string[];
    }> = [];

    for (const [value, rows] of valueMap.entries()) {
      if (rows.length > 1) {
        duplicates.push({
          value,
          rows,
          count: rows.length,
          cellAddresses: rows.map(r => `${column}${r}`)
        });
      }
    }

    return {
      success: true,
      data: {
        column,
        duplicates,
        totalDuplicates: duplicates.length,
        totalValues: data.length,
        uniqueValues: valueMap.size
      }
    };
  }

  /**
   * Helper: Convert column number to letter (1 -> A, 27 -> AA)
   */
  private columnNumberToLetter(columnNumber: number): string {
    let letter = '';
    while (columnNumber > 0) {
      const remainder = (columnNumber - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      columnNumber = Math.floor((columnNumber - 1) / 26);
    }
    return letter;
  }

  /**
   * Helper: Convert column letter to number (A -> 1, AA -> 27)
   */
  private columnLetterToNumber(columnLetter: string): number | null {
    columnLetter = columnLetter.toUpperCase();
    let result = 0;
    for (let i = 0; i < columnLetter.length; i++) {
      const char = columnLetter.charCodeAt(i) - 64;
      if (char < 1 || char > 26) {
        return null;
      }
      result = result * 26 + char;
    }
    return result;
  }
}