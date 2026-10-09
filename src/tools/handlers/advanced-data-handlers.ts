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

    // Read range data
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

    // Read range data
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
    const listResult = await this.excelService.getWorksheets(filename);
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

    // Read entire column
    const endCell = this.columnNumberToLetter(columnNumber) + sheetInfo.rowCount;
    const range = this.parseCellRange(`${column}1`, endCell);
    if (!range) {
      return { success: false, error: 'Invalid column range' };
    }

    // Read column data
    const readResult = await this.excelService.readRange(filename, worksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    // Find duplicates in column
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

  /**
   * Split data from one worksheet into multiple worksheets based on column values
   */
  async handleSplitDataToWorksheets(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const sourceWorksheet = this.getStringArg(args, 'sourceWorksheet');
    const columnLetter = this.getStringArg(args, 'columnLetter');
    const sheetNames = this.getObjectArg(args, 'sheetNames');
    const includeHeader = this.getBooleanArg(args, 'includeHeader') ?? true;

    if (!filename || !sourceWorksheet || !columnLetter) {
      return { success: false, error: 'Missing required parameters' };
    }

    // Parse column to number
    const columnNumber = this.columnLetterToNumber(columnLetter);
    if (!columnNumber) {
      return { success: false, error: 'Invalid column format' };
    }

    // Get worksheet info to determine range
    const listResult = await this.excelService.getWorksheets(filename);
    if (!listResult.success || !listResult.data) {
      return listResult;
    }

    const sheets = listResult.data as Array<{ name: string; rowCount: number }>;
    const sheetInfo = sheets.find(s => s.name === sourceWorksheet);
    if (!sheetInfo) {
      return { success: false, error: 'Source worksheet not found' };
    }

    // Read all data from source worksheet
    const endCell = this.columnNumberToLetter(sheetInfo.rowCount) + sheetInfo.rowCount;
    const range = this.parseCellRange('A1', endCell);
    if (!range) {
      return { success: false, error: 'Invalid range format' };
    }

    const readResult = await this.excelService.readRange(filename, sourceWorksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    const data = readResult.data as Array<Array<{ value: unknown }>>;
    if (data.length === 0) {
      return { success: false, error: 'No data found' };
    }

    // Group data by column value
    const groupedData = new Map<string, Array<Array<unknown>>>();
    const headerRow = includeHeader ? data[0] : null;

    const startRow = includeHeader ? 1 : 0;
    for (let row = startRow; row < data.length; row++) {
      const rowValues = data[row];
      const cellValue = rowValues[columnNumber - 1];
      const stringValue = cellValue !== undefined && cellValue !== null ? String(cellValue.value) : '';
      
      if (stringValue) {
        if (!groupedData.has(stringValue)) {
          groupedData.set(stringValue, []);
        }
        groupedData.get(stringValue)!.push(rowValues);
      }
    }

    // Create worksheets for each group
    const results: Array<{ sheetName: string; rowCount: number }> = [];
    
    for (const [value, rows] of groupedData.entries()) {
      const targetSheetName: string = sheetNames && sheetNames[value] ? String(sheetNames[value]) : value;
      
      // Create worksheet
      const createResult = await this.excelService.addWorksheet(filename, targetSheetName);
      if (!createResult.success) {
        return createResult;
      }

      // Write data
      let startRow = 1;
      if (includeHeader && headerRow) {
        const headerResult = await this.excelService.writeBatch(
          filename,
          targetSheetName,
          this.convertHeaderToBatch(headerRow, 1)
        );
        if (!headerResult.success) {
          return headerResult;
        }
        startRow = 2;
      }

      // Write data rows
      for (let i = 0; i < rows.length; i++) {
        const rowData = rows[i] as Array<unknown>;
        for (let col = 0; col < rowData.length; col++) {
          const cellAddress = this.columnNumberToLetter(col + 1) + (startRow + i);
          const cellValue = rowData[col] !== undefined ? rowData[col] : '';
          
          const writeResult = await this.excelService.writeCell(
            filename,
            targetSheetName,
            cellAddress,
            String(cellValue)
          );
          
          if (!writeResult.success) {
            return writeResult;
          }
        }
      }

      results.push({
        sheetName: targetSheetName,
        rowCount: rows.length
      });
    }

    return {
      success: true,
      data: {
        sheetsCreated: results.length,
        details: results
      }
    };
  }

  /**
   * Convert header row to batch format
   */
  private convertHeaderToBatch(headerRow: Array<{ value: unknown }>, row: number): Array<{ cellAddress: string; value: string }> {
    return headerRow.map((cell, index) => ({
      cellAddress: this.columnNumberToLetter(index + 1) + row,
      value: cell !== undefined && cell !== null ? String(cell.value) : ''
    }));
  }

  /**
   * Get unique values from a specific column
   */
  async handleGetUniqueValues(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const columnLetter = this.getStringArg(args, 'columnLetter');
    const hasHeader = this.getBooleanArg(args, 'hasHeader') ?? true;

    if (!filename || !worksheet || !columnLetter) {
      return { success: false, error: 'Missing required parameters' };
    }

    // Get worksheet info
    const listResult = await this.excelService.getWorksheets(filename);
    if (!listResult.success || !listResult.data) {
      return listResult;
    }

    const sheets = listResult.data as Array<{ name: string; rowCount: number }>;
    const sheetInfo = sheets.find(s => s.name === worksheet);
    if (!sheetInfo) {
      return { success: false, error: 'Worksheet not found' };
    }

    // Parse column to number
    const columnNumber = this.columnLetterToNumber(columnLetter);
    if (!columnNumber) {
      return { success: false, error: 'Invalid column format' };
    }

    // Read entire column
    const endCell = this.columnNumberToLetter(columnNumber) + sheetInfo.rowCount;
    const range = this.parseCellRange(`${columnLetter}1`, endCell);
    if (!range) {
      return { success: false, error: 'Invalid column range' };
    }

    const readResult = await this.excelService.readRange(filename, worksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    const data = readResult.data as Array<Array<{ value: unknown }>>;
    const startRow = hasHeader ? 1 : 0;

    // Extract values
    const values: string[] = [];
    for (let row = startRow; row < data.length; row++) {
      const cell = data[row][0];
      if (cell && cell.value !== undefined && cell.value !== null) {
        values.push(String(cell.value));
      }
    }

    // Get unique values
    const uniqueValues = [...new Set(values)];
    const valueCount = new Map<string, number>();
    
    for (const value of values) {
      valueCount.set(value, (valueCount.get(value) || 0) + 1);
    }

    const sortedUnique = uniqueValues.sort((a, b) => {
      const countDiff = (valueCount.get(b) || 0) - (valueCount.get(a) || 0);
      return countDiff !== 0 ? countDiff : a.localeCompare(b);
    });

    return {
      success: true,
      data: {
        column: columnLetter,
        totalValues: values.length,
        uniqueValues: sortedUnique,
        uniqueCount: sortedUnique.length,
        valueCounts: Object.fromEntries(valueCount.entries())
      }
    };
  }

  /**
   * Filter data based on conditions and copy to a new worksheet
   */
  async handleFilterData(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const sourceWorksheet = this.getStringArg(args, 'sourceWorksheet');
    const targetWorksheet = this.getStringArg(args, 'targetWorksheet');
    const filtersArg = this.getArrayArg(args, 'filters');
    const sourceRangeArg = this.getStringArg(args, 'sourceRange');
    const targetCellArg = this.getStringArg(args, 'targetCell') || 'A1';

    if (!filename || !sourceWorksheet || !targetWorksheet || !filtersArg) {
      return { success: false, error: 'Missing required parameters' };
    }

    // Parse filters
    const filters: Array<{ column: string; value: string }> = [];
    for (const filter of filtersArg) {
      if (typeof filter === 'object' && filter !== null) {
        const f = filter as Record<string, unknown>;
        const column = f.column as string;
        const value = f.value as string;
        filters.push({ column, value });
      }
    }

    // Determine source range
    let startCell = 'A1';
    let endCell = 'Z999';
    
    if (sourceRangeArg) {
      const range = this.parseCellRangeString(sourceRangeArg);
      if (range) {
        startCell = range.startCell;
        endCell = range.endCell;
      }
    } else {
      // Get worksheet info
      const listResult = await this.excelService.getWorksheets(filename);
      if (!listResult.success || !listResult.data) {
        return listResult;
      }

      const sheets = listResult.data as Array<{ name: string; rowCount: number }>;
      const sheetInfo = sheets.find(s => s.name === sourceWorksheet);
      if (sheetInfo) {
        endCell = 'Z' + sheetInfo.rowCount;
      }
    }

    const range = this.parseCellRange(startCell, endCell);
    if (!range) {
      return { success: false, error: 'Invalid range format' };
    }

    // Read source data
    const readResult = await this.excelService.readRange(filename, sourceWorksheet, range);
    if (!readResult.success || !readResult.data) {
      return readResult;
    }

    const data = readResult.data as Array<Array<{ value: unknown }>>;

    // Filter data
    const filteredRows: Array<Array<unknown>> = [];
    for (const row of data) {
      let matchesAllFilters = true;
      
      for (const filter of filters) {
        const colNumber = this.columnLetterToNumber(filter.column);
        if (!colNumber) {
          return { success: false, error: `Invalid column format: ${filter.column}` };
        }
        
        const cellValue = row[colNumber - 1];
        const stringValue = cellValue !== undefined && cellValue !== null ? String(cellValue.value) : '';
        
        if (stringValue !== filter.value) {
          matchesAllFilters = false;
          break;
        }
      }
      
      if (matchesAllFilters) {
        filteredRows.push(row);
      }
    }

    if (filteredRows.length === 0) {
      return {
        success: true,
        data: {
          message: 'No rows matched filter criteria',
          filteredRows: 0
        }
      };
    }

    // Create target worksheet
    const createResult = await this.excelService.addWorksheet(filename, targetWorksheet);
    if (!createResult.success) {
      return createResult;
    }

    // Parse target cell
    const targetCellParts = targetCellArg.match(/([A-Z]+)(\d+)/);
    const targetRow = targetCellParts ? parseInt(targetCellParts[2]) : 1;

    // Write filtered data
    for (let rowIndex = 0; rowIndex < filteredRows.length; rowIndex++) {
      const row = filteredRows[rowIndex];
      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const cellAddress = this.columnNumberToLetter(colIndex + 1) + (targetRow + rowIndex);
        const cellValue = row[colIndex] !== undefined ? row[colIndex] : '';
        
        const writeResult = await this.excelService.writeCell(
          filename,
          targetWorksheet,
          cellAddress,
          String(cellValue)
        );
        
        if (!writeResult.success) {
          return writeResult;
        }
      }
    }

    return {
      success: true,
      data: {
        sourceRange: `${startCell}:${endCell}`,
        targetWorksheet,
        filters,
        filteredRows: filteredRows.length,
        message: `Filtered ${filteredRows.length} rows from ${data.length} total rows`
      }
    };
  }

  /**
   * Parse cell range string to components
   */
  private parseCellRangeString(rangeStr: string): { startCell: string; endCell: string } | null {
    const parts = rangeStr.split(':');
    if (parts.length !== 2) {
      return null;
    }
    return {
      startCell: parts[0],
      endCell: parts[1]
    };
  }
}