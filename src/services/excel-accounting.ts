/**
 * Excel Accounting Service
 * Single Responsibility: Accounting-specific operations for financial data
 */

import ExcelJS from 'exceljs';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult } from '../types/index.js';
import {
  columnLetterToNumber,
  parseCellAddress,
  parseCellRange,
} from '../utils/excel-coords.js';

export class ExcelAccounting {
  private logger: Logger;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(
    _permissionChecker: PermissionChecker,
    logger: Logger,
    activeWorkbooks: Map<string, ExcelJS.Workbook>
  ) {
    this.logger = logger;
    this.activeWorkbooks = activeWorkbooks;
  }

  // ============================================
  // Financial Calculations
  // ============================================

  /**
   * Calculate sum with optional criteria
   */
  async calculateSum(
    filename: string,
    worksheetName: string,
    rangeStart: string,
    rangeEnd: string,
    criteriaColumn?: string,
    criteriaValue?: string
  ): Promise<OperationResult<{ sum: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const data = this.getRangeData(worksheet, rangeStart, rangeEnd);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      let sum = 0;

      if (criteriaColumn && criteriaValue) {
        // Sum with criteria
        const criteriaColIndex = this.getColumnIndex(rangeStart, criteriaColumn);

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          const cellVal = row[criteriaColIndex];
          const matches = cellVal !== null && cellVal !== undefined && (String(cellVal) === String(criteriaValue) || cellVal === criteriaValue);

          if (matches) {
            for (let c = 0; c < row.length; c++) {
              if (c !== criteriaColIndex && typeof row[c] === 'number') {
                sum += row[c] as number;
              }
            }
          }
        }
      } else {
        // Multi-column simple sum
        for (const row of data) {
          for (let c = 0; c < row.length; c++) {
            if (typeof row[c] === 'number') {
              sum += row[c] as number;
            }
          }
        }
      }

      sum = Math.round(sum * 1e6) / 1e6;

      this.logger.info(`Calculated sum for ${filename}!${worksheetName}: ${sum}`);
      return { success: true, data: { sum } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate sum: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate average
   */
  async calculateAverage(
    filename: string,
    worksheetName: string,
    rangeStart: string,
    rangeEnd: string
  ): Promise<OperationResult<{ average: number; count: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const data = this.getRangeData(worksheet, rangeStart, rangeEnd);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      let sum = 0;
      let count = 0;

      for (const row of data) {
        for (let c = 0; c < row.length; c++) {
          if (typeof row[c] === 'number') {
            sum += row[c] as number;
            count++;
          }
        }
      }

      const rawAvg = count > 0 ? sum / count : 0;
      const average = Math.round(rawAvg * 1e6) / 1e6;

      this.logger.info(`Calculated average for ${filename}!${worksheetName}: ${average}`);
      return { success: true, data: { average, count } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate average: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate running total
   */
  async calculateRunningTotal(
    filename: string,
    worksheetName: string,
    valueStartCell: string,
    valueEndCell: string,
    outputStartCell: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const valueData = this.getRangeData(worksheet, valueStartCell, valueEndCell);
      if (!valueData || valueData.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const valueColIndex = this.getColumnIndex(valueStartCell, valueStartCell.match(/^([A-Z]+)/)![1]);
      const outAddr = parseCellAddress(outputStartCell);
      if (!outAddr) {
        return { success: false, error: `Invalid outputStartCell "${outputStartCell}"` };
      }
      const outputColLetter = outAddr.colLetter;
      const outputStartRow = outAddr.row;

      let runningTotal = 0;

      for (let i = 0; i < valueData.length; i++) {
        const row = valueData[i];
        const value = typeof row[valueColIndex] === 'number' ? row[valueColIndex] as number : 0;
        runningTotal += value;

        const cell = worksheet.getCell(`${outputColLetter}${outputStartRow + i}`);
        cell.value = runningTotal;
      }

      this.logger.info(`Calculated running total for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate running total: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate percentage of total
   */
  async calculatePercentageOfTotal(
    filename: string,
    worksheetName: string,
    valueStartCell: string,
    valueEndCell: string,
    outputStartCell: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const valueData = this.getRangeData(worksheet, valueStartCell, valueEndCell);
      if (!valueData || valueData.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const valueColIndex = this.getColumnIndex(valueStartCell, valueStartCell.match(/^([A-Z]+)/)![1]);
      const outAddr = parseCellAddress(outputStartCell);
      if (!outAddr) {
        return { success: false, error: `Invalid outputStartCell "${outputStartCell}"` };
      }
      const outputColLetter = outAddr.colLetter;
      const outputStartRow = outAddr.row;

      // Calculate total
      let total = 0;
      for (const row of valueData) {
        if (typeof row[valueColIndex] === 'number') {
          total += row[valueColIndex] as number;
        }
      }

      // Calculate percentages
      for (let i = 0; i < valueData.length; i++) {
        const row = valueData[i];
        const value = typeof row[valueColIndex] === 'number' ? row[valueColIndex] as number : 0;
        const percentage = total !== 0 ? (value / total) * 100 : 0;

        const cell = worksheet.getCell(`${outputColLetter}${outputStartRow + i}`);
        cell.value = percentage;
        cell.numFmt = '0.00"%"';
      }

      this.logger.info(`Calculated percentage of total for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate percentage of total: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate year-to-date (YTD) values
   */
  async calculateYTD(
    filename: string,
    worksheetName: string,
    dateColumn: string,
    valueColumn: string,
    rangeStart: string,
    rangeEnd: string,
    outputColumn: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const range = parseCellRange(rangeStart, rangeEnd);
      if (!range) {
        return { success: false, error: 'Invalid data range' };
      }

      const data = this.getRangeData(worksheet, rangeStart, rangeEnd);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const dateColIndex = this.getColumnIndex(rangeStart, dateColumn);
      const valueColIndex = this.getColumnIndex(rangeStart, valueColumn);

      // Determine output column letter and starting row
      const outAddr = parseCellAddress(outputColumn);
      const outputColLetter = outAddr ? outAddr.colLetter : (outputColumn.match(/^([A-Za-z]+)/)?.[1]?.toUpperCase() ?? 'A');
      const outputStartRow = outAddr ? outAddr.row : range.start.row;

      let ytdTotal = 0;
      let currentYear = 0;

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const dateValue = row[dateColIndex];
        const value = typeof row[valueColIndex] === 'number' ? row[valueColIndex] as number : 0;

        // Extract year from date (Date object, Excel serial number, or date string)
        let year = 0;
        if (dateValue instanceof Date) {
          year = Number.isNaN(dateValue.getTime()) ? 0 : dateValue.getFullYear();
        } else if (typeof dateValue === 'number' && dateValue > 0) {
          // Excel serial date (days since 1899-12-30)
          const excelEpoch = new Date(1899, 11, 30);
          const date = new Date(excelEpoch.getTime() + dateValue * 86400000);
          year = Number.isNaN(date.getTime()) ? 0 : date.getFullYear();
        } else if (typeof dateValue === 'string' && dateValue.trim() !== '') {
          const parsed = new Date(dateValue);
          if (!Number.isNaN(parsed.getTime())) {
            year = parsed.getFullYear();
          }
        }

        // Reset YTD on year change
        if (year !== 0 && year !== currentYear) {
          ytdTotal = 0;
          currentYear = year;
        }

        ytdTotal += value;

        const cell = worksheet.getCell(`${outputColLetter}${outputStartRow + i}`);
        cell.value = ytdTotal;
      }

      this.logger.info(`Calculated YTD for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate YTD: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  // ============================================
  // Accounting Formats
  // ============================================

  /**
   * Apply accounting number format
   */
  async applyAccountingFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    options: {
      decimalPlaces?: number;
      useSeparator?: boolean;
      showNegativeInRed?: boolean;
      showNegativeInParentheses?: boolean;
      currencySymbol?: string;
    } = {}
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const useSeparator = options.useSeparator ?? true;
      const symbol = options.currencySymbol ?? '';
      const showParentheses = options.showNegativeInParentheses ?? true;
      const showRed = options.showNegativeInRed ?? true;
      const redTag = showRed ? '[Red]' : '';

      let formatCode = '';

      if (showParentheses) {
        // Format: (1,234.56) for negative with optional [Red]
        formatCode = `_(${symbol}* #,##0.00_);${redTag}_(${symbol} \\(#,##0.00\\);_(${symbol}*"-"_);_(@_)`;
      } else {
        formatCode = `${symbol}#,##0.00;${redTag}-${symbol}#,##0.00`;
      }

      if (!useSeparator) {
        formatCode = formatCode.replace(/#,##0/g, '#0');
      }

      this.applyFormatToRange(worksheet, startCell, endCell, {
        numFmt: formatCode,
      });

      this.logger.info(`Applied accounting format to ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to apply accounting format: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Apply VND currency format
   */
  async applyVNDCurrencyFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    decimalPlaces: number = 0
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      let formatCode = '';
      if (decimalPlaces === 0) {
        formatCode = '#,##0" đ"\\';
      } else {
        formatCode = `#,##0.${'0'.repeat(decimalPlaces)}" đ"\\`;
      }

      this.applyFormatToRange(worksheet, startCell, endCell, { numFmt: formatCode });

      this.logger.info(`Applied VND format to ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to apply VND format: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Apply negative numbers in red format
   */
  async applyNegativeRedFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    decimalPlaces: number = 2
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const formatCode = `#,##0.${'0'.repeat(decimalPlaces)};[Red]-#,##0.${'0'.repeat(decimalPlaces)}`;

      this.applyFormatToRange(worksheet, startCell, endCell, { numFmt: formatCode });

      this.logger.info(`Applied negative red format to ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to apply negative red format: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Show zeros instead of empty cells
   */
  async showZerosInsteadOfEmpty(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const range = parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: 'Invalid cell range' };
      }

      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          const cellValue = cell.value;
          if (cellValue === null || cellValue === undefined || cellValue === '') {
            cell.value = 0;
          }
        }
      }

      this.logger.info(`Set zeros for empty cells in ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to set zeros: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  // ============================================
  // Financial Analysis
  // ============================================

  /**
   * Calculate period comparison (previous vs current)
   */
  async calculatePeriodComparison(
    filename: string,
    worksheetName: string,
    currentValueRange: string,
    previousValueRange: string,
    outputRange: string,
    showPercentage: boolean = true
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const curRange = parseCellRange(currentValueRange);
      const prevRange = parseCellRange(previousValueRange);
      if (!curRange || !prevRange) {
        return { success: false, error: 'Invalid ranges or mismatched data' };
      }

      const currentData = this.getRangeData(worksheet, `${curRange.start.colLetter}${curRange.start.row}`, `${curRange.end.colLetter}${curRange.end.row}`);
      const previousData = this.getRangeData(worksheet, `${prevRange.start.colLetter}${prevRange.start.row}`, `${prevRange.end.colLetter}${prevRange.end.row}`);

      if (!currentData || !previousData || currentData.length !== previousData.length) {
        return { success: false, error: 'Invalid ranges or mismatched data' };
      }

      const currentColIndex = 0;
      const prevColIndex = 0;

      const outAddr = parseCellAddress(outputRange) ?? parseCellRange(outputRange)?.start;
      const outputCol = outAddr ? outAddr.colLetter : (outputRange.match(/^([A-Za-z]+)/)?.[1]?.toUpperCase() ?? 'A');
      const outputStartRow = outAddr ? outAddr.row : curRange.start.row;

      for (let i = 0; i < currentData.length; i++) {
        const current = typeof currentData[i][currentColIndex] === 'number' ? currentData[i][currentColIndex] as number : 0;
        const previous = typeof previousData[i][prevColIndex] === 'number' ? previousData[i][prevColIndex] as number : 0;

        const variance = current - previous;
        const percentage = previous !== 0 ? ((current - previous) / previous) * 100 : 0;

        const cell = worksheet.getCell(`${outputCol}${outputStartRow + i}`);
        if (showPercentage) {
          cell.value = percentage;
          cell.numFmt = '0.00"%"';
        } else {
          cell.value = variance;
          cell.numFmt = '#,##0.00';
        }
      }

      this.logger.info(`Calculated period comparison for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate period comparison: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate variance (budget vs actual)
   */
  async calculateVariance(
    filename: string,
    worksheetName: string,
    budgetRange: string,
    actualRange: string,
    outputRange: string,
    showPercentage: boolean = true
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const bRange = parseCellRange(budgetRange);
      const aRange = parseCellRange(actualRange);
      if (!bRange || !aRange) {
        return { success: false, error: 'Invalid ranges or mismatched data' };
      }

      const budgetData = this.getRangeData(worksheet, `${bRange.start.colLetter}${bRange.start.row}`, `${bRange.end.colLetter}${bRange.end.row}`);
      const actualData = this.getRangeData(worksheet, `${aRange.start.colLetter}${aRange.start.row}`, `${aRange.end.colLetter}${aRange.end.row}`);

      if (!budgetData || !actualData || budgetData.length !== actualData.length) {
        return { success: false, error: 'Invalid ranges or mismatched data' };
      }

      const budgetColIndex = 0;
      const actualColIndex = 0;

      const outAddr = parseCellAddress(outputRange) ?? parseCellRange(outputRange)?.start;
      const outputCol = outAddr ? outAddr.colLetter : (outputRange.match(/^([A-Za-z]+)/)?.[1]?.toUpperCase() ?? 'A');
      const outputStartRow = outAddr ? outAddr.row : bRange.start.row;

      for (let i = 0; i < budgetData.length; i++) {
        const budget = typeof budgetData[i][budgetColIndex] === 'number' ? budgetData[i][budgetColIndex] as number : 0;
        const actual = typeof actualData[i][actualColIndex] === 'number' ? actualData[i][actualColIndex] as number : 0;

        const variance = actual - budget;
        const percentage = budget !== 0 ? ((actual - budget) / budget) * 100 : 0;

        const cell = worksheet.getCell(`${outputCol}${outputStartRow + i}`);
        if (showPercentage) {
          cell.value = percentage;
          cell.numFmt = '0.00"%"';
        } else {
          cell.value = variance;
          cell.numFmt = '#,##0.00';
        }
      }

      this.logger.info(`Calculated variance for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate variance: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  // ============================================
  // Validation & Checks
  // ============================================

  /**
   * Check if debits equal credits
   */
  async checkBalance(
    filename: string,
    worksheetName: string,
    debitRange: string,
    creditRange: string
  ): Promise<OperationResult<{ isBalanced: boolean; debitTotal: number; creditTotal: number; difference: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const dRange = parseCellRange(debitRange);
      const cRange = parseCellRange(creditRange);
      if (!dRange || !cRange) {
        return { success: false, error: 'Invalid debit or credit range' };
      }

      const debitData = this.getRangeData(worksheet, `${dRange.start.colLetter}${dRange.start.row}`, `${dRange.end.colLetter}${dRange.end.row}`);
      const creditData = this.getRangeData(worksheet, `${cRange.start.colLetter}${cRange.start.row}`, `${cRange.end.colLetter}${cRange.end.row}`);

      if (!debitData || !creditData) {
        return { success: false, error: 'Invalid ranges' };
      }

      const debitColIndex = this.getColumnIndex(debitRange.match(/^([A-Z]+\d+)/)![1], debitRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const creditColIndex = this.getColumnIndex(creditRange.match(/^([A-Z]+\d+)/)![1], creditRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);

      let debitTotal = 0;
      let creditTotal = 0;

      for (const row of debitData) {
        if (typeof row[debitColIndex] === 'number') {
          debitTotal += row[debitColIndex] as number;
        }
      }

      for (const row of creditData) {
        if (typeof row[creditColIndex] === 'number') {
          creditTotal += row[creditColIndex] as number;
        }
      }

      const isBalanced = Math.abs(debitTotal - creditTotal) < 0.01;
      const difference = debitTotal - creditTotal;

      this.logger.info(`Balance check for ${filename}!${worksheetName}: ${isBalanced ? 'Balanced' : 'Not Balanced'}`);
      return { success: true, data: { isBalanced, debitTotal, creditTotal, difference } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to check balance: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Find anomalous values (outliers)
   */
  async findAnomalies(
    filename: string,
    worksheetName: string,
    rangeStart: string,
    rangeEnd: string,
    stdDevThreshold: number = 2
  ): Promise<OperationResult<{ anomalies: Array<{ cell: string; value: number; deviation: number }> }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const data = this.getRangeData(worksheet, rangeStart, rangeEnd);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const colIndex = this.getColumnIndex(rangeStart, rangeStart.match(/^([A-Z]+)/)![1]);
      const values: number[] = [];

      for (const row of data) {
        if (typeof row[colIndex] === 'number') {
          values.push(row[colIndex] as number);
        }
      }

      if (values.length < 2) {
        return { success: true, data: { anomalies: [] } };
      }

      // Calculate mean and standard deviation
      const mean = values.reduce((sum, v) => sum + v, 0) / values.length;
      const variance = values.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / values.length;
      const stdDev = Math.sqrt(variance);

      // Find anomalies
      const anomalies: Array<{ cell: string; value: number; deviation: number }> = [];
      const startCol = rangeStart.match(/^([A-Z]+)/)![1];
      const startRow = parseInt(rangeStart.match(/\d+/)![0], 10);

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const value = typeof row[colIndex] === 'number' ? row[colIndex] as number : 0;
        const zScore = stdDev !== 0 ? Math.abs((value - mean) / stdDev) : 0;

        if (zScore > stdDevThreshold) {
          anomalies.push({
            cell: `${startCol}${startRow + i}`,
            value,
            deviation: zScore
          });
        }
      }

      this.logger.info(`Found ${anomalies.length} anomalies in ${filename}!${worksheetName}`);
      return { success: true, data: { anomalies } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to find anomalies: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  // ============================================
  // Helper Methods
  // ============================================

  private getRangeData(worksheet: ExcelJS.Worksheet, startCell: string, endCell?: string): (string | number | boolean | Date | null)[][] {
    const range = parseCellRange(startCell, endCell);
    if (!range) return [];

    const data: (string | number | boolean | Date | null)[][] = [];

    for (let row = range.start.row; row <= range.end.row; row++) {
      const rowData: (string | number | boolean | Date | null)[] = [];
      for (let col = range.start.column; col <= range.end.column; col++) {
        const cell = worksheet.getCell(row, col);
        const cellValue = cell.value;
        if (cellValue !== undefined) {
          rowData.push(cellValue as string | number | boolean | Date | null);
        } else {
          rowData.push(null);
        }
      }
      data.push(rowData);
    }

    return data;
  }

  private getColumnIndex(rangeStart: string, columnLetter: string): number {
    const addr = parseCellAddress(rangeStart);
    const startCol = addr ? addr.column : this.columnToNumber(rangeStart.match(/^([A-Za-z]+)/)?.[1] ?? 'A');
    return this.columnToNumber(columnLetter) - startCol;
  }

  private columnToNumber(column: string): number {
    return columnLetterToNumber(column);
  }

  private applyFormatToRange(
    worksheet: ExcelJS.Worksheet,
    startCell: string,
    endCell: string,
    format: Partial<ExcelJS.Style>
  ): void {
    const range = parseCellRange(startCell, endCell);
    if (!range) return;

    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.column; col <= range.end.column; col++) {
        const cell = worksheet.getCell(row, col);
        Object.assign(cell, format);
      }
    }
  }
}