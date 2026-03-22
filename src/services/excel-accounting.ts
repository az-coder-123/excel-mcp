/**
 * Excel Accounting Service
 * Single Responsibility: Accounting-specific operations for financial data
 */

import ExcelJS from 'exceljs';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult } from '../types/index.js';

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
        const targetColIndex = this.getColumnIndex(rangeStart, rangeStart.match(/^([A-Z]+)/)![1]);

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          if (row[criteriaColIndex] === criteriaValue && typeof row[targetColIndex] === 'number') {
            sum += row[targetColIndex] as number;
          }
        }
      } else {
        // Simple sum
        const targetColIndex = this.getColumnIndex(rangeStart, rangeStart.match(/^([A-Z]+)/)![1]);
        for (const row of data) {
          if (typeof row[targetColIndex] === 'number') {
            sum += row[targetColIndex] as number;
          }
        }
      }

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
      const targetColIndex = this.getColumnIndex(rangeStart, rangeStart.match(/^([A-Z]+)/)![1]);

      for (const row of data) {
        if (typeof row[targetColIndex] === 'number') {
          sum += row[targetColIndex] as number;
          count++;
        }
      }

      const average = count > 0 ? sum / count : 0;

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
      const outputColLetter = outputStartCell.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputStartCell.match(/\d+/)![0], 10);

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
      const outputColLetter = outputStartCell.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputStartCell.match(/\d+/)![0], 10);

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

      const data = this.getRangeData(worksheet, rangeStart, rangeEnd);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const dateColIndex = this.getColumnIndex(rangeStart, dateColumn);
      const valueColIndex = this.getColumnIndex(rangeStart, valueColumn);

      let ytdTotal = 0;
      let currentYear = 0;

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const dateValue = row[dateColIndex];
        const value = typeof row[valueColIndex] === 'number' ? row[valueColIndex] as number : 0;

        // Extract year from date
        let year = 0;
        if (dateValue instanceof Date) {
          year = dateValue.getFullYear();
        } else if (typeof dateValue === 'number' && dateValue > 0) {
          // Excel serial date (days since 1/1/1900)
          const excelEpoch = new Date(1899, 11, 30);
          const date = new Date(excelEpoch.getTime() + dateValue * 86400000);
          year = date.getFullYear();
        }

        // Reset YTD on year change
        if (year !== 0 && year !== currentYear) {
          ytdTotal = 0;
          currentYear = year;
        }

        ytdTotal += value;

        const cell = worksheet.getCell(`${outputColumn}${i + 1}`);
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

      let formatCode = '';

      if (showParentheses) {
        // Format: (1,234.56) for negative
        formatCode = `_(${symbol}* #,##0.00_);_(${symbol} \\(#,##0.00\\);_(${symbol}*"-"_);_(@_)`;
      } else {
        formatCode = `${symbol}#,##0.00;-${symbol}#,##0.00`;
      }

      if (!useSeparator) {
        formatCode = formatCode.replace(/#,##0/g, '#0');
      }

      this.applyFormatToRange(worksheet, startCell, endCell, {
        numFmt: formatCode,
        font: showRed ? { color: { argb: 'FF000000' } } : undefined
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

      const startRow = parseInt(startCell.match(/\d+/)![0], 10);
      const endRow = parseInt(endCell.match(/\d+/)![0], 10);
      const startCol = this.columnToNumber(startCell.match(/^([A-Z]+)/)![1]);
      const endCol = this.columnToNumber(endCell.match(/^([A-Z]+)/)![1]);

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
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

      const currentData = this.getRangeData(worksheet, currentValueRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], currentValueRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      const previousData = this.getRangeData(worksheet, previousValueRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], previousValueRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);

      if (!currentData || !previousData || currentData.length !== previousData.length) {
        return { success: false, error: 'Invalid ranges or mismatched data' };
      }

      const currentColIndex = this.getColumnIndex(currentValueRange.match(/^([A-Z]+\d+)/)![1], currentValueRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const prevColIndex = this.getColumnIndex(previousValueRange.match(/^([A-Z]+\d+)/)![1], previousValueRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);

      const outputCol = outputRange.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputRange.match(/\d+/)![0], 10);

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

      const budgetData = this.getRangeData(worksheet, budgetRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], budgetRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      const actualData = this.getRangeData(worksheet, actualRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], actualRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);

      if (!budgetData || !actualData || budgetData.length !== actualData.length) {
        return { success: false, error: 'Invalid ranges or mismatched data' };
      }

      const budgetColIndex = this.getColumnIndex(budgetRange.match(/^([A-Z]+\d+)/)![1], budgetRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const actualColIndex = this.getColumnIndex(actualRange.match(/^([A-Z]+\d+)/)![1], actualRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);

      const outputCol = outputRange.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputRange.match(/\d+/)![0], 10);

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

      const debitData = this.getRangeData(worksheet, debitRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], debitRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      const creditData = this.getRangeData(worksheet, creditRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], creditRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);

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

  private getRangeData(worksheet: ExcelJS.Worksheet, startCell: string, endCell: string): (string | number | boolean | Date | null)[][] {
    const startRow = parseInt(startCell.match(/\d+/)![0], 10);
    const endRow = parseInt(endCell.match(/\d+/)![0], 10);
    const startCol = this.columnToNumber(startCell.match(/^([A-Z]+)/)![1]);
    const endCol = this.columnToNumber(endCell.match(/^([A-Z]+)/)![1]);

    const data: (string | number | boolean | Date | null)[][] = [];

    for (let row = startRow; row <= endRow; row++) {
      const rowData: (string | number | boolean | Date | null)[] = [];
      for (let col = startCol; col <= endCol; col++) {
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
    return this.columnToNumber(columnLetter) - this.columnToNumber(rangeStart.match(/^([A-Z]+)/)![1]);
  }

  private columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  }

  private applyFormatToRange(
    worksheet: ExcelJS.Worksheet,
    startCell: string,
    endCell: string,
    format: Partial<ExcelJS.Style>
  ): void {
    const startRow = parseInt(startCell.match(/\d+/)![0], 10);
    const endRow = parseInt(endCell.match(/\d+/)![0], 10);
    const startCol = this.columnToNumber(startCell.match(/^([A-Z]+)/)![1]);
    const endCol = this.columnToNumber(endCell.match(/^([A-Z]+)/)![1]);

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = worksheet.getCell(row, col);
        Object.assign(cell, format);
      }
    }
  }
}