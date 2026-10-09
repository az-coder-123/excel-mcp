/**
 * Excel Advanced Accounting Service
 * Single Responsibility: Advanced accounting-specific operations for financial data
 */

import ExcelJS from 'exceljs';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult } from '../types/index.js';

export class ExcelAdvancedAccounting {
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
  // Advanced Financial Calculations
  // ============================================

  /**
   * Calculate NPV (Net Present Value)
   */
  async calculateNPV(
    filename: string,
    worksheetName: string,
    rate: number,
    valuesRange: string
  ): Promise<OperationResult<{ npv: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const data = this.getRangeData(worksheet, valuesRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], valuesRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const colIndex = this.getColumnIndex(valuesRange.match(/^([A-Z]+\d+)/)![1], valuesRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const values: number[] = [];

      for (const row of data) {
        if (typeof row[colIndex] === 'number') {
          values.push(row[colIndex] as number);
        }
      }

      // Calculate NPV
      let npv = 0;
      for (let i = 0; i < values.length; i++) {
        npv += values[i] / Math.pow(1 + rate, i + 1);
      }

      this.logger.info(`Calculated NPV for ${filename}!${worksheetName}: ${npv}`);
      return { success: true, data: { npv } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate NPV: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate IRR (Internal Rate of Return)
   */
  async calculateIRR(
    filename: string,
    worksheetName: string,
    valuesRange: string,
    guess: number = 0.1
  ): Promise<OperationResult<{ irr: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const data = this.getRangeData(worksheet, valuesRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], valuesRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const colIndex = this.getColumnIndex(valuesRange.match(/^([A-Z]+\d+)/)![1], valuesRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const values: number[] = [];

      for (const row of data) {
        if (typeof row[colIndex] === 'number') {
          values.push(row[colIndex] as number);
        }
      }

      // Validate cash flows before solving (audit Finding 4.1)
      if (values.length < 2) {
        return { success: false, error: 'IRR requires at least two cash flow values' };
      }
      const hasNegative = values.some((v) => v < 0);
      const hasPositive = values.some((v) => v > 0);
      if (!hasNegative || !hasPositive) {
        return {
          success: false,
          error: 'IRR requires at least one negative and one positive cash flow',
        };
      }

      // Standard IRR with Excel semantics: solve NPV(r) = Σ_{t=0}^{n-1} v_t / (1+r)^t = 0
      // via Newton-Raphson using the mathematically correct derivative:
      //   NPV'(r) = Σ_{t=0}^{n-1} -t · v_t / (1+r)^(t+1)
      let rate = guess;
      let converged = false;
      const maxIterations = 100;
      const tolerance = 1e-10;

      for (let i = 0; i < maxIterations; i++) {
        let npv = 0;
        let dnpv = 0;

        for (let t = 0; t < values.length; t++) {
          npv += values[t] / Math.pow(1 + rate, t);
          dnpv -= (t * values[t]) / Math.pow(1 + rate, t + 1);
        }

        if (Math.abs(npv) < tolerance) {
          converged = true;
          break;
        }

        // Guard against a flat derivative (division by zero → NaN/Infinity)
        if (!Number.isFinite(dnpv) || Math.abs(dnpv) < 1e-12) {
          break;
        }

        const nextRate = rate - npv / dnpv;
        if (!Number.isFinite(nextRate)) {
          break;
        }
        // Keep the rate economically valid (r > -100%)
        rate = Math.max(nextRate, -0.999999);
      }

      if (!converged) {
        this.logger.warn(`IRR calculation did not converge for ${filename}!${worksheetName}`);
        return {
          success: false,
          error: 'IRR calculation did not converge; try a different guess value',
        };
      }

      const irr = rate * 100; // Convert to percentage

      this.logger.info(`Calculated IRR for ${filename}!${worksheetName}: ${irr}%`);
      return { success: true, data: { irr } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate IRR: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate financial ratios
   */
  async calculateFinancialRatios(
    filename: string,
    worksheetName: string,
    ratioType: 'current' | 'quick' | 'debt-to-equity' | 'return-on-equity' | 'profit-margin',
    numeratorRange: string,
    denominatorRange: string
  ): Promise<OperationResult<{ ratio: number; name: string }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const numeratorData = this.getRangeData(worksheet, numeratorRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], numeratorRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      const denominatorData = this.getRangeData(worksheet, denominatorRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], denominatorRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);

      if (!numeratorData || !denominatorData) {
        return { success: false, error: 'Invalid ranges' };
      }

      const numColIndex = this.getColumnIndex(numeratorRange.match(/^([A-Z]+\d+)/)![1], numeratorRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const denColIndex = this.getColumnIndex(denominatorRange.match(/^([A-Z]+\d+)/)![1], denominatorRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);

      let numerator = 0;
      let denominator = 0;

      for (const row of numeratorData) {
        if (typeof row[numColIndex] === 'number') {
          numerator += row[numColIndex] as number;
        }
      }

      for (const row of denominatorData) {
        if (typeof row[denColIndex] === 'number') {
          denominator += row[denColIndex] as number;
        }
      }

      if (denominator === 0) {
        return { success: false, error: 'Denominator cannot be zero' };
      }

      const ratio = numerator / denominator;
      const ratioNames = {
        'current': 'Current Ratio',
        'quick': 'Quick Ratio',
        'debt-to-equity': 'Debt-to-Equity Ratio',
        'return-on-equity': 'Return on Equity',
        'profit-margin': 'Profit Margin'
      };

      this.logger.info(`Calculated ${ratioNames[ratioType]} for ${filename}!${worksheetName}: ${ratio}`);
      return { success: true, data: { ratio, name: ratioNames[ratioType] } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate financial ratio: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Create amortization schedule
   */
  async createAmortizationSchedule(
    filename: string,
    worksheetName: string,
    startCell: string,
    principal: number,
    annualRate: number,
    numberOfPeriods: number
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

      const startCol = startCell.match(/^([A-Z]+)/)![1];
      const startRow = parseInt(startCell.match(/\d+/)![0], 10);

      // Calculate monthly payment
      const monthlyRate = annualRate / 12;
      const payment = principal * (monthlyRate * Math.pow(1 + monthlyRate, numberOfPeriods)) / (Math.pow(1 + monthlyRate, numberOfPeriods) - 1);

      let balance = principal;

      // Write headers
      worksheet.getCell(`${startCol}${startRow}`).value = 'Period';
      worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 1)}${startRow}`).value = 'Payment';
      worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 2)}${startRow}`).value = 'Principal';
      worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 3)}${startRow}`).value = 'Interest';
      worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 4)}${startRow}`).value = 'Balance';

      // Calculate schedule
      for (let i = 1; i <= numberOfPeriods; i++) {
        const interest = balance * monthlyRate;
        const principalPaid = payment - interest;
        balance -= principalPaid;

        const row = startRow + i;

        worksheet.getCell(`${startCol}${row}`).value = i;
        worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 1)}${row}`).value = payment;
        worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 2)}${row}`).value = principalPaid;
        worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 3)}${row}`).value = interest;
        worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + 4)}${row}`).value = Math.max(0, balance);

        // Format as currency
        for (let j = 1; j <= 4; j++) {
          const cell = worksheet.getCell(`${this.numberToColumn(this.columnToNumber(startCol) + j)}${row}`);
          cell.numFmt = '#,##0.00';
        }
      }

      this.logger.info(`Created amortization schedule for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create amortization schedule: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Create aging report for accounts receivable/payable
   */
  async createAgingReport(
    filename: string,
    worksheetName: string,
    invoiceDateColumn: string,
    amountColumn: string,
    asOfDate: Date,
    outputStartCell: string
  ): Promise<OperationResult<{ summary: { [key: string]: number } }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      const lastRow = worksheet.rowCount;
      const data = this.getRangeData(worksheet, `${invoiceDateColumn}1`, `${amountColumn}${lastRow}`);

      if (!data || data.length === 0) {
        return { success: false, error: 'No data found' };
      }

      // Fix: compute both indexes relative to the SAME range start so each
      // column maps to its own slot. The previous code passed each column as
      // its own range start, making getColumnIndex always return 0 and causing
      // amounts to be read from the DATE column (caught by aging tests).
      const rangeStartCell = `${invoiceDateColumn}1`;
      const dateColIndex = this.getColumnIndex(rangeStartCell, invoiceDateColumn);
      const amountColIndex = this.getColumnIndex(rangeStartCell, amountColumn);

      const agingBuckets: { [key: string]: number } = {
        '0-30': 0,
        '31-60': 0,
        '61-90': 0,
        '91-120': 0,
        '120+': 0
      };

      for (const row of data) {
        const dateValue = row[dateColIndex];
        const amount = typeof row[amountColIndex] === 'number' ? row[amountColIndex] as number : 0;

        // Support Date objects, raw Excel serial numbers, and date strings (audit Finding 4.2)
        const invoiceDate = this.toDate(dateValue);
        if (!invoiceDate) {
          continue; // Skip headers, blanks, and values that are not dates
        }

        const daysPastDue = Math.floor((asOfDate.getTime() - invoiceDate.getTime()) / (1000 * 60 * 60 * 24));

        if (daysPastDue <= 30) {
          agingBuckets['0-30'] += amount;
        } else if (daysPastDue <= 60) {
          agingBuckets['31-60'] += amount;
        } else if (daysPastDue <= 90) {
          agingBuckets['61-90'] += amount;
        } else if (daysPastDue <= 120) {
          agingBuckets['91-120'] += amount;
        } else {
          agingBuckets['120+'] += amount;
        }
      }

      // Write aging report
      const outputCol = outputStartCell.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputStartCell.match(/\d+/)![0], 10);

      const agingHeaders = Object.keys(agingBuckets);
      for (let i = 0; i < agingHeaders.length; i++) {
        const cell = worksheet.getCell(`${this.numberToColumn(this.columnToNumber(outputCol) + i)}${outputStartRow}`);
        cell.value = agingHeaders[i];
        cell.font = { bold: true };

        const valueCell = worksheet.getCell(`${this.numberToColumn(this.columnToNumber(outputCol) + i)}${outputStartRow + 1}`);
        valueCell.value = agingBuckets[agingHeaders[i]];
        valueCell.numFmt = '#,##0.00';
      }

      this.logger.info(`Created aging report for ${filename}!${worksheetName}`);
      return { success: true, data: { summary: agingBuckets } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to create aging report: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Calculate tax
   */
  async calculateTax(
    filename: string,
    worksheetName: string,
    amountRange: string,
    taxRate: number,
    outputRange: string
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

      const data = this.getRangeData(worksheet, amountRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], amountRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const colIndex = this.getColumnIndex(amountRange.match(/^([A-Z]+\d+)/)![1], amountRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const outputCol = outputRange.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputRange.match(/\d+/)![0], 10);

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const amount = typeof row[colIndex] === 'number' ? row[colIndex] as number : 0;
        const tax = amount * (taxRate / 100);

        const cell = worksheet.getCell(`${outputCol}${outputStartRow + i}`);
        cell.value = tax;
        cell.numFmt = '#,##0.00';
      }

      this.logger.info(`Calculated tax at ${taxRate}% for ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to calculate tax: ${errorMessage}`);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Currency conversion
   */
  async convertCurrency(
    filename: string,
    worksheetName: string,
    amountRange: string,
    exchangeRate: number,
    outputRange: string
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

      const data = this.getRangeData(worksheet, amountRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![1], amountRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/)![2]);
      if (!data || data.length === 0) {
        return { success: false, error: 'No data found in range' };
      }

      const colIndex = this.getColumnIndex(amountRange.match(/^([A-Z]+\d+)/)![1], amountRange.match(/^([A-Z]+\d+)/)![1].match(/^([A-Z]+)/)![1]);
      const outputCol = outputRange.match(/^([A-Z]+)/)![1];
      const outputStartRow = parseInt(outputRange.match(/\d+/)![0], 10);

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const amount = typeof row[colIndex] === 'number' ? row[colIndex] as number : 0;
        const converted = amount * exchangeRate;

        const cell = worksheet.getCell(`${outputCol}${outputStartRow + i}`);
        cell.value = converted;
        cell.numFmt = '#,##0.00';
      }

      this.logger.info(`Converted currency for ${filename}!${worksheetName} at rate ${exchangeRate}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to convert currency: ${errorMessage}`);
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

  /**
   * Convert a cell value to a Date, supporting JavaScript Date objects,
   * raw Excel serial numbers (days since 1899-12-30), and parseable date strings.
   * Returns null when the value cannot be interpreted as a date.
   *
   * Serial numbers below 61 are rejected: Excel's Lotus-1900 compatibility bug
   * (phantom Feb 29, 1900) makes serials 1–60 ambiguous, and real-world invoice
   * dates always postdate 1900-03-01 anyway.
   */
  private toDate(value: string | number | boolean | Date | null): Date | null {
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? null : value;
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      // Excel serial dates: valid range 61 (1900-03-01) .. 2958465 (9999-12-31)
      if (value < 61 || value > 2958465) {
        return null;
      }
      const MS_PER_DAY = 86_400_000;
      const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
      return new Date(EXCEL_EPOCH_MS + Math.round(value * MS_PER_DAY));
    }
    if (typeof value === 'string' && value.trim() !== '') {
      const parsed = new Date(value);
      return Number.isNaN(parsed.getTime()) ? null : parsed;
    }
    return null;
  }

  private columnToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return result;
  }

  private numberToColumn(num: number): string {
    let column = '';
    while (num > 0) {
      num--;
      column = String.fromCharCode(65 + (num % 26)) + column;
      num = Math.floor(num / 26);
    }
    return column;
  }
}