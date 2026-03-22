/**
 * Advanced Accounting Tool Handlers
 * Handlers for advanced accounting operations
 */

import { ExcelAdvancedAccounting } from '../../services/excel-advanced-accounting.js';
import { OperationResult } from '../../types/index.js';

type ToolHandler = (args: Record<string, unknown>) => Promise<OperationResult>;

export class AdvancedAccountingHandlers {
  private advancedAccounting: ExcelAdvancedAccounting;

  constructor(advancedAccounting: ExcelAdvancedAccounting) {
    this.advancedAccounting = advancedAccounting;
  }

  /**
   * Calculate NPV (Net Present Value)
   */
  calculateNPV: ToolHandler = async (args: any) => {
    const result = await this.advancedAccounting.calculateNPV(
      args.filename,
      args.worksheet,
      args.rate,
      args.valuesRange
    );
    return result;
  };

  /**
   * Calculate IRR (Internal Rate of Return)
   */
  calculateIRR: ToolHandler = async (args: any) => {
    const result = await this.advancedAccounting.calculateIRR(
      args.filename,
      args.worksheet,
      args.valuesRange,
      args.guess
    );
    return result;
  };

  /**
   * Calculate Financial Ratios
   */
  calculateFinancialRatio: ToolHandler = async (args: any) => {
    const result = await this.advancedAccounting.calculateFinancialRatios(
      args.filename,
      args.worksheet,
      args.ratioType,
      args.numeratorRange,
      args.denominatorRange
    );
    return result;
  };

  /**
   * Create Amortization Schedule
   */
  createAmortizationSchedule: ToolHandler = async (args: any) => {
    const result = await this.advancedAccounting.createAmortizationSchedule(
      args.filename,
      args.worksheet,
      args.startCell,
      args.principal,
      args.annualRate,
      args.numberOfPeriods
    );
    return result;
  };

  /**
   * Create Aging Report
   */
  createAgingReport: ToolHandler = async (args: any) => {
    const date = new Date(args.asOfDate);
    if (isNaN(date.getTime())) {
      return {
        success: false,
        error: 'Invalid date format. Use ISO format: YYYY-MM-DD'
      };
    }

    const result = await this.advancedAccounting.createAgingReport(
      args.filename,
      args.worksheet,
      args.invoiceDateColumn,
      args.amountColumn,
      date,
      args.outputStartCell
    );
    return result;
  };

  /**
   * Calculate Tax
   */
  calculateTax: ToolHandler = async (args: any) => {
    const result = await this.advancedAccounting.calculateTax(
      args.filename,
      args.worksheet,
      args.amountRange,
      args.taxRate,
      args.outputRange
    );
    return result;
  };

  /**
   * Convert Currency
   */
  convertCurrency: ToolHandler = async (args: any) => {
    const result = await this.advancedAccounting.convertCurrency(
      args.filename,
      args.worksheet,
      args.amountRange,
      args.exchangeRate,
      args.outputRange
    );
    return result;
  };
}
