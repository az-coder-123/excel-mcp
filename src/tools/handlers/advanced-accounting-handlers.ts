/**
 * Advanced Accounting Tool Handlers
 * Handlers for advanced accounting operations
 */

import { ExcelAdvancedAccounting } from '../../services/excel-advanced-accounting.js';
import { OperationResult } from '../../types/index.js';

type ToolHandler = (args: Record<string, unknown>) => Promise<OperationResult>;
type RatioType = 'current' | 'quick' | 'debt-to-equity' | 'return-on-equity' | 'profit-margin';

/** Typed argument extraction — replaces the former `args: any` direct access (audit Finding 2.2). */
const strArg = (args: Record<string, unknown>, key: string): string => {
  const value = args[key];
  if (typeof value !== 'string') {
    throw new Error(`Missing or invalid required parameter: ${key}`);
  }
  return value;
};

const reqNum = (args: Record<string, unknown>, key: string): number => {
  const value = args[key];
  if (typeof value !== 'number') {
    throw new Error(`Missing or invalid required parameter: ${key}`);
  }
  return value;
};

const optNum = (args: Record<string, unknown>, key: string): number | undefined =>
  typeof args[key] === 'number' ? (args[key] as number) : undefined;

export class AdvancedAccountingHandlers {
  private advancedAccounting: ExcelAdvancedAccounting;

  constructor(advancedAccounting: ExcelAdvancedAccounting) {
    this.advancedAccounting = advancedAccounting;
  }

  /**
   * Calculate NPV (Net Present Value)
   */
  calculateNPV: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.advancedAccounting.calculateNPV(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      reqNum(args, 'rate'),
      strArg(args, 'valuesRange')
    );
    return result;
  };

  /**
   * Calculate IRR (Internal Rate of Return)
   */
  calculateIRR: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.advancedAccounting.calculateIRR(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'valuesRange'),
      optNum(args, 'guess')
    );
    return result;
  };

  /**
   * Calculate Financial Ratios
   */
  calculateFinancialRatio: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.advancedAccounting.calculateFinancialRatios(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'ratioType') as RatioType,
      strArg(args, 'numeratorRange'),
      strArg(args, 'denominatorRange')
    );
    return result;
  };

  /**
   * Create Amortization Schedule
   */
  createAmortizationSchedule: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.advancedAccounting.createAmortizationSchedule(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'startCell'),
      reqNum(args, 'principal'),
      reqNum(args, 'annualRate'),
      reqNum(args, 'numberOfPeriods')
    );
    return result;
  };

  /**
   * Create Aging Report
   */
  createAgingReport: ToolHandler = async (args: Record<string, unknown>) => {
    const asOf = strArg(args, 'asOfDate');
    const date = new Date(asOf);
    if (isNaN(date.getTime())) {
      return {
        success: false,
        error: 'Invalid date format. Use ISO format: YYYY-MM-DD'
      };
    }

    const result = await this.advancedAccounting.createAgingReport(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'invoiceDateColumn'),
      strArg(args, 'amountColumn'),
      date,
      strArg(args, 'outputStartCell')
    );
    return result;
  };

  /**
   * Calculate Tax
   */
  calculateTax: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.advancedAccounting.calculateTax(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'amountRange'),
      reqNum(args, 'taxRate'),
      strArg(args, 'outputRange')
    );
    return result;
  };

  /**
   * Convert Currency
   */
  convertCurrency: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.advancedAccounting.convertCurrency(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'amountRange'),
      reqNum(args, 'exchangeRate'),
      strArg(args, 'outputRange')
    );
    return result;
  };
}
