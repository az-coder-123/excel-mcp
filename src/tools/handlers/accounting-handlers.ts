/**
 * Accounting Tool Handlers
 * Single Responsibility: Handle accounting-specific operations
 */

import { ExcelAccounting } from '../../services/excel-accounting.js';
import { OperationResult } from '../../types/index.js';

type ToolHandler = (args: Record<string, unknown>) => Promise<OperationResult>;

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

const optBool = (args: Record<string, unknown>, key: string): boolean | undefined =>
  typeof args[key] === 'boolean' ? (args[key] as boolean) : undefined;

const optNum = (args: Record<string, unknown>, key: string): number | undefined =>
  typeof args[key] === 'number' ? (args[key] as number) : undefined;

const optStr = (args: Record<string, unknown>, key: string): string | undefined =>
  typeof args[key] === 'string' ? (args[key] as string) : undefined;

export class AccountingHandlers {
  private accountingService: ExcelAccounting;

  constructor(accountingService: ExcelAccounting) {
    this.accountingService = accountingService;
  }

  // Financial Calculations
  financialSum: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculateSum(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'rangeStart'),
      strArg(args, 'rangeEnd'),
      optStr(args, 'criteriaColumn'),
      optStr(args, 'criteriaValue')
    );
    return result;
  };

  financialAverage: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculateAverage(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'rangeStart'),
      strArg(args, 'rangeEnd')
    );
    return result;
  };

  runningTotal: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculateRunningTotal(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'valueStartCell'),
      strArg(args, 'valueEndCell'),
      strArg(args, 'outputStartCell')
    );
    return result;
  };

  percentageOfTotal: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculatePercentageOfTotal(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'valueStartCell'),
      strArg(args, 'valueEndCell'),
      strArg(args, 'outputStartCell')
    );
    return result;
  };

  yearToDate: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculateYTD(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'dateColumn'),
      strArg(args, 'valueColumn'),
      strArg(args, 'rangeStart'),
      strArg(args, 'rangeEnd'),
      strArg(args, 'outputColumn')
    );
    return result;
  };

  // Accounting Formats
  accountingFormat: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.applyAccountingFormat(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'startCell'),
      strArg(args, 'endCell'),
      {
        decimalPlaces: optNum(args, 'decimalPlaces'),
        useSeparator: optBool(args, 'useSeparator'),
        showNegativeInRed: optBool(args, 'showNegativeInRed'),
        showNegativeInParentheses: optBool(args, 'showNegativeInParentheses'),
        currencySymbol: optStr(args, 'currencySymbol'),
      }
    );
    return result;
  };

  vndCurrencyFormat: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.applyVNDCurrencyFormat(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'startCell'),
      strArg(args, 'endCell'),
      optNum(args, 'decimalPlaces')
    );
    return result;
  };

  negativeRedFormat: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.applyNegativeRedFormat(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'startCell'),
      strArg(args, 'endCell'),
      optNum(args, 'decimalPlaces')
    );
    return result;
  };

  showZerosInsteadOfEmpty: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.showZerosInsteadOfEmpty(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'startCell'),
      strArg(args, 'endCell')
    );
    return result;
  };

  // Financial Analysis
  periodComparison: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculatePeriodComparison(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'currentValueRange'),
      strArg(args, 'previousValueRange'),
      strArg(args, 'outputRange'),
      optBool(args, 'showPercentage')
    );
    return result;
  };

  varianceAnalysis: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.calculateVariance(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'budgetRange'),
      strArg(args, 'actualRange'),
      strArg(args, 'outputRange'),
      optBool(args, 'showPercentage')
    );
    return result;
  };

  // Validation & Checks
  checkBalance: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.checkBalance(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'debitRange'),
      strArg(args, 'creditRange')
    );
    return result;
  };

  findAnomalies: ToolHandler = async (args: Record<string, unknown>) => {
    const result = await this.accountingService.findAnomalies(
      strArg(args, 'filename'),
      strArg(args, 'worksheet'),
      strArg(args, 'rangeStart'),
      strArg(args, 'rangeEnd'),
      reqNum(args, 'stdDevThreshold')
    );
    return result;
  };
}