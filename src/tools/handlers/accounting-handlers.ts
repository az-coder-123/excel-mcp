/**
 * Accounting Tool Handlers
 * Single Responsibility: Handle accounting-specific operations
 */

import { ExcelAccounting } from '../../services/excel-accounting.js';
import { OperationResult } from '../../types/index.js';

type ToolHandler = (args: Record<string, unknown>) => Promise<OperationResult>;

export class AccountingHandlers {
  private accountingService: ExcelAccounting;

  constructor(accountingService: ExcelAccounting) {
    this.accountingService = accountingService;
  }

  // Financial Calculations
  financialSum: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculateSum(
      args.filename,
      args.worksheet,
      args.rangeStart,
      args.rangeEnd,
      args.criteriaColumn,
      args.criteriaValue
    );
    return result;
  };

  financialAverage: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculateAverage(
      args.filename,
      args.worksheet,
      args.rangeStart,
      args.rangeEnd
    );
    return result;
  };

  runningTotal: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculateRunningTotal(
      args.filename,
      args.worksheet,
      args.valueStartCell,
      args.valueEndCell,
      args.outputStartCell
    );
    return result;
  };

  percentageOfTotal: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculatePercentageOfTotal(
      args.filename,
      args.worksheet,
      args.valueStartCell,
      args.valueEndCell,
      args.outputStartCell
    );
    return result;
  };

  yearToDate: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculateYTD(
      args.filename,
      args.worksheet,
      args.dateColumn,
      args.valueColumn,
      args.rangeStart,
      args.rangeEnd,
      args.outputColumn
    );
    return result;
  };

  // Accounting Formats
  accountingFormat: ToolHandler = async (args: any) => {
    const result = await this.accountingService.applyAccountingFormat(
      args.filename,
      args.worksheet,
      args.startCell,
      args.endCell,
      {
        decimalPlaces: args.decimalPlaces,
        useSeparator: args.useSeparator,
        showNegativeInRed: args.showNegativeInRed,
        showNegativeInParentheses: args.showNegativeInParentheses,
        currencySymbol: args.currencySymbol,
      }
    );
    return result;
  };

  vndCurrencyFormat: ToolHandler = async (args: any) => {
    const result = await this.accountingService.applyVNDCurrencyFormat(
      args.filename,
      args.worksheet,
      args.startCell,
      args.endCell,
      args.decimalPlaces
    );
    return result;
  };

  negativeRedFormat: ToolHandler = async (args: any) => {
    const result = await this.accountingService.applyNegativeRedFormat(
      args.filename,
      args.worksheet,
      args.startCell,
      args.endCell,
      args.decimalPlaces
    );
    return result;
  };

  showZerosInsteadOfEmpty: ToolHandler = async (args: any) => {
    const result = await this.accountingService.showZerosInsteadOfEmpty(
      args.filename,
      args.worksheet,
      args.startCell,
      args.endCell
    );
    return result;
  };

  // Financial Analysis
  periodComparison: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculatePeriodComparison(
      args.filename,
      args.worksheet,
      args.currentValueRange,
      args.previousValueRange,
      args.outputRange,
      args.showPercentage
    );
    return result;
  };

  varianceAnalysis: ToolHandler = async (args: any) => {
    const result = await this.accountingService.calculateVariance(
      args.filename,
      args.worksheet,
      args.budgetRange,
      args.actualRange,
      args.outputRange,
      args.showPercentage
    );
    return result;
  };

  // Validation & Checks
  checkBalance: ToolHandler = async (args: any) => {
    const result = await this.accountingService.checkBalance(
      args.filename,
      args.worksheet,
      args.debitRange,
      args.creditRange
    );
    return result;
  };

  findAnomalies: ToolHandler = async (args: any) => {
    const result = await this.accountingService.findAnomalies(
      args.filename,
      args.worksheet,
      args.rangeStart,
      args.rangeEnd,
      args.stdDevThreshold
    );
    return result;
  };
}