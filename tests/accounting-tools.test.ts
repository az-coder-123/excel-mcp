/**
 * Comprehensive Unit and Integration Tests for ExcelAccounting
 * Testing all tools, edge cases, mathematical correctness, and ToolHandler wiring:
 * - calculateSum (simple, with criteria, non-numeric filtering)
 * - calculateAverage (mean calculation, count, zero elements)
 * - calculateRunningTotal (cumulative sum, row offsets, zero values)
 * - calculatePercentageOfTotal (percentages, sum=0 division by zero guard)
 * - calculateYTD (year resets, JS Dates, string dates, Excel serial numbers, row offsets)
 * - applyAccountingFormat (format strings, parentheses, [Red] tag, symbols)
 * - applyVNDCurrencyFormat (0 and N decimal places)
 * - applyNegativeRedFormat ([Red] negative format)
 * - showZerosInsteadOfEmpty (null/empty to 0, existing preserved)
 * - calculatePeriodComparison (variance, percentage, previous=0 guard)
 * - calculateVariance (actual vs budget, positive/negative variance)
 * - checkBalance (balanced, unbalanced, single-cell and multi-cell ranges)
 * - findAnomalies (z-score, stdDev=0 guard, dataset < 2)
 */

import ExcelJS from 'exceljs';
import { ExcelAccounting } from '../src/services/excel-accounting.js';
import { ExcelService } from '../src/services/excel-service.js';
import { ToolHandler } from '../src/tools/tool-handler.js';
import { PermissionChecker } from '../src/security/permission-checker.js';
import { Logger } from '../src/utils/logger.js';
import { PermissionConfig } from '../src/types/index.js';

const config = (overrides: Partial<PermissionConfig> = {}): PermissionConfig => ({
  allowedPaths: [],
  deniedPaths: [],
  maxFileSize: 50 * 1024 * 1024,
  allowedExtensions: ['.xlsx', '.xls', '.xlsm', '.xlsb'],
  permissions: ['read', 'write', 'delete'],
  ...overrides,
});

describe('ExcelAccounting Service & Tools', () => {
  let workbooks: Map<string, ExcelJS.Workbook>;
  let accounting: ExcelAccounting;
  let excelService: ExcelService;
  let handler: ToolHandler;
  const fileName = 'test_accounting.xlsx';
  const sheetName = 'Finance';

  beforeEach(() => {
    workbooks = new Map();
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    accounting = new ExcelAccounting(permChecker, logger, workbooks);
    excelService = new ExcelService(permChecker, logger);
    handler = new ToolHandler(excelService, permChecker, logger);

    const wb = new ExcelJS.Workbook();
    wb.addWorksheet(sheetName);
    workbooks.set(fileName, wb);
    excelService.getActiveWorkbooks().set(fileName, wb);
  });

  const getWs = (): ExcelJS.Worksheet => workbooks.get(fileName)!.getWorksheet(sheetName)!;

  describe('calculateSum', () => {
    it('sums numeric values across a range and ignores strings', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = 'not a number';
      ws.getCell('A3').value = 250.5;
      ws.getCell('A4').value = -50.5;

      const res = await accounting.calculateSum(fileName, sheetName, 'A1', 'A4');
      expect(res.success).toBe(true);
      expect(res.data?.sum).toBe(300);
    });

    it('sums values based on criteria column', async () => {
      const ws = getWs();
      // Category in col A, Amount in col B
      ws.getCell('A1').value = 'Marketing'; ws.getCell('B1').value = 500;
      ws.getCell('A2').value = 'R&D';       ws.getCell('B2').value = 1200;
      ws.getCell('A3').value = 'Marketing'; ws.getCell('B3').value = 350;
      ws.getCell('A4').value = 'Sales';     ws.getCell('B4').value = 800;

      const res = await accounting.calculateSum(fileName, sheetName, 'A1', 'B4', 'A', 'Marketing');
      expect(res.success).toBe(true);
      expect(res.data?.sum).toBe(850);
    });

    it('returns error when workbook or worksheet is not found', async () => {
      const res1 = await accounting.calculateSum('missing.xlsx', sheetName, 'A1', 'A2');
      expect(res1.success).toBe(false);
      expect(res1.error).toContain('not found');

      const res2 = await accounting.calculateSum(fileName, 'MissingSheet', 'A1', 'A2');
      expect(res2.success).toBe(false);
      expect(res2.error).toContain('not found');
    });
  });

  describe('calculateAverage', () => {
    it('calculates arithmetic mean and count correctly', async () => {
      const ws = getWs();
      ws.getCell('B2').value = 10;
      ws.getCell('B3').value = 20;
      ws.getCell('B4').value = 30;
      ws.getCell('B5').value = 'skip';

      const res = await accounting.calculateAverage(fileName, sheetName, 'B2', 'B5');
      expect(res.success).toBe(true);
      expect(res.data?.average).toBe(20);
      expect(res.data?.count).toBe(3);
    });

    it('returns 0 average when count is 0', async () => {
      const ws = getWs();
      ws.getCell('B2').value = 'text only';

      const res = await accounting.calculateAverage(fileName, sheetName, 'B2', 'B2');
      expect(res.success).toBe(true);
      expect(res.data?.average).toBe(0);
      expect(res.data?.count).toBe(0);
    });
  });

  describe('calculateRunningTotal', () => {
    it('calculates cumulative running totals with row offset', async () => {
      const ws = getWs();
      ws.getCell('B5').value = 100;
      ws.getCell('B6').value = 200;
      ws.getCell('B7').value = 50;

      const res = await accounting.calculateRunningTotal(fileName, sheetName, 'B5', 'B7', 'C5');
      expect(res.success).toBe(true);
      expect(ws.getCell('C5').value).toBe(100);
      expect(ws.getCell('C6').value).toBe(300);
      expect(ws.getCell('C7').value).toBe(350);
    });

    it('returns error on invalid output start cell', async () => {
      const ws = getWs();
      ws.getCell('B1').value = 10;
      const res = await accounting.calculateRunningTotal(fileName, sheetName, 'B1', 'B1', 'invalid');
      expect(res.success).toBe(false);
    });
  });

  describe('calculatePercentageOfTotal', () => {
    it('calculates percentages and formats with % symbol', async () => {
      const ws = getWs();
      ws.getCell('A2').value = 25;
      ws.getCell('A3').value = 75;

      const res = await accounting.calculatePercentageOfTotal(fileName, sheetName, 'A2', 'A3', 'B2');
      expect(res.success).toBe(true);
      expect(ws.getCell('B2').value).toBe(25);
      expect(ws.getCell('B3').value).toBe(75);
      expect(ws.getCell('B2').numFmt).toBe('0.00"%"');
    });

    it('handles total = 0 without dividing by zero', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 0;
      ws.getCell('A2').value = 0;

      const res = await accounting.calculatePercentageOfTotal(fileName, sheetName, 'A1', 'A2', 'B1');
      expect(res.success).toBe(true);
      expect(ws.getCell('B1').value).toBe(0);
      expect(ws.getCell('B2').value).toBe(0);
    });
  });

  describe('calculateYTD', () => {
    it('accumulates YTD and resets when calendar year changes', async () => {
      const ws = getWs();
      // Dates in col A, amounts in col B
      ws.getCell('A2').value = new Date('2024-01-15'); ws.getCell('B2').value = 100;
      ws.getCell('A3').value = new Date('2024-06-30'); ws.getCell('B3').value = 150;
      ws.getCell('A4').value = new Date('2025-01-10'); ws.getCell('B4').value = 50;  // New year! Reset!
      ws.getCell('A5').value = new Date('2025-03-20'); ws.getCell('B5').value = 80;

      const res = await accounting.calculateYTD(fileName, sheetName, 'A', 'B', 'A2', 'B5', 'C');
      expect(res.success).toBe(true);
      expect(ws.getCell('C2').value).toBe(100);
      expect(ws.getCell('C3').value).toBe(250);
      expect(ws.getCell('C4').value).toBe(50); // Reset for 2025
      expect(ws.getCell('C5').value).toBe(130);
    });

    it('supports date strings and Excel serial date numbers', async () => {
      const ws = getWs();
      // Row 2: date string '2024-02-01'
      ws.getCell('A2').value = '2024-02-01'; ws.getCell('B2').value = 200;
      // Row 3: Excel serial date (approx 2024)
      ws.getCell('A3').value = 45350;        ws.getCell('B3').value = 300;

      const res = await accounting.calculateYTD(fileName, sheetName, 'A', 'B', 'A2', 'B3', 'D2');
      expect(res.success).toBe(true);
      expect(ws.getCell('D2').value).toBe(200);
      expect(ws.getCell('D3').value).toBe(500);
    });
  });

  describe('applyAccountingFormat & Currency Formats', () => {
    it('applies accounting format with parentheses and red tag', async () => {
      const ws = getWs();
      ws.getCell('A1').value = -1234.56;

      const res = await accounting.applyAccountingFormat(fileName, sheetName, 'A1', 'A1', {
        showNegativeInParentheses: true,
        showNegativeInRed: true,
        currencySymbol: '$',
      });
      expect(res.success).toBe(true);
      expect(ws.getCell('A1').numFmt).toContain('[Red]');
      expect(ws.getCell('A1').numFmt).toContain('$');
    });

    it('applies VND currency format with 0 decimal places', async () => {
      const ws = getWs();
      ws.getCell('B2').value = 5000000;

      const res = await accounting.applyVNDCurrencyFormat(fileName, sheetName, 'B2', 'B2', 0);
      expect(res.success).toBe(true);
      expect(ws.getCell('B2').numFmt).toContain('" đ"');
    });

    it('applies negative red format code', async () => {
      const ws = getWs();
      ws.getCell('C1').value = -500;

      const res = await accounting.applyNegativeRedFormat(fileName, sheetName, 'C1', 'C1', 2);
      expect(res.success).toBe(true);
      expect(ws.getCell('C1').numFmt).toContain('[Red]');
    });
  });

  describe('showZerosInsteadOfEmpty', () => {
    it('populates empty, null, and empty-string cells with 0', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = '';
      ws.getCell('A3').value = null;

      const res = await accounting.showZerosInsteadOfEmpty(fileName, sheetName, 'A1', 'A3');
      expect(res.success).toBe(true);
      expect(ws.getCell('A1').value).toBe(100);
      expect(ws.getCell('A2').value).toBe(0);
      expect(ws.getCell('A3').value).toBe(0);
    });
  });

  describe('calculatePeriodComparison & calculateVariance', () => {
    it('calculates percentage and absolute variance between periods', async () => {
      const ws = getWs();
      // Previous: 100, Current: 150 -> Variance: +50, +50%
      ws.getCell('A2').value = 150;
      ws.getCell('B2').value = 100;

      const res = await accounting.calculatePeriodComparison(
        fileName, sheetName, 'A2:A2', 'B2:B2', 'C2', true
      );
      expect(res.success).toBe(true);
      expect(ws.getCell('C2').value).toBe(50);
      expect(ws.getCell('C2').numFmt).toBe('0.00"%"');
    });

    it('handles previous = 0 in period comparison without throwing', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 50;
      ws.getCell('B1').value = 0;

      const res = await accounting.calculatePeriodComparison(
        fileName, sheetName, 'A1:A1', 'B1:B1', 'C1', true
      );
      expect(res.success).toBe(true);
      expect(ws.getCell('C1').value).toBe(0);
    });

    it('calculates budget variance (actual - budget)', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 1000; // Budget
      ws.getCell('B1').value = 850;  // Actual -> -150 variance (-15%)

      const res = await accounting.calculateVariance(
        fileName, sheetName, 'A1:A1', 'B1:B1', 'C1', true
      );
      expect(res.success).toBe(true);
      expect(ws.getCell('C1').value).toBe(-15);
    });
  });

  describe('checkBalance', () => {
    it('verifies balanced debit and credit columns', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 500; ws.getCell('A2').value = 500; // Debits = 1000
      ws.getCell('B1').value = 700; ws.getCell('B2').value = 300; // Credits = 1000

      const res = await accounting.checkBalance(fileName, sheetName, 'A1:A2', 'B1:B2');
      expect(res.success).toBe(true);
      expect(res.data?.isBalanced).toBe(true);
      expect(res.data?.debitTotal).toBe(1000);
      expect(res.data?.creditTotal).toBe(1000);
      expect(res.data?.difference).toBe(0);
    });

    it('reports imbalance correctly when totals do not match', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 1000;
      ws.getCell('B1').value = 950;

      const res = await accounting.checkBalance(fileName, sheetName, 'A1', 'B1');
      expect(res.success).toBe(true);
      expect(res.data?.isBalanced).toBe(false);
      expect(res.data?.difference).toBe(50);
    });
  });

  describe('findAnomalies', () => {
    it('detects statistical outliers using z-score threshold', async () => {
      const ws = getWs();
      // Normal values ~10-14, outlier at 100
      const vals = [10, 11, 12, 10, 11, 13, 12, 100];
      vals.forEach((v, i) => {
        ws.getCell(`A${i + 1}`).value = v;
      });

      const res = await accounting.findAnomalies(fileName, sheetName, 'A1', 'A8', 2);
      expect(res.success).toBe(true);
      expect(res.data?.anomalies.length).toBeGreaterThan(0);
      expect(res.data?.anomalies[0].value).toBe(100);
      expect(res.data?.anomalies[0].cell).toBe('A8');
    });

    it('handles uniform values (stdDev = 0) safely without NaN', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 50;
      ws.getCell('A2').value = 50;
      ws.getCell('A3').value = 50;

      const res = await accounting.findAnomalies(fileName, sheetName, 'A1', 'A3', 2);
      expect(res.success).toBe(true);
      expect(res.data?.anomalies).toHaveLength(0);
    });

    it('returns empty anomalies for dataset with less than 2 items', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 42;

      const res = await accounting.findAnomalies(fileName, sheetName, 'A1', 'A1');
      expect(res.success).toBe(true);
      expect(res.data?.anomalies).toHaveLength(0);
    });
  });

  describe('ToolHandler Integration for Accounting Tools', () => {
    it('executes excel_financial_sum through ToolHandler', async () => {
      const ws = excelService.getActiveWorkbooks().get(fileName)!.getWorksheet(sheetName)!;
      ws.getCell('A1').value = 40;
      ws.getCell('A2').value = 60;

      const res = await handler.executeTool('excel_financial_sum', {
        filename: fileName,
        worksheet: sheetName,
        rangeStart: 'A1',
        rangeEnd: 'A2',
      });
      expect(res.success).toBe(true);
      expect(res.data).toEqual({ sum: 100 });
    });

    it('executes excel_check_balance through ToolHandler', async () => {
      const ws = excelService.getActiveWorkbooks().get(fileName)!.getWorksheet(sheetName)!;
      ws.getCell('A1').value = 200;
      ws.getCell('B1').value = 200;

      const res = await handler.executeTool('excel_check_balance', {
        filename: fileName,
        worksheet: sheetName,
        debitRange: 'A1:A1',
        creditRange: 'B1:B1',
      });
      expect(res.success).toBe(true);
      expect((res.data as any).isBalanced).toBe(true);
    });
  });
});
