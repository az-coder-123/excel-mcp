/**
 * Tests for new accounting tools:
 * - calculateDepreciation (Straight-Line, DDB, SYD)
 * - calculateProgressiveTax (graduated brackets)
 * - calculateXIRR (irregular cash flows)
 */
import ExcelJS from 'exceljs';
import { ExcelAdvancedAccounting } from '../src/services/excel-advanced-accounting.js';
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

const makeService = () => {
  const workbooks = new Map<string, ExcelJS.Workbook>();
  const service = new ExcelAdvancedAccounting(
    new PermissionChecker(config()),
    new Logger('silent'),
    workbooks
  );
  return { service, workbooks };
};

/** Helper: register a workbook with a sheet containing a single column of numbers */
const addColumnSheet = (
  workbooks: Map<string, ExcelJS.Workbook>,
  name: string,
  sheetName: string,
  values: (number | string | Date)[]
): void => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  values.forEach((v, i) => {
    ws.getCell(`A${i + 1}`).value = v;
  });
  workbooks.set(name, wb);
};

/** Helper: register a workbook with two columns (dates + values) */
const addDateValueSheet = (
  workbooks: Map<string, ExcelJS.Workbook>,
  name: string,
  sheetName: string,
  dates: Date[],
  values: number[]
): void => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  for (let i = 0; i < dates.length; i++) {
    ws.getCell(`A${i + 1}`).value = dates[i];
    ws.getCell(`B${i + 1}`).value = values[i];
  }
  workbooks.set(name, wb);
};

// ===================================================================
// Depreciation Tests
// ===================================================================
describe('calculateDepreciation', () => {
  describe('Straight-Line', () => {
    it('calculates equal annual depreciation', async () => {
      const { service, workbooks } = makeService();
      // Need a workbook with a sheet to write to
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        10000, 2000, 5, 'straight-line'
      );

      expect(result.success).toBe(true);
      expect(result.data?.schedule).toHaveLength(5);

      // Depreciable base = 10000 - 2000 = 8000, annual = 8000/5 = 1600
      for (const entry of result.data!.schedule) {
        expect(entry.depreciation).toBeCloseTo(1600, 2);
      }

      // Final book value should equal salvage
      expect(result.data!.schedule[4].bookValue).toBeCloseTo(2000, 2);
      expect(result.data!.schedule[4].accumulated).toBeCloseTo(8000, 2);
    });
  });

  describe('Double-Declining Balance', () => {
    it('calculates accelerated depreciation correctly', async () => {
      const { service, workbooks } = makeService();
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        10000, 1000, 5, 'double-declining'
      );

      expect(result.success).toBe(true);
      expect(result.data?.schedule).toHaveLength(5);

      // DDB rate = 2/5 = 40%
      // Year 1: 10000 * 0.4 = 4000 → BV 6000
      expect(result.data!.schedule[0].depreciation).toBeCloseTo(4000, 2);
      expect(result.data!.schedule[0].bookValue).toBeCloseTo(6000, 2);

      // Year 2: 6000 * 0.4 = 2400 → BV 3600
      expect(result.data!.schedule[1].depreciation).toBeCloseTo(2400, 2);
      expect(result.data!.schedule[1].bookValue).toBeCloseTo(3600, 2);

      // Final book value should not drop below salvage (1000)
      expect(result.data!.schedule[4].bookValue).toBeGreaterThanOrEqual(1000);
    });

    it('stops depreciating when book value equals salvage value', async () => {
      const { service, workbooks } = makeService();
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      // High salvage relative to cost
      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        10000, 5000, 3, 'double-declining'
      );

      expect(result.success).toBe(true);
      // Book value should never go below 5000
      for (const entry of result.data!.schedule) {
        expect(entry.bookValue).toBeGreaterThanOrEqual(4999.99); // rounding tolerance
      }
    });
  });

  describe('Sum-of-Years-Digits', () => {
    it('calculates SYD depreciation correctly', async () => {
      const { service, workbooks } = makeService();
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        10000, 1000, 5, 'sum-of-years-digits'
      );

      expect(result.success).toBe(true);
      expect(result.data?.schedule).toHaveLength(5);

      // SYD denominator = 5*6/2 = 15, depreciable = 9000
      // Year 1: 9000 * 5/15 = 3000
      expect(result.data!.schedule[0].depreciation).toBeCloseTo(3000, 2);
      // Year 2: 9000 * 4/15 = 2400
      expect(result.data!.schedule[1].depreciation).toBeCloseTo(2400, 2);
      // Year 5: 9000 * 1/15 = 600
      expect(result.data!.schedule[4].depreciation).toBeCloseTo(600, 2);

      // Final book value should equal salvage
      expect(result.data!.schedule[4].bookValue).toBeCloseTo(1000, 2);
    });
  });

  describe('Edge cases', () => {
    it('rejects zero cost', async () => {
      const { service, workbooks } = makeService();
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        0, 0, 5, 'straight-line'
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Cost must be greater than zero');
    });

    it('rejects salvage >= cost', async () => {
      const { service, workbooks } = makeService();
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        10000, 10000, 5, 'straight-line'
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Salvage value must be less than cost');
    });

    it('rejects non-integer useful life', async () => {
      const { service, workbooks } = makeService();
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Assets');
      workbooks.set('assets.xlsx', wb);

      const result = await service.calculateDepreciation(
        'assets.xlsx', 'Assets', 'A1',
        10000, 1000, 2.5, 'straight-line'
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('Useful life must be a positive integer');
    });

    it('returns error for missing workbook', async () => {
      const { service } = makeService();
      const result = await service.calculateDepreciation(
        'missing.xlsx', 'Sheet', 'A1',
        10000, 1000, 5, 'straight-line'
      );
      expect(result.success).toBe(false);
      expect(result.error).toContain('not found');
    });
  });
});

// ===================================================================
// Progressive Tax Tests
// ===================================================================
describe('calculateProgressiveTax', () => {
  it('calculates Vietnam PIT-like progressive tax correctly', async () => {
    const { service, workbooks } = makeService();
    addColumnSheet(workbooks, 'tax.xlsx', 'Income', [60_000_000]);

    // Simplified VN PIT brackets (monthly, VND)
    const brackets = [
      { threshold: 0,          rate: 5 },
      { threshold: 5_000_000,  rate: 10 },
      { threshold: 10_000_000, rate: 15 },
      { threshold: 18_000_000, rate: 20 },
      { threshold: 32_000_000, rate: 25 },
      { threshold: 52_000_000, rate: 30 },
      { threshold: 80_000_000, rate: 35 },
    ];

    const result = await service.calculateProgressiveTax(
      'tax.xlsx', 'Income', 'A1:A1', brackets, 'B1'
    );

    expect(result.success).toBe(true);
    expect(result.data?.details).toHaveLength(1);

    // Manual calculation for 60,000,000:
    // 0–5M:  5M * 5%   = 250,000
    // 5–10M: 5M * 10%  = 500,000
    // 10–18M: 8M * 15% = 1,200,000
    // 18–32M: 14M * 20% = 2,800,000
    // 32–52M: 20M * 25% = 5,000,000
    // 52–60M: 8M * 30%  = 2,400,000
    // Total = 12,150,000
    expect(result.data!.details[0].tax).toBeCloseTo(12_150_000, 0);
    expect(result.data!.totalTax).toBeCloseTo(12_150_000, 0);
  });

  it('returns 0 tax for amount below first bracket', async () => {
    const { service, workbooks } = makeService();
    addColumnSheet(workbooks, 'tax.xlsx', 'Income', [3_000_000]);

    const brackets = [
      { threshold: 5_000_000,  rate: 10 },
      { threshold: 10_000_000, rate: 20 },
    ];

    const result = await service.calculateProgressiveTax(
      'tax.xlsx', 'Income', 'A1:A1', brackets, 'B1'
    );

    expect(result.success).toBe(true);
    expect(result.data!.details[0].tax).toBe(0);
  });

  it('handles multiple incomes in range', async () => {
    const { service, workbooks } = makeService();
    addColumnSheet(workbooks, 'tax.xlsx', 'Income', [10_000, 20_000, 50_000]);

    const brackets = [
      { threshold: 0,      rate: 10 },
      { threshold: 15_000, rate: 20 },
      { threshold: 30_000, rate: 30 },
    ];

    const result = await service.calculateProgressiveTax(
      'tax.xlsx', 'Income', 'A1:A3', brackets, 'B1'
    );

    expect(result.success).toBe(true);
    expect(result.data?.details).toHaveLength(3);

    // Income 10,000: 0–10,000 @ 10% = 1,000
    expect(result.data!.details[0].tax).toBeCloseTo(1_000, 2);

    // Income 20,000: 0–15,000 @ 10% + 15,000–20,000 @ 20% = 1,500 + 1,000 = 2,500
    expect(result.data!.details[1].tax).toBeCloseTo(2_500, 2);
  });

  it('rejects unsorted brackets', async () => {
    const { service, workbooks } = makeService();
    addColumnSheet(workbooks, 'tax.xlsx', 'Income', [10_000]);

    const brackets = [
      { threshold: 20_000, rate: 20 },
      { threshold: 10_000, rate: 10 },
    ];

    const result = await service.calculateProgressiveTax(
      'tax.xlsx', 'Income', 'A1:A1', brackets, 'B1'
    );

    expect(result.success).toBe(false);
    expect(result.error).toContain('sorted');
  });

  it('rejects empty brackets', async () => {
    const { service, workbooks } = makeService();
    addColumnSheet(workbooks, 'tax.xlsx', 'Income', [10_000]);

    const result = await service.calculateProgressiveTax(
      'tax.xlsx', 'Income', 'A1:A1', [], 'B1'
    );

    expect(result.success).toBe(false);
    expect(result.error).toContain('At least one tax bracket');
  });
});

// ===================================================================
// XIRR Tests
// ===================================================================
describe('calculateXIRR', () => {
  it('computes XIRR for regular periodic flows (should ≈ IRR)', async () => {
    const { service, workbooks } = makeService();

    // Investment: -1000 on Jan 1, then 3 equal payments of 400 every year
    // Regular IRR ≈ 9.7%
    const dates = [
      new Date('2024-01-01'),
      new Date('2025-01-01'),
      new Date('2026-01-01'),
      new Date('2027-01-01'),
    ];
    const values = [-1000, 400, 400, 400];

    addDateValueSheet(workbooks, 'xirr.xlsx', 'CF', dates, values);

    const result = await service.calculateXIRR(
      'xirr.xlsx', 'CF', 'A1:A4', 'B1:B4', 0.1
    );

    expect(result.success).toBe(true);
    // Should be close to standard IRR for equal periods (~9.7%)
    expect(result.data!.xirr).toBeCloseTo(9.7, 0);
  });

  it('computes XIRR for irregular cash flows', async () => {
    const { service, workbooks } = makeService();

    // Invest -10000 on Jan 1, get 2750 on Mar 1, 4250 on Jul 15, 3250 on Dec 1
    const dates = [
      new Date('2024-01-01'),
      new Date('2024-03-01'),
      new Date('2024-07-15'),
      new Date('2024-12-01'),
    ];
    const values = [-10000, 2750, 4250, 3250];

    addDateValueSheet(workbooks, 'xirr.xlsx', 'CF', dates, values);

    const result = await service.calculateXIRR(
      'xirr.xlsx', 'CF', 'A1:A4', 'B1:B4', 0.1
    );

    expect(result.success).toBe(true);
    // The XIRR should be a positive rate (the investor gets 250 in profit within a year)
    expect(result.data!.xirr).toBeGreaterThan(0);
  });

  it('computes known XIRR: [-100, 110] after exactly 1 year = 10%', async () => {
    const { service, workbooks } = makeService();

    const dates = [
      new Date('2024-01-01'),
      new Date('2025-01-01'),
    ];
    const values = [-100, 110];

    addDateValueSheet(workbooks, 'xirr.xlsx', 'CF', dates, values);

    const result = await service.calculateXIRR(
      'xirr.xlsx', 'CF', 'A1:A2', 'B1:B2', 0.1
    );

    expect(result.success).toBe(true);
    expect(result.data!.xirr).toBeCloseTo(10, 1);
  });

  it('rejects all-positive cash flows', async () => {
    const { service, workbooks } = makeService();

    const dates = [new Date('2024-01-01'), new Date('2025-01-01')];
    const values = [100, 200];

    addDateValueSheet(workbooks, 'xirr.xlsx', 'CF', dates, values);

    const result = await service.calculateXIRR(
      'xirr.xlsx', 'CF', 'A1:A2', 'B1:B2'
    );

    expect(result.success).toBe(false);
    expect(result.error).toContain('negative and one positive');
  });

  it('rejects mismatched date/value lengths', async () => {
    const { service, workbooks } = makeService();

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('CF');
    ws.getCell('A1').value = new Date('2024-01-01');
    ws.getCell('A2').value = new Date('2025-01-01');
    ws.getCell('B1').value = -100;
    // B2 is missing → 2 date rows but effectively 1 value row won't cause length mismatch
    // since getRangeData reads by row. Let's set up explicit length mismatch:
    // Use different row counts for date and value ranges
    ws.getCell('A3').value = new Date('2026-01-01');
    // B2 and B3 are empty — still 3 rows each, but values will be skipped as non-numeric
    workbooks.set('xirr.xlsx', wb);

    const result = await service.calculateXIRR(
      'xirr.xlsx', 'CF', 'A1:A3', 'B1:B2'
    );

    expect(result.success).toBe(false);
    expect(result.error).toContain('same number of rows');
  });

  it('returns error for missing workbook', async () => {
    const { service } = makeService();
    const result = await service.calculateXIRR(
      'missing.xlsx', 'CF', 'A1:A2', 'B1:B2'
    );
    expect(result.success).toBe(false);
    expect(result.error).toContain('not found');
  });
});

describe('ToolHandler integration & ExcelService Facade', () => {
  let excelService: ExcelService;
  let handler: ToolHandler;
  const testFile = 'integration.xlsx';

  beforeEach(async () => {
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    excelService = new ExcelService(permChecker, logger);
    handler = new ToolHandler(excelService, permChecker, logger);

    // Setup an in-memory workbook directly in activeWorkbooks
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1');
    excelService.getActiveWorkbooks().set(testFile, wb);
  });

  it('executes excel_calculate_depreciation through ToolHandler', async () => {
    const res = await handler.executeTool('excel_calculate_depreciation', {
      filename: testFile,
      worksheet: 'Sheet1',
      startCell: 'A1',
      cost: 100000,
      salvageValue: 10000,
      usefulLife: 5,
      method: 'straight-line',
    });

    expect(res.success).toBe(true);
    const data = res.data as { schedule: { year: number; depreciation: number; accumulated: number; bookValue: number }[] };
    expect(data.schedule).toHaveLength(5);
    expect(data.schedule[0].depreciation).toBe(18000);
  });

  it('executes excel_calculate_progressive_tax through ToolHandler', async () => {
    const ws = excelService.getActiveWorkbooks().get(testFile)!.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = 10000000;

    const res = await handler.executeTool('excel_calculate_progressive_tax', {
      filename: testFile,
      worksheet: 'Sheet1',
      amountRange: 'A1:A1',
      brackets: [
        { threshold: 5000000, rate: 5 },
        { threshold: 10000000, rate: 10 },
      ],
      outputRange: 'B1',
    });

    expect(res.success).toBe(true);
    const data = res.data as { totalTax: number };
    expect(data.totalTax).toBe(250000);
    expect(ws.getCell('B1').value).toBe(250000);
  });

  it('executes excel_calculate_xirr through ToolHandler', async () => {
    const ws = excelService.getActiveWorkbooks().get(testFile)!.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = new Date('2024-01-01');
    ws.getCell('A2').value = new Date('2025-01-01');
    ws.getCell('B1').value = -1000;
    ws.getCell('B2').value = 1100;

    const res = await handler.executeTool('excel_calculate_xirr', {
      filename: testFile,
      worksheet: 'Sheet1',
      dateRange: 'A1:A2',
      valuesRange: 'B1:B2',
    });

    expect(res.success).toBe(true);
    const data = res.data as { xirr: number };
    expect(data.xirr).toBeCloseTo(10, 1);
  });

  it('validates invalid params via ToolHandler Zod schema', async () => {
    // Missing required field 'cost'
    const res = await handler.executeTool('excel_calculate_depreciation', {
      filename: testFile,
      worksheet: 'Sheet1',
      startCell: 'A1',
      // cost missing
      salvageValue: 1000,
      usefulLife: 5,
      method: 'straight-line',
    });

    expect(res.success).toBe(false);
    expect(res.error).toContain('Invalid parameters for excel_calculate_depreciation');
  });

  it('calls ExcelService facade methods directly', async () => {
    const depResult = await excelService.calculateDepreciation(
      testFile,
      'Sheet1',
      'D1',
      50000,
      5000,
      3,
      'straight-line'
    );
    expect(depResult.success).toBe(true);
    expect(depResult.data?.schedule).toHaveLength(3);

    const ws = excelService.getActiveWorkbooks().get(testFile)!.getWorksheet('Sheet1')!;
    ws.getCell('F1').value = 15000000;
    const taxResult = await excelService.calculateProgressiveTax(
      testFile,
      'Sheet1',
      'F1:F1',
      [{ threshold: 10000000, rate: 10 }],
      'G1'
    );
    expect(taxResult.success).toBe(true);
    expect(taxResult.data?.totalTax).toBe(500000);

    ws.getCell('A1').value = new Date('2024-01-01');
    ws.getCell('A2').value = new Date('2025-01-01');
    ws.getCell('B1').value = -1000;
    ws.getCell('B2').value = 1100;

    const xirrResult = await excelService.calculateXIRR(
      testFile,
      'Sheet1',
      'A1:A2',
      'B1:B2'
    );
    expect(xirrResult.success).toBe(true);
    expect(xirrResult.data?.xirr).toBeCloseTo(10, 1);
  });
});

