/**
 * Financial correctness tests for ExcelAdvancedAccounting
 *
 * Regression coverage for audit findings:
 * - 4.1: IRR Newton-Raphson derivative was mathematically wrong
 * - 4.2: Aging report ignored raw Excel serial dates
 */
import ExcelJS from 'exceljs';
import { ExcelAdvancedAccounting } from '../src/services/excel-advanced-accounting.js';
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
    new Logger('error'),
    workbooks
  );
  return { service, workbooks };
};

/** Registers a workbook named 'fin.xlsx' with one sheet holding a column of values. */
const addCashflowSheet = (
  workbooks: Map<string, ExcelJS.Workbook>,
  sheetName: string,
  values: (number | string | Date)[]
): void => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  values.forEach((v, i) => {
    ws.getCell(`A${i + 1}`).value = v;
  });
  workbooks.set('fin.xlsx', wb);
};

describe('calculateIRR — audit Finding 4.1: Newton-Raphson correctness', () => {
  it('computes the exact IRR for a single-period flow: [-100, 110] → 10%', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-100, 110]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A2', 0.1);
    expect(result.success).toBe(true);
    expect(result.data?.irr).toBeCloseTo(10, 6);
  });

  it('computes IRR for multi-period flows: [-1000, 500, 500, 500] ≈ 23.37%', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-1000, 500, 500, 500]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A4', 0.1);
    expect(result.success).toBe(true);
    expect(result.data?.irr).toBeCloseTo(23.37, 1);
  });

  it('uses Excel semantics: the first cash flow sits at t=0 (undiscounted)', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-100, 110]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A2', 0.05);
    expect(result.success).toBe(true);
    expect(result.data?.irr).toBeCloseTo(10, 6);
  });

  it('rejects all-positive flows (no sign change → no IRR exists)', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [100, 110, 120]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A3', 0.1);
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/negative and one positive/i);
  });

  it('rejects all-negative flows', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-100, -110]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A2', 0.1);
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/negative and one positive/i);
  });

  it('rejects fewer than two values', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-100]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A1', 0.1);
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/at least two/i);
  });

  it('skips non-numeric cells in the range', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-100, 'n/a', 110]);
    const result = await service.calculateIRR('fin.xlsx', 'CF', 'A1:A3', 0.1);
    expect(result.success).toBe(true);
    expect(result.data?.irr).toBeCloseTo(10, 6);
  });

  it('fails when the workbook was never opened', async () => {
    const { service } = makeService();
    const result = await service.calculateIRR('ghost.xlsx', 'CF', 'A1:A2', 0.1);
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  it('fails when the worksheet does not exist', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [-100, 110]);
    const result = await service.calculateIRR('fin.xlsx', 'Missing', 'A1:A2', 0.1);
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });
});

describe('calculateNPV — Excel NPV semantics (first value discounted at t=1)', () => {
  it('computes NPV for [100, 100, 100] at 10% ≈ 248.69', async () => {
    const { service, workbooks } = makeService();
    addCashflowSheet(workbooks, 'CF', [100, 100, 100]);
    const result = await service.calculateNPV('fin.xlsx', 'CF', 0.1, 'A1:A3');
    expect(result.success).toBe(true);
    expect(result.data?.npv).toBeCloseTo(248.69, 1);
  });
});

describe('createAgingReport — audit Finding 4.2: Excel serial date support', () => {
  // Canonical anchor: Excel serial 25569 === 1970-01-01 (the Unix epoch)
  const AS_OF = new Date('1970-03-15T00:00:00Z');

  const buildAgingSheet = (rows: { date: number | string | Date | null; amount: number }[]) => {
    const { service, workbooks } = makeService();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('AR');
    ws.getCell('A1').value = 'Invoice Date'; // header must be skipped
    ws.getCell('B1').value = 'Amount';
    rows.forEach((r, i) => {
      const row = i + 2;
      if (r.date !== null) {
        ws.getCell(`A${row}`).value = r.date;
      }
      ws.getCell(`B${row}`).value = r.amount;
    });
    workbooks.set('aging.xlsx', wb);
    return { service, workbooks, ws };
  };

  it('buckets invoices stored as raw Excel serial numbers (serial 25569 → 73 days → 61-90)', async () => {
    const { service } = buildAgingSheet([{ date: 25569, amount: 100 }]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    expect(result.data?.summary['61-90']).toBe(100);
  });

  it('buckets invoices stored as JavaScript Date objects (1970-01-15 → 59 days → 31-60)', async () => {
    const { service } = buildAgingSheet([{ date: new Date('1970-01-15T00:00:00Z'), amount: 200 }]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    expect(result.data?.summary['31-60']).toBe(200);
  });

  it('buckets invoices stored as date strings (1970-01-10 → 64 days → 61-90)', async () => {
    const { service } = buildAgingSheet([{ date: '1970-01-10', amount: 50 }]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    expect(result.data?.summary['61-90']).toBe(50);
  });

  it('places future-dated invoices into the 0-30 bucket', async () => {
    const { service } = buildAgingSheet([{ date: new Date('2999-01-01T00:00:00Z'), amount: 75 }]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    expect(result.data?.summary['0-30']).toBe(75);
  });

  it('ignores serial numbers inside the ambiguous Lotus-1900 window (serial < 61)', async () => {
    const { service } = buildAgingSheet([{ date: 60, amount: 500 }]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    const summary = result.data?.summary as Record<string, number>;
    expect(Object.values(summary).reduce((a, b) => a + b, 0)).toBe(0);
  });

  it('skips headers, garbage strings, and blank date cells but keeps valid rows', async () => {
    const { service } = buildAgingSheet([
      { date: 25569, amount: 100 },
      { date: 'not-a-date', amount: 999 },
      { date: null, amount: 80 },
      { date: new Date('1969-01-01T00:00:00Z'), amount: 25 },
    ]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    expect(result.data?.summary).toEqual({
      '0-30': 0,
      '31-60': 0,
      '61-90': 100,
      '91-120': 0,
      '120+': 25,
    });
  });

  it('computes a correct combined summary across mixed date representations', async () => {
    const { service } = buildAgingSheet([
      { date: 25569, amount: 100 },
      { date: new Date('1970-01-15T00:00:00Z'), amount: 200 },
      { date: '1970-01-10', amount: 50 },
      { date: new Date('2999-01-01T00:00:00Z'), amount: 75 },
    ]);
    const result = await service.createAgingReport('aging.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(true);
    expect(result.data?.summary).toEqual({
      '0-30': 75,
      '31-60': 200,
      '61-90': 150,
      '91-120': 0,
      '120+': 0,
    });
  });

  it('fails when the workbook was never opened', async () => {
    const { service } = makeService();
    const result = await service.createAgingReport('ghost.xlsx', 'AR', 'A', 'B', AS_OF, 'D1');
    expect(result.success).toBe(false);
  });
});
