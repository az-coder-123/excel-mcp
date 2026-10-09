/**
 * Tests for Excel coordinate utilities, formula auto-detection,
 * multi-column accounting calculations, and SystemHandlers typed dependencies.
 */

import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  columnLetterToNumber,
  columnToNumber,
  numberToColumn,
  parseCellAddress,
  parseCellRange,
} from '../src/utils/excel-coords.js';
import { ExcelCellOperations } from '../src/services/excel-cell-operations.js';
import { ExcelAccounting } from '../src/services/excel-accounting.js';
import { ExcelService } from '../src/services/excel-service.js';
import { SystemHandlers } from '../src/tools/handlers/system-handlers.js';
import { PermissionChecker } from '../src/security/permission-checker.js';
import { Logger } from '../src/utils/logger.js';
import { PermissionConfig } from '../src/types/index.js';

describe('Excel Coordinate Utilities (excel-coords)', () => {
  describe('columnLetterToNumber & columnToNumber', () => {
    it('correctly converts single-letter columns', () => {
      expect(columnLetterToNumber('A')).toBe(1);
      expect(columnLetterToNumber('B')).toBe(2);
      expect(columnLetterToNumber('Z')).toBe(26);
    });

    it('correctly converts multi-letter columns', () => {
      expect(columnLetterToNumber('AA')).toBe(27);
      expect(columnLetterToNumber('AB')).toBe(28);
      expect(columnLetterToNumber('AZ')).toBe(52);
      expect(columnLetterToNumber('BA')).toBe(53);
      expect(columnLetterToNumber('ZZ')).toBe(702);
      expect(columnLetterToNumber('AAA')).toBe(703);
    });

    it('handles lowercase column letters gracefully', () => {
      expect(columnLetterToNumber('a')).toBe(1);
      expect(columnLetterToNumber('z')).toBe(26);
      expect(columnLetterToNumber('aa')).toBe(27);
    });

    it('columnToNumber is identical alias', () => {
      expect(columnToNumber('C')).toBe(3);
      expect(columnToNumber('XFD')).toBe(16384);
    });
  });

  describe('numberToColumn', () => {
    it('converts indices back to column letters', () => {
      expect(numberToColumn(1)).toBe('A');
      expect(numberToColumn(26)).toBe('Z');
      expect(numberToColumn(27)).toBe('AA');
      expect(numberToColumn(28)).toBe('AB');
      expect(numberToColumn(702)).toBe('ZZ');
      expect(numberToColumn(703)).toBe('AAA');
      expect(numberToColumn(16384)).toBe('XFD');
    });

    it('round-trips index -> column -> index for large sample', () => {
      for (const idx of [1, 5, 26, 27, 52, 53, 702, 703, 1000, 16384]) {
        expect(columnLetterToNumber(numberToColumn(idx))).toBe(idx);
      }
    });
  });

  describe('parseCellAddress', () => {
    it('parses valid cell addresses', () => {
      expect(parseCellAddress('A1')).toEqual({ row: 1, column: 1, colLetter: 'A' });
      expect(parseCellAddress('b10')).toEqual({ row: 10, column: 2, colLetter: 'B' });
      expect(parseCellAddress('AA999')).toEqual({ row: 999, column: 27, colLetter: 'AA' });
    });

    it('returns null for invalid inputs', () => {
      expect(parseCellAddress('invalid')).toBeNull();
      expect(parseCellAddress('1A')).toBeNull();
      expect(parseCellAddress('')).toBeNull();
    });
  });

  describe('parseCellRange', () => {
    it('parses colon-separated ranges', () => {
      const parsed = parseCellRange('A1:C10');
      expect(parsed).toEqual({
        start: { row: 1, column: 1, colLetter: 'A' },
        end: { row: 10, column: 3, colLetter: 'C' },
      });
    });

    it('normalizes reversed cell ranges', () => {
      const parsed = parseCellRange('C10', 'A1');
      expect(parsed).toEqual({
        start: { row: 1, column: 1, colLetter: 'A' },
        end: { row: 10, column: 3, colLetter: 'C' },
      });
    });

    it('parses single cell as 1x1 range', () => {
      const parsed = parseCellRange('B5');
      expect(parsed).toEqual({
        start: { row: 5, column: 2, colLetter: 'B' },
        end: { row: 5, column: 2, colLetter: 'B' },
      });
    });
  });
});

describe('Formula Auto-Detection in Cell Operations', () => {
  let tmpRoot: string;
  let activeWorkbooks: Map<string, ExcelJS.Workbook>;
  let cellOps: ExcelCellOperations;
  let testWorkbook: ExcelJS.Workbook;
  let testWorksheet: ExcelJS.Worksheet;
  const filename = 'formula-test.xlsx';

  beforeAll(() => {
    tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-mcp-formula-'));
    const checker = new PermissionChecker({
      allowedPaths: [tmpRoot],
      deniedPaths: [],
      maxFileSize: 50 * 1024 * 1024,
      allowedExtensions: ['.xlsx'],
      permissions: ['read', 'write', 'delete'],
    });
    activeWorkbooks = new Map();
    cellOps = new ExcelCellOperations(checker, new Logger('silent'), activeWorkbooks);

    testWorkbook = new ExcelJS.Workbook();
    testWorksheet = testWorkbook.addWorksheet('Sheet1');
    activeWorkbooks.set(filename, testWorkbook);
  });

  afterAll(() => {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  });

  it('automatically sets formula when string starts with =', async () => {
    const result = await cellOps.writeCell(filename, 'Sheet1', 'C1', '=SUM(A1:B1)');
    expect(result.success).toBe(true);

    const cell = testWorksheet.getCell('C1');
    expect(cell.formula).toBe('SUM(A1:B1)');
  });

  it('preserves regular text strings that do not start with =', async () => {
    const result = await cellOps.writeCell(filename, 'Sheet1', 'A1', 'Total Revenue');
    expect(result.success).toBe(true);

    const cell = testWorksheet.getCell('A1');
    expect(cell.value).toBe('Total Revenue');
    expect(cell.formula).toBeUndefined();
  });

  it('preserves numbers and boolean values intact', async () => {
    await cellOps.writeCell(filename, 'Sheet1', 'A2', 45000);
    await cellOps.writeCell(filename, 'Sheet1', 'A3', true);

    expect(testWorksheet.getCell('A2').value).toBe(45000);
    expect(testWorksheet.getCell('A3').value).toBe(true);
  });

  it('supports formula auto-detection in writeBatch', async () => {
    const batchData = [
      { cellAddress: 'D1', value: 10 },
      { cellAddress: 'D2', value: 20 },
      { cellAddress: 'D3', value: '=AVERAGE(D1:D2)' },
    ];
    const result = await cellOps.writeBatch(filename, 'Sheet1', batchData);
    expect(result.success).toBe(true);
    expect(result.data?.count).toBe(3);

    expect(testWorksheet.getCell('D1').value).toBe(10);
    expect(testWorksheet.getCell('D2').value).toBe(20);
    expect(testWorksheet.getCell('D3').formula).toBe('AVERAGE(D1:D2)');
  });
});

describe('Multi-Column Accounting Operations', () => {
  let tmpRoot: string;
  let activeWorkbooks: Map<string, ExcelJS.Workbook>;
  let accounting: ExcelAccounting;
  const filename = 'multi-col-test.xlsx';

  beforeAll(() => {
    tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-mcp-multicol-'));
    const checker = new PermissionChecker({
      allowedPaths: [tmpRoot],
      deniedPaths: [],
      maxFileSize: 50 * 1024 * 1024,
      allowedExtensions: ['.xlsx'],
      permissions: ['read', 'write', 'delete'],
    });
    activeWorkbooks = new Map();
    accounting = new ExcelAccounting(checker, new Logger('silent'), activeWorkbooks);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');

    // Matrix 3x3:
    // Row 1: [10, 20, 30]
    // Row 2: [40, 50, 60]
    // Row 3: [70, 80, 90]
    ws.getCell('A1').value = 10;
    ws.getCell('B1').value = 20;
    ws.getCell('C1').value = 30;

    ws.getCell('A2').value = 40;
    ws.getCell('B2').value = 50;
    ws.getCell('C2').value = 60;

    ws.getCell('A3').value = 70;
    ws.getCell('B3').value = 80;
    ws.getCell('C3').value = 90;

    // Criteria sheet
    const ws2 = wb.addWorksheet('Criteria');
    ws2.getCell('A1').value = 'North';
    ws2.getCell('B1').value = 100;
    ws2.getCell('C1').value = 50;

    ws2.getCell('A2').value = 'South';
    ws2.getCell('B2').value = 200;
    ws2.getCell('C2').value = 75;

    ws2.getCell('A3').value = 'North';
    ws2.getCell('B3').value = 300;
    ws2.getCell('C3').value = 25;

    activeWorkbooks.set(filename, wb);
  });

  afterAll(() => {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  });

  it('sums all numbers across a multi-column range (A1:C3 = 450)', async () => {
    const result = await accounting.calculateSum(filename, 'Data', 'A1', 'C3');
    expect(result.success).toBe(true);
    expect(result.data?.sum).toBe(450); // 10+20+30 + 40+50+60 + 70+80+90
  });

  it('calculates average across a multi-column range (A1:C3 = 50 across 9 cells)', async () => {
    const result = await accounting.calculateAverage(filename, 'Data', 'A1', 'C3');
    expect(result.success).toBe(true);
    expect(result.data?.average).toBe(50);
    expect(result.data?.count).toBe(9);
  });

  it('calculates sum with criteria across multiple columns', async () => {
    // Sum rows where column A is 'North':
    // Row 1: B1 (100) + C1 (50) = 150
    // Row 3: B3 (300) + C3 (25) = 325
    // Total = 475
    const result = await accounting.calculateSum(filename, 'Criteria', 'A1', 'C3', 'A', 'North');
    expect(result.success).toBe(true);
    expect(result.data?.sum).toBe(475);
  });
});

describe('SystemHandlers Typed Dependencies', () => {
  it('instantiates cleanly with default fallback dependencies', async () => {
    const permConfig: PermissionConfig = {
      allowedPaths: [os.tmpdir()],
      deniedPaths: [],
      maxFileSize: 50 * 1024 * 1024,
      allowedExtensions: ['.xlsx'],
      permissions: ['read', 'write', 'delete', 'admin'],
    };
    const checker = new PermissionChecker(permConfig);
    const logger = new Logger('silent');
    const service = new ExcelService(checker, logger);

    const handler = new SystemHandlers(service, checker, logger);
    const health = await handler.healthCheck();

    expect(health.success).toBe(true);
    const data = health.data as { status: string; server: unknown; dependencies: unknown };
    expect(data.status).toBeDefined();
    expect(data.dependencies).toBeDefined();
  });
});
