/**
 * Workbook lifecycle hardening tests (Phase 3 C3 + audit Finding 6.2)
 *
 * Covers:
 * - LRU eviction at MAX_ACTIVE_WORKBOOKS = 10 (Finding 1.2 memory leak)
 * - touch-on-access keeps recently used workbooks alive
 * - saveWorkbook() without outputPath returns to the ORIGINAL source path (Finding 1.3)
 * - exportWorksheetToNewFile preserves styles, number formats and merges (Finding 6.2)
 */
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { ExcelWorkbookManager } from '../src/services/excel-workbook-manager.js';
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

let tmpRoot: string;
let activeWorkbooks: Map<string, ExcelJS.Workbook>;
let manager: ExcelWorkbookManager;

beforeAll(() => {
  tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-mcp-lru-'));
  activeWorkbooks = new Map();
  manager = new ExcelWorkbookManager(
    new PermissionChecker(config({ allowedPaths: [tmpRoot] })),
    new Logger('silent'),
    activeWorkbooks
  );
});

afterAll(() => {
  fs.rmSync(tmpRoot, { recursive: true, force: true });
});

describe('LRU eviction — audit Finding 1.2 (unbounded Map growth)', () => {
  it('evicts the OLDEST workbook when exceeding the cap of 10', async () => {
    for (let i = 1; i <= 10; i++) {
      await manager.createWorkbook(path.join(tmpRoot, `lru-a-${i}.xlsx`));
    }
    expect(activeWorkbooks.size).toBe(10);

    // The 11th registration must evict the least-recently-used entry (lru-a-1)
    await manager.createWorkbook(path.join(tmpRoot, 'lru-a-11.xlsx'));
    expect(activeWorkbooks.size).toBe(10);

    const evicted = await manager.saveWorkbook(path.join(tmpRoot, 'lru-a-1.xlsx'));
    expect(evicted.success).toBe(false);
    expect(evicted.error).toMatch(/not opened/i);

    const survivor = await manager.saveWorkbook(path.join(tmpRoot, 'lru-a-11.xlsx'));
    expect(survivor.success).toBe(true);
  });

  it('keeps touched workbooks alive and evicts the next-oldest instead', async () => {
    for (let i = 1; i <= 10; i++) {
      await manager.createWorkbook(path.join(tmpRoot, `lru-b-${i}.xlsx`));
    }
    // Touch lru-b-1 by opening it (moves it to most-recently-used)
    const touched = await manager.openWorkbook(path.join(tmpRoot, 'lru-b-1.xlsx'));
    expect(touched.success).toBe(true);

    await manager.createWorkbook(path.join(tmpRoot, 'lru-b-11.xlsx'));
    expect(activeWorkbooks.size).toBe(10);

    // lru-b-1 survives (recently touched); lru-b-2 is now the oldest → evicted
    const survivor = await manager.saveWorkbook(path.join(tmpRoot, 'lru-b-1.xlsx'));
    expect(survivor.success).toBe(true);
    const evicted = await manager.saveWorkbook(path.join(tmpRoot, 'lru-b-2.xlsx'));
    expect(evicted.success).toBe(false);
    expect(evicted.error).toMatch(/not opened/i);
  });
});

describe('saveWorkbook without outputPath — audit Finding 1.3 (lost working paths)', () => {
  it('saves back to the ORIGINAL source path instead of process.cwd()', async () => {
    const file = path.join(tmpRoot, 'source-save.xlsx');
    await manager.createWorkbook(file);

    const wb = activeWorkbooks.get(file);
    expect(wb).toBeDefined();
    wb!.getWorksheet('Sheet1')!.getCell('A1').value = 99;

    // No outputPath provided — must write to the registered source path
    const res = await manager.saveWorkbook(file);
    expect(res.success).toBe(true);

    const verify = new ExcelJS.Workbook();
    await verify.xlsx.readFile(file);
    expect(verify.getWorksheet('Sheet1')!.getCell('A1').value).toBe(99);
  });
});

describe('exportWorksheetToNewFile fidelity — audit Finding 6.2 (format stripping)', () => {
  it('preserves values, fonts, number formats, merges, and column widths', async () => {
    const file = path.join(tmpRoot, 'fancy.xlsx');
    await manager.createWorkbook(file);

    const wb = activeWorkbooks.get(file)!;
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = 1234.5;
    ws.getCell('A1').font = { bold: true };
    ws.getCell('A1').numFmt = '#,##0.00';
    ws.mergeCells('A1:B1');
    ws.getColumn(1).width = 42;

    const target = path.join(tmpRoot, 'fancy-export.xlsx');
    const res = await manager.exportWorksheetToNewFile(file, 'Sheet1', target);
    expect(res.success).toBe(true);

    const verify = new ExcelJS.Workbook();
    await verify.xlsx.readFile(target);
    const out = verify.getWorksheet('Sheet1')!;
    expect(out.getCell('A1').value).toBe(1234.5);
    expect(out.getCell('A1').font?.bold).toBe(true);
    expect(out.getCell('A1').numFmt).toBe('#,##0.00');
    expect(out.model.merges).toContain('A1:B1');
    expect(out.getColumn(1).width).toBe(42);
  });

  it('exports ONLY the requested worksheet, dropping the others', async () => {
    const file = path.join(tmpRoot, 'multi.xlsx');
    await manager.createWorkbook(file);
    const wb = activeWorkbooks.get(file)!;
    wb.addWorksheet('Extra');
    wb.getWorksheet('Sheet1')!.getCell('A1').value = 'keep';
    wb.getWorksheet('Extra')!.getCell('A1').value = 'drop';

    const target = path.join(tmpRoot, 'multi-export.xlsx');
    const res = await manager.exportWorksheetToNewFile(file, 'Sheet1', target);
    expect(res.success).toBe(true);

    const verify = new ExcelJS.Workbook();
    await verify.xlsx.readFile(target);
    expect(verify.worksheets.map((w) => w.name)).toEqual(['Sheet1']);
    expect(verify.getWorksheet('Sheet1')!.getCell('A1').value).toBe('keep');
  });
});
