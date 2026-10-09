/**
 * Workbook lifecycle + security enforcement tests for ExcelWorkbookManager
 *
 * Regression coverage for audit Finding 5.3 (A4 remediation):
 * every write-target operation must run the FULL validateFileAccess pipeline
 * (permission + path containment + extension whitelist) — including the
 * previously completely unvalidated exportWorksheetToNewFile.
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
  tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-mcp-wb-'));
  const strictChecker = new PermissionChecker(config({ allowedPaths: [tmpRoot] }));
  activeWorkbooks = new Map();
  manager = new ExcelWorkbookManager(strictChecker, new Logger('error'), activeWorkbooks);
});

afterAll(() => {
  fs.rmSync(tmpRoot, { recursive: true, force: true });
});

describe('createWorkbook — A4: full validateFileAccess on write targets', () => {
  it('creates a workbook inside the allowed root and writes it to disk', async () => {
    const file = path.join(tmpRoot, 'created.xlsx');
    const res = await manager.createWorkbook(file);
    expect(res.success).toBe(true);
    expect(fs.existsSync(file)).toBe(true);
    expect(res.data?.worksheetCount).toBe(1);
    expect(res.data?.worksheets[0]?.name).toBe('Sheet1');
  });

  it('REJECTS creation outside allowed paths (.. traversal)', async () => {
    const escape = path.join(tmpRoot, '..', 'escape-a4-create.xlsx');
    const res = await manager.createWorkbook(escape);
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/does not match/i);
    expect(fs.existsSync(path.resolve(escape))).toBe(false);
  });

  it('REJECTS non-whitelisted extensions (.txt)', async () => {
    const res = await manager.createWorkbook(path.join(tmpRoot, 'bad.txt'));
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/extension/i);
    expect(fs.existsSync(path.join(tmpRoot, 'bad.txt'))).toBe(false);
  });

  it('REJECTS when the caller lacks write permission', async () => {
    const readOnlyChecker = new PermissionChecker(
      config({ allowedPaths: [tmpRoot], permissions: ['read'] })
    );
    const readOnlyManager = new ExcelWorkbookManager(readOnlyChecker, new Logger('error'), new Map());
    const res = await readOnlyManager.createWorkbook(path.join(tmpRoot, 'denied.xlsx'));
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/permission/i);
  });
});

describe('saveWorkbook — A4: validation on the save target', () => {
  it('saves to an allowed output path', async () => {
    const file = path.join(tmpRoot, 'save-src.xlsx');
    await manager.createWorkbook(file);
    const out = path.join(tmpRoot, 'save-out.xlsx');
    const res = await manager.saveWorkbook(file, out);
    expect(res.success).toBe(true);
    expect(fs.existsSync(out)).toBe(true);
  });

  it('REJECTS saving outside allowed paths', async () => {
    const file = path.join(tmpRoot, 'save-src2.xlsx');
    await manager.createWorkbook(file);
    const res = await manager.saveWorkbook(file, path.join(tmpRoot, '..', 'escape-save.xlsx'));
    expect(res.success).toBe(false);
    expect(fs.existsSync(path.join(tmpRoot, '..', 'escape-save.xlsx'))).toBe(false);
  });

  it('REJECTS saving to a non-whitelisted extension', async () => {
    const file = path.join(tmpRoot, 'save-src3.xlsx');
    await manager.createWorkbook(file);
    const res = await manager.saveWorkbook(file, path.join(tmpRoot, 'save-out.exe'));
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/extension/i);
  });

  it('fails for a workbook that was never opened', async () => {
    const res = await manager.saveWorkbook(path.join(tmpRoot, 'ghost.xlsx'));
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/not opened/i);
  });
});

describe('exportWorksheetToNewFile — A4: previously completely unvalidated', () => {
  it('REJECTS export to a non-whitelisted extension (.txt)', async () => {
    const file = path.join(tmpRoot, 'exp-ext.xlsx');
    await manager.createWorkbook(file);
    const target = path.join(tmpRoot, 'exp-out.txt');
    const res = await manager.exportWorksheetToNewFile(file, 'Sheet1', target);
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/extension/i);
    expect(fs.existsSync(target)).toBe(false);
  });

  it('REJECTS export outside allowed paths', async () => {
    const file = path.join(tmpRoot, 'exp-path.xlsx');
    await manager.createWorkbook(file);
    const res = await manager.exportWorksheetToNewFile(
      file,
      'Sheet1',
      path.join(tmpRoot, '..', 'escape-export.xlsx')
    );
    expect(res.success).toBe(false);
    expect(fs.existsSync(path.join(tmpRoot, '..', 'escape-export.xlsx'))).toBe(false);
  });

  it('REJECTS export when the caller lacks write permission', async () => {
    const readOnlyChecker = new PermissionChecker(
      config({ allowedPaths: [tmpRoot], permissions: ['read'] })
    );
    const roManager = new ExcelWorkbookManager(readOnlyChecker, new Logger('error'), new Map());
    const res = await roManager.exportWorksheetToNewFile(
      'irrelevant.xlsx',
      'Sheet1',
      path.join(tmpRoot, 'ro-export.xlsx')
    );
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/permission/i);
  });

  it('exports worksheet values to an allowed target', async () => {
    const file = path.join(tmpRoot, 'exp-ok.xlsx');
    await manager.createWorkbook(file);
    const wb = activeWorkbooks.get(file);
    expect(wb).toBeDefined();
    wb!.getWorksheet('Sheet1')!.getCell('A1').value = 42;

    const target = path.join(tmpRoot, 'exported.xlsx');
    const res = await manager.exportWorksheetToNewFile(file, 'Sheet1', target);
    expect(res.success).toBe(true);
    expect(fs.existsSync(target)).toBe(true);

    const verify = new ExcelJS.Workbook();
    await verify.xlsx.readFile(target);
    expect(verify.getWorksheet('Sheet1')!.getCell('A1').value).toBe(42);
  });

  it('fails when the source workbook is not open', async () => {
    const res = await manager.exportWorksheetToNewFile(
      'ghost.xlsx',
      'Sheet1',
      path.join(tmpRoot, 'ghost-export.xlsx')
    );
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/not opened/i);
  });

  it('fails when the source worksheet does not exist', async () => {
    const file = path.join(tmpRoot, 'exp-ws.xlsx');
    await manager.createWorkbook(file);
    const res = await manager.exportWorksheetToNewFile(file, 'Missing', path.join(tmpRoot, 'ws-export.xlsx'));
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/not found/i);
  });
});

describe('openWorkbook / closeWorkbook — lifecycle', () => {
  it('opens a real file inside the allowed root and reports its sheets', async () => {
    const file = path.join(tmpRoot, 'opened.xlsx');
    await manager.createWorkbook(file);
    const res = await manager.openWorkbook(file);
    expect(res.success).toBe(true);
    expect(res.data?.filename).toBe('opened.xlsx'); // keyed by basename
    expect(res.data?.worksheetCount).toBeGreaterThanOrEqual(1);
  });

  it('REJECTS opening a path outside allowed paths (A1 integration)', async () => {
    const res = await manager.openWorkbook(path.join(tmpRoot, '..', '..', 'etc', 'passwd.xlsx'));
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/does not match/i);
  });

  it('fails gracefully for a missing file inside the allowed root', async () => {
    const res = await manager.openWorkbook(path.join(tmpRoot, 'does-not-exist.xlsx'));
    expect(res.success).toBe(false);
  });

  it('closes an open workbook and reports failure on double close', async () => {
    const file = path.join(tmpRoot, 'closed.xlsx');
    await manager.createWorkbook(file);
    expect(manager.closeWorkbook(file).success).toBe(true);
    const second = manager.closeWorkbook(file);
    expect(second.success).toBe(false);
    expect(second.error).toMatch(/not found/i);
  });
});
