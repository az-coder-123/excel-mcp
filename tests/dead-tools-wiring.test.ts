/**
 * End-to-end tests for the 24 previously-undispatched tools
 * (Phase 3 directive: dispatch them instead of deleting them).
 *
 * Every test drives the REAL ToolHandler.executeTool pipeline (validation +
 * permission + dispatch) against a real in-memory workbook, then inspects the
 * resulting workbook state directly.
 */
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
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

let tmpRoot: string;
let excelService: ExcelService;
let handler: ToolHandler;

beforeAll(() => {
  tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-mcp-dead-'));
  excelService = new ExcelService(
    new PermissionChecker(config({ allowedPaths: [tmpRoot] })),
    new Logger('error')
  );
  handler = new ToolHandler(
    excelService,
    new PermissionChecker(config({ allowedPaths: [tmpRoot] })),
    new Logger('error')
  );
});

afterAll(() => {
  fs.rmSync(tmpRoot, { recursive: true, force: true });
});

/** Creates a workbook through the real tool pipeline and returns its handle. */
const createBook = async (name: string): Promise<string> => {
  const file = path.join(tmpRoot, name);
  const res = await handler.executeTool('excel_create_workbook', {
    filename: file,
    outputPath: file,
  });
  expect(res.success).toBe(true);
  return file;
};

const sheet = (file: string, name = 'Sheet1'): ExcelJS.Worksheet =>
  excelService.activeWorkbooks.get(file)!.getWorksheet(name)!;

describe('CSV transfer tools', () => {
  it('excel_import_csv writes parsed rows (numbers coerced)', async () => {
    const file = await createBook('csv.xlsx');
    const res = await handler.executeTool('excel_import_csv', {
      filename: file,
      worksheet: 'Sheet1',
      csvContent: 'h1,h2\n1,2',
      startCell: 'A1',
    });
    expect(res.success).toBe(true);
    expect(res.data).toEqual({ rowsImported: 2 });
    expect(sheet(file).getCell('A2').value).toBe(1);
  });

  it('excel_export_csv round-trips a range', async () => {
    const file = await createBook('csv2.xlsx');
    await handler.executeTool('excel_import_csv', {
      filename: file, worksheet: 'Sheet1', csvContent: 'h1,h2\n1,2', startCell: 'A1',
    });
    const res = await handler.executeTool('excel_export_csv', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'B2',
    });
    expect(res.success).toBe(true);
    expect(res.data).toEqual({ csv: 'h1,h2\n1,2' });
  });
});

describe('Data cleanup tools', () => {
  it('excel_remove_duplicates keeps first occurrences and reports counts', async () => {
    const file = await createBook('dup.xlsx');
    const ws = sheet(file);
    ws.getCell('A1').value = 'x'; ws.getCell('B1').value = '1';
    ws.getCell('A2').value = 'y'; ws.getCell('B2').value = '2';
    ws.getCell('A3').value = 'x'; ws.getCell('B3').value = '1';

    const res = await handler.executeTool('excel_remove_duplicates', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'B3',
    });
    expect(res.success).toBe(true);
    expect(res.data).toEqual({ removed: 1, remaining: 2 });
    expect(ws.getCell('A3').value).toBeNull();
  });

  it('excel_text_to_columns splits one cell across columns', async () => {
    const file = await createBook('ttc.xlsx');
    sheet(file).getCell('A1').value = 'a-b-c';
    const res = await handler.executeTool('excel_text_to_columns', {
      filename: file, worksheet: 'Sheet1',
      sourceCell: 'A1', targetCell: 'B1', delimiter: '-', numberOfColumns: 3,
    });
    expect(res.success).toBe(true);
    expect(sheet(file).getCell('B1').value).toBe('a');
    expect(sheet(file).getCell('D1').value).toBe('c');
  });

  it('excel_flash_fill infers first-token transform and fills targets', async () => {
    const file = await createBook('flash.xlsx');
    const ws = sheet(file);
    ws.getCell('B1').value = 'John Smith'; ws.getCell('C1').value = 'John';
    ws.getCell('B2').value = 'Jane Doe';  ws.getCell('C2').value = 'Jane';
    ws.getCell('B3').value = 'Bob Jones';

    const res = await handler.executeTool('excel_flash_fill', {
      filename: file, worksheet: 'Sheet1', sourceRange: 'B1:C3', targetRange: 'D1:D3',
    });
    expect(res.success).toBe(true);
    expect(res.data).toEqual({ filled: 3, transform: 'first-token' });
    expect(ws.getCell('D3').value).toBe('Bob');
  });
});

describe('Lookup formula writers', () => {
  it('excel_vlookup stores a VLOOKUP formula cell', async () => {
    const file = await createBook('vl.xlsx');
    const res = await handler.executeTool('excel_vlookup', {
      filename: file, worksheet: 'Sheet1', targetCell: 'D5',
      lookupValue: 'Alice', tableArray: 'A1:B10', colIndex: 2, rangeLookup: false,
    });
    expect(res.success).toBe(true);
    expect((res.data as { formula?: string }).formula).toBe('VLOOKUP("Alice",A1:B10,2,false)');
    expect(sheet(file).getCell('D5').value).toEqual({ formula: 'VLOOKUP("Alice",A1:B10,2,false)' });
  });

  it('excel_index_match stores an INDEX/MATCH formula cell', async () => {
    const file = await createBook('im.xlsx');
    const res = await handler.executeTool('excel_index_match', {
      filename: file, worksheet: 'Sheet1', targetCell: 'E1',
      returnRange: 'B1:B10', lookupRange: 'A1:A10', lookupValue: 42,
    });
    expect(res.success).toBe(true);
    expect((res.data as { formula?: string }).formula).toBe('INDEX(B1:B10,MATCH(42,A1:A10,0))');
  });
});

describe('Pivot summary tool', () => {
  it('excel_create_pivot_table groups and sums by the row field', async () => {
    const file = await createBook('piv.xlsx');
    const wb = excelService.activeWorkbooks.get(file)!;
    const data = wb.addWorksheet('Data');
    data.getCell('A1').value = 'Region'; data.getCell('B1').value = 'Amount';
    data.getCell('A2').value = 'North'; data.getCell('B2').value = 10;
    data.getCell('A3').value = 'South'; data.getCell('B3').value = 20;
    data.getCell('A4').value = 'North'; data.getCell('B4').value = 5;

    const res = await handler.executeTool('excel_create_pivot_table', {
      filename: file, sourceWorksheet: 'Data', sourceStartCell: 'A1', sourceEndCell: 'B4',
      targetWorksheet: 'Pivot', targetCell: 'A1', rowFields: ['Region'], valueFields: ['Amount'],
    });
    expect(res.success).toBe(true);
    expect(res.data).toEqual({ rows: 2, valueFields: ['Amount'] });
    const pivot = wb.getWorksheet('Pivot')!;
    expect(pivot.getCell('A2').value).toBe('North');
    expect(pivot.getCell('B2').value).toBe(15);
    expect(pivot.getCell('B3').value).toBe(20);
  });
});

describe('Gradient & conditional formatting tools', () => {
  it('excel_set_gradient_fill applies a two-stop gradient', async () => {
    const file = await createBook('grad.xlsx');
    const res = await handler.executeTool('excel_set_gradient_fill', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'B1',
      color1: 'FF0000', color2: '0000FF', type: 'vertical',
    });
    expect(res.success).toBe(true);
    const fill = sheet(file).getCell('A1').fill as ExcelJS.FillGradientAngle;
    expect(fill.type).toBe('gradient');
    expect(fill.stops).toHaveLength(2);
  });

  it('excel_add_conditional_format + excel_remove_conditional_format manage rules', async () => {
    const file = await createBook('cf.xlsx');
    const add = await handler.executeTool('excel_add_conditional_format', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'A5',
      ruleType: 'cellValue', operator: 'greaterThan', formula1: '100',
      format: { fill: 'FFC7CE' },
    });
    expect(add.success).toBe(true);

    const ws = sheet(file) as unknown as { conditionalFormattings: Array<{ ref: string }> };
    expect(ws.conditionalFormattings.some((cf) => cf.ref === 'A1:A5')).toBe(true);

    const remove = await handler.executeTool('excel_remove_conditional_format', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'A5',
    });
    expect(remove.success).toBe(true);
    expect(remove.data).toEqual({ removed: 1 });
    expect(ws.conditionalFormattings.some((cf) => cf.ref === 'A1:A5')).toBe(false);
  });

  it('excel_add_data_bar / color_scale / icon_set add visual rules', async () => {
    const file = await createBook('vis.xlsx');
    expect((await handler.executeTool('excel_add_data_bar', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'A5', color: '638EC6',
    })).success).toBe(true);
    expect((await handler.executeTool('excel_add_color_scale', {
      filename: file, worksheet: 'Sheet1', startCell: 'B1', endCell: 'B5',
      minColor: 'F8696B', maxColor: '63BE7B',
    })).success).toBe(true);
    expect((await handler.executeTool('excel_add_icon_set', {
      filename: file, worksheet: 'Sheet1', startCell: 'C1', endCell: 'C5', iconSet: '3Arrows',
    })).success).toBe(true);

    const ws = sheet(file) as unknown as { conditionalFormattings: Array<{ ref: string }> };
    expect(ws.conditionalFormattings.length).toBe(3);
  });
});

describe('Protection tools', () => {
  it('excel_protect_worksheet protects and excel_unprotect_worksheet reverses', async () => {
    const file = await createBook('prot.xlsx');
    const protect = await handler.executeTool('excel_protect_worksheet', {
      filename: file, worksheet: 'Sheet1', password: 'secret',
    });
    expect(protect.success).toBe(true);
    const prot = sheet(file) as unknown as { sheetProtection?: unknown };
    expect(prot.sheetProtection).toBeDefined();

    const unprotect = await handler.executeTool('excel_unprotect_worksheet', {
      filename: file, worksheet: 'Sheet1', password: 'secret',
    });
    expect(unprotect.success).toBe(true);
  });

  it('excel_protect_cells sets the locked flag on a range', async () => {
    const file = await createBook('lock.xlsx');
    const res = await handler.executeTool('excel_protect_cells', {
      filename: file, worksheet: 'Sheet1', startCell: 'A1', endCell: 'A3', locked: true,
    });
    expect(res.success).toBe(true);
    expect(sheet(file).getCell('A2').protection?.locked).toBe(true);
  });

  it('excel_protect_workbook returns an honest engine-limitation error', async () => {
    const file = await createBook('pwb.xlsx');
    const res = await handler.executeTool('excel_protect_workbook', {
      filename: file, password: 'x',
    });
    expect(res.success).toBe(false);
    expect(res.error).toMatch(/not supported by the ExcelJS engine/i);
  });
});

describe('Chart tools', () => {
  it('excel_add_chart + excel_list_charts register and report charts', async () => {
    const file = await createBook('chart.xlsx');
    const ws = sheet(file);
    ws.getCell('A1').value = 'Q1'; ws.getCell('B1').value = 10;
    ws.getCell('A2').value = 'Q2'; ws.getCell('B2').value = 20;

    const add = await handler.executeTool('excel_add_chart', {
      filename: file, worksheet: 'Sheet1', dataStartCell: 'A1', dataEndCell: 'B2',
      chartType: 'column', targetCell: 'E1', title: 'Sales',
    });
    expect(add.success).toBe(true);

    const list = await handler.executeTool('excel_list_charts', {
      filename: file, worksheet: 'Sheet1',
    });
    expect(list.success).toBe(true);
    expect(list.data).toEqual([{ index: 0, type: 'column', title: 'Sales' }]);
  });
});

describe('Dispatch alias', () => {
  it('excel_check_circular answers through the formula analyzer', async () => {
    const file = await createBook('circ.xlsx');
    const res = await handler.executeTool('excel_check_circular', {
      filename: file, worksheet: 'Sheet1',
    });
    expect(res.success).toBe(true);
  });
});
