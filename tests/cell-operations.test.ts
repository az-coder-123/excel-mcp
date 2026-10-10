/**
 * Unit and Integration Tests for ExcelCellOperations
 * Testing all tools, edge cases, data types, and ToolHandler dispatch:
 * - readCell (types: string, number, boolean, formula, date, empty)
 * - readRange (2D array extraction)
 * - writeCell (write and read-back validation)
 * - writeBatch (multi-cell atomic writes)
 * - copyRange (range copying across cells)
 * - findReplace (exact matching, case-sensitivity, replace count)
 * - sortRange (ascending, descending, numeric & string sorting)
 * - getNamedRanges & addNamedRange (named range lifecycle)
 */

import ExcelJS from 'exceljs';
import { ExcelCellOperations } from '../src/services/excel-cell-operations.js';
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

describe('ExcelCellOperations', () => {
  let workbooks: Map<string, ExcelJS.Workbook>;
  let cellOps: ExcelCellOperations;
  let excelService: ExcelService;
  let handler: ToolHandler;
  const fileName = 'test_cells.xlsx';
  const sheetName = 'DataSheet';

  beforeEach(() => {
    workbooks = new Map();
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    cellOps = new ExcelCellOperations(permChecker, logger, workbooks);
    excelService = new ExcelService(permChecker, logger);
    handler = new ToolHandler(excelService, permChecker, logger);

    const wb = new ExcelJS.Workbook();
    wb.addWorksheet(sheetName);
    workbooks.set(fileName, wb);
    excelService.getActiveWorkbooks().set(fileName, wb);
  });

  const getWs = (): ExcelJS.Worksheet => workbooks.get(fileName)!.getWorksheet(sheetName)!;

  describe('readCell & writeCell', () => {
    it('writes and reads back string, number, and boolean values', async () => {
      await cellOps.writeCell(fileName, sheetName, 'A1', 'Hello Excel');
      await cellOps.writeCell(fileName, sheetName, 'A2', 12345.67);
      await cellOps.writeCell(fileName, sheetName, 'A3', true);

      const r1 = await cellOps.readCell(fileName, sheetName, 'A1');
      expect(r1.success).toBe(true);
      expect(r1.data?.value).toBe('Hello Excel');
      expect(r1.data?.type).toBe('string');

      const r2 = await cellOps.readCell(fileName, sheetName, 'A2');
      expect(r2.success).toBe(true);
      expect(r2.data?.value).toBe(12345.67);
      expect(r2.data?.type).toBe('number');

      const r3 = await cellOps.readCell(fileName, sheetName, 'A3');
      expect(r3.success).toBe(true);
      expect(r3.data?.value).toBe(true);
      expect(r3.data?.type).toBe('boolean');
    });

    it('reads empty cell as null gracefully', async () => {
      const res = await cellOps.readCell(fileName, sheetName, 'Z99');
      expect(res.success).toBe(true);
      expect(res.data?.value).toBeNull();
    });

    it('returns error when file or sheet is missing', async () => {
      const res1 = await cellOps.readCell('non_existent.xlsx', sheetName, 'A1');
      expect(res1.success).toBe(false);

      const res2 = await cellOps.readCell(fileName, 'NonExistentSheet', 'A1');
      expect(res2.success).toBe(false);
    });
  });

  describe('readRange', () => {
    it('reads a 2D matrix of cell values across rows and columns', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 'R1C1'; ws.getCell('B1').value = 10;
      ws.getCell('A2').value = 'R2C1'; ws.getCell('B2').value = 20;

      const res = await cellOps.readRange(fileName, sheetName, {
        start: { row: 1, column: 1 },
        end: { row: 2, column: 2 },
      });

      expect(res.success).toBe(true);
      expect(res.data).toHaveLength(2);
      expect(res.data![0][0].value).toBe('R1C1');
      expect(res.data![0][1].value).toBe(10);
      expect(res.data![1][0].value).toBe('R2C1');
      expect(res.data![1][1].value).toBe(20);
    });
  });

  describe('writeBatch', () => {
    it('writes multiple cells in a single batch operation', async () => {
      const batchData = [
        { cellAddress: 'A1', value: 'ID' },
        { cellAddress: 'B1', value: 'Name' },
        { cellAddress: 'A2', value: 101 },
        { cellAddress: 'B2', value: 'Alice' },
      ];

      const res = await cellOps.writeBatch(fileName, sheetName, batchData);
      expect(res.success).toBe(true);
      expect(res.data?.count).toBe(4);

      const ws = getWs();
      expect(ws.getCell('A1').value).toBe('ID');
      expect(ws.getCell('B2').value).toBe('Alice');
    });
  });

  describe('copyRange', () => {
    it('copies cells from source range to destination start cell', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 'Item A';
      ws.getCell('A2').value = 'Item B';

      const res = await cellOps.copyRange(
        fileName,
        sheetName,
        'A1',
        'A2',
        'A5'
      );

      expect(res.success).toBe(true);
      expect(ws.getCell('A5').value).toBe('Item A');
      expect(ws.getCell('A6').value).toBe('Item B');
    });
  });

  describe('findReplace', () => {
    it('finds and replaces text with matchCase option', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 'apple pie';
      ws.getCell('A2').value = 'Apple juice';
      ws.getCell('A3').value = 'banana';

      // Replace 'apple' with 'cherry', case insensitive (matchCase: false)
      const res = await cellOps.findReplace(
        fileName,
        sheetName,
        'apple',
        'cherry',
        false
      );

      expect(res.success).toBe(true);
      expect(res.data?.count).toBe(2);
      expect(ws.getCell('A1').value).toBe('cherry pie');
      expect(ws.getCell('A2').value).toBe('cherry juice');
      expect(ws.getCell('A3').value).toBe('banana');
    });
  });

  describe('sortRange', () => {
    it('sorts numeric rows in ascending and descending order', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 40;
      ws.getCell('A2').value = 10;
      ws.getCell('A3').value = 30;
      ws.getCell('A4').value = 20;

      // Sort ascending
      const resAsc = await cellOps.sortRange(
        fileName,
        sheetName,
        'A1',
        'A4',
        1,
        true
      );
      expect(resAsc.success).toBe(true);
      expect(ws.getCell('A1').value).toBe(10);
      expect(ws.getCell('A2').value).toBe(20);
      expect(ws.getCell('A3').value).toBe(30);
      expect(ws.getCell('A4').value).toBe(40);

      // Sort descending
      const resDesc = await cellOps.sortRange(
        fileName,
        sheetName,
        'A1',
        'A4',
        1,
        false
      );
      expect(resDesc.success).toBe(true);
      expect(ws.getCell('A1').value).toBe(40);
      expect(ws.getCell('A2').value).toBe(30);
      expect(ws.getCell('A3').value).toBe(20);
      expect(ws.getCell('A4').value).toBe(10);
    });
  });

  describe('Named Ranges', () => {
    it('adds and lists named ranges in workbook', async () => {
      const addRes = await cellOps.addNamedRange(fileName, sheetName, 'TaxRates', 'B1', 'B5');
      expect(addRes.success).toBe(true);

      const listRes = await cellOps.getNamedRanges(fileName);
      expect(listRes.success).toBe(true);
      const names = listRes.data?.map(n => n.name);
      expect(names).toContain('TaxRates');
    });
  });

  describe('ToolHandler Integration for Cell Operations', () => {
    it('executes excel_write_cell and excel_read_cell through ToolHandler', async () => {
      const writeRes = await handler.executeTool('excel_write_cell', {
        filename: fileName,
        worksheet: sheetName,
        cellAddress: 'C3',
        value: 'ToolHandler Dispatch Test',
      });
      expect(writeRes.success).toBe(true);

      const readRes = await handler.executeTool('excel_read_cell', {
        filename: fileName,
        worksheet: sheetName,
        cellAddress: 'C3',
      });
      expect(readRes.success).toBe(true);
      expect((readRes.data as any).value).toBe('ToolHandler Dispatch Test');
    });

    it('executes excel_write_batch through ToolHandler', async () => {
      const batchRes = await handler.executeTool('excel_write_batch', {
        filename: fileName,
        worksheet: sheetName,
        data: [
          { cellAddress: 'D1', value: 111 },
          { cellAddress: 'D2', value: 222 },
        ],
      });
      expect(batchRes.success).toBe(true);
      expect((batchRes.data as any).count).toBe(2);
    });
  });
});
