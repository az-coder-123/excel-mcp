/**
 * Unit and Integration Tests for ExcelStructureOperations & WorksheetHandlers
 * Testing all worksheet, structure manipulation tools, edge cases, and ToolHandler dispatch:
 * - listWorksheets (metadata, row/col counts, hidden status)
 * - addWorksheet (valid creation, duplicate prevention, empty name check)
 * - deleteWorksheet (valid deletion, only sheet protection, missing sheet check)
 * - renameWorksheet (valid rename, collision prevention, missing source check)
 * - copyWorksheet (content cloning across sheets, duplicate target check)
 * - insertRows & deleteRows (row splicing, count/row boundary validation)
 * - insertColumns & deleteColumns (column splicing, case-insensitive letters, boundary validation)
 * - mergeCells & unmergeCells (range merging, unmerging, invalid range coordinate check)
 * - addTable (table creation, headerRow, styling, range validation)
 * - addFilter & removeFilter (autoFilter assignment and clearance)
 * - calculateFormula (stub execution)
 * - protectWorksheet, unprotectWorksheet, protectCells (sheet protection lifecycle and cell locking)
 */

import ExcelJS from 'exceljs';
import { ExcelStructureOperations } from '../src/services/excel-structure-operations.js';
import { ExcelService } from '../src/services/excel-service.js';
import { ToolHandler } from '../src/tools/tool-handler.js';
import { PermissionChecker } from '../src/security/permission-checker.js';
import { Logger } from '../src/utils/logger.js';
import { PermissionConfig, WorksheetInfo } from '../src/types/index.js';

const config = (overrides: Partial<PermissionConfig> = {}): PermissionConfig => ({
  allowedPaths: [],
  deniedPaths: [],
  maxFileSize: 50 * 1024 * 1024,
  allowedExtensions: ['.xlsx', '.xls', '.xlsm', '.xlsb'],
  permissions: ['read', 'write', 'delete'],
  ...overrides,
});

describe('ExcelStructureOperations & Worksheet Tools', () => {
  let workbooks: Map<string, ExcelJS.Workbook>;
  let structureOps: ExcelStructureOperations;
  let excelService: ExcelService;
  let handler: ToolHandler;
  const fileName = 'test_structure.xlsx';

  beforeEach(() => {
    workbooks = new Map();
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    structureOps = new ExcelStructureOperations(permChecker, logger, workbooks);
    excelService = new ExcelService(permChecker, logger);
    handler = new ToolHandler(excelService, permChecker, logger);

    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1');
    workbooks.set(fileName, wb);
    excelService.getActiveWorkbooks().set(fileName, wb);
  });

  const getWb = (): ExcelJS.Workbook => workbooks.get(fileName)!;

  describe('Worksheet Lifecycle (list, add, rename, copy, delete)', () => {
    it('lists worksheets with metadata', async () => {
      const res = structureOps.getWorksheets(fileName);
      expect(res.success).toBe(true);
      expect(res.data).toHaveLength(1);
      const sheet = (res.data as WorksheetInfo[])[0];
      expect(sheet.name).toBe('Sheet1');
      expect(sheet.index).toBe(0);
      expect(sheet.hidden).toBe(false);
    });

    it('adds a new worksheet successfully', async () => {
      const res = await structureOps.addWorksheet(fileName, 'Sheet2');
      expect(res.success).toBe(true);
      expect(res.data?.name).toBe('Sheet2');
      expect(getWb().worksheets).toHaveLength(2);
    });

    it('rejects adding duplicate or empty worksheet name', async () => {
      const emptyRes = await structureOps.addWorksheet(fileName, '  ');
      expect(emptyRes.success).toBe(false);
      expect(emptyRes.error).toMatch(/cannot be empty/i);

      const dupRes = await structureOps.addWorksheet(fileName, 'Sheet1');
      expect(dupRes.success).toBe(false);
      expect(dupRes.error).toMatch(/already exists/i);
    });

    it('renames a worksheet successfully', async () => {
      const res = await structureOps.renameWorksheet(fileName, 'Sheet1', 'RenamedSheet');
      expect(res.success).toBe(true);
      expect(getWb().getWorksheet('RenamedSheet')).toBeDefined();
      expect(getWb().getWorksheet('Sheet1')).toBeUndefined();
    });

    it('rejects renaming to existing name or empty name', async () => {
      await structureOps.addWorksheet(fileName, 'Sheet2');
      const dupRes = await structureOps.renameWorksheet(fileName, 'Sheet1', 'Sheet2');
      expect(dupRes.success).toBe(false);
      expect(dupRes.error).toMatch(/already exists/i);

      const emptyRes = await structureOps.renameWorksheet(fileName, 'Sheet1', '');
      expect(emptyRes.success).toBe(false);
      expect(emptyRes.error).toMatch(/cannot be empty/i);

      const missingRes = await structureOps.renameWorksheet(fileName, 'NonExistent', 'NewName');
      expect(missingRes.success).toBe(false);
      expect(missingRes.error).toMatch(/not found/i);
    });

    it('copies a worksheet with cell data preserved', async () => {
      const ws1 = getWb().getWorksheet('Sheet1')!;
      ws1.getCell('A1').value = 'Original Data';
      ws1.getCell('B2').value = 999;

      const res = await structureOps.copyWorksheet(fileName, 'Sheet1', 'Sheet1_Copy');
      expect(res.success).toBe(true);

      const copyWs = getWb().getWorksheet('Sheet1_Copy')!;
      expect(copyWs).toBeDefined();
      expect(copyWs.getCell('A1').value).toBe('Original Data');
      expect(copyWs.getCell('B2').value).toBe(999);
    });

    it('rejects copying with invalid names or missing source', async () => {
      const missingRes = await structureOps.copyWorksheet(fileName, 'NoSheet', 'Copy');
      expect(missingRes.success).toBe(false);
      expect(missingRes.error).toMatch(/not found/i);

      const emptyRes = await structureOps.copyWorksheet(fileName, 'Sheet1', '');
      expect(emptyRes.success).toBe(false);
      expect(emptyRes.error).toMatch(/cannot be empty/i);

      const dupRes = await structureOps.copyWorksheet(fileName, 'Sheet1', 'Sheet1');
      expect(dupRes.success).toBe(false);
      expect(dupRes.error).toMatch(/already exists/i);
    });

    it('deletes a worksheet when multiple exist', async () => {
      await structureOps.addWorksheet(fileName, 'Sheet2');
      expect(getWb().worksheets).toHaveLength(2);

      const res = await structureOps.deleteWorksheet(fileName, 'Sheet1');
      expect(res.success).toBe(true);
      expect(getWb().worksheets).toHaveLength(1);
      expect(getWb().getWorksheet('Sheet1')).toBeUndefined();
    });

    it('protects against deleting the only worksheet in workbook', async () => {
      expect(getWb().worksheets).toHaveLength(1);
      const res = await structureOps.deleteWorksheet(fileName, 'Sheet1');
      expect(res.success).toBe(false);
      expect(res.error).toMatch(/only worksheet/i);
    });

    it('returns error if worksheet to delete is not found', async () => {
      await structureOps.addWorksheet(fileName, 'Sheet2');
      const res = await structureOps.deleteWorksheet(fileName, 'NoSuchSheet');
      expect(res.success).toBe(false);
      expect(res.error).toMatch(/not found/i);
    });
  });

  describe('Row and Column Manipulations', () => {
    it('inserts rows into worksheet', async () => {
      const ws = getWb().getWorksheet('Sheet1')!;
      ws.getCell('A1').value = 'Row 1';
      ws.getCell('A2').value = 'Row 2';

      const res = await structureOps.insertRows(fileName, 'Sheet1', 2, 2);
      expect(res.success).toBe(true);
    });

    it('rejects invalid startRow or count in insertRows', async () => {
      const zeroRowRes = await structureOps.insertRows(fileName, 'Sheet1', 0, 1);
      expect(zeroRowRes.success).toBe(false);
      expect(zeroRowRes.error).toMatch(/must be positive/i);

      const zeroCountRes = await structureOps.insertRows(fileName, 'Sheet1', 1, 0);
      expect(zeroCountRes.success).toBe(false);
      expect(zeroCountRes.error).toMatch(/greater than 0/i);
    });

    it('deletes rows from worksheet', async () => {
      const ws = getWb().getWorksheet('Sheet1')!;
      ws.getCell('A1').value = 'Row 1';
      ws.getCell('A2').value = 'Row 2';

      const res = await structureOps.deleteRows(fileName, 'Sheet1', 1, 1);
      expect(res.success).toBe(true);
    });

    it('rejects invalid parameters in deleteRows', async () => {
      const invalidRow = await structureOps.deleteRows(fileName, 'Sheet1', -1, 1);
      expect(invalidRow.success).toBe(false);

      const invalidCount = await structureOps.deleteRows(fileName, 'Sheet1', 1, -1);
      expect(invalidCount.success).toBe(false);
    });

    it('inserts and deletes columns with case-insensitive column letters', async () => {
      const ws = getWb().getWorksheet('Sheet1')!;
      ws.getCell('A1').value = 'Col A';
      ws.getCell('B1').value = 'Col B';

      // Lowercase column letter 'b'
      const insertRes = await structureOps.insertColumns(fileName, 'Sheet1', 'b', 1);
      expect(insertRes.success).toBe(true);

      const deleteRes = await structureOps.deleteColumns(fileName, 'Sheet1', 'B', 1);
      expect(deleteRes.success).toBe(true);
    });

    it('rejects invalid column letters or non-positive count', async () => {
      const badCol = await structureOps.insertColumns(fileName, 'Sheet1', '123', 1);
      expect(badCol.success).toBe(false);
      expect(badCol.error).toMatch(/invalid column letter/i);

      const badCount = await structureOps.deleteColumns(fileName, 'Sheet1', 'A', 0);
      expect(badCount.success).toBe(false);
      expect(badCount.error).toMatch(/greater than 0/i);
    });
  });

  describe('Cell Merging, Tables, and Filters', () => {
    it('merges and unmerges a cell range', async () => {
      const mergeRes = await structureOps.mergeCells(fileName, 'Sheet1', 'A1', 'B2');
      expect(mergeRes.success).toBe(true);

      const ws = getWb().getWorksheet('Sheet1')!;
      expect(ws.getCell('A1').isMerged).toBe(true);

      const unmergeRes = await structureOps.unmergeCells(fileName, 'Sheet1', 'A1');
      expect(unmergeRes.success).toBe(true);
      expect(ws.getCell('A1').isMerged).toBe(false);
    });

    it('rejects merge with invalid range', async () => {
      const res = await structureOps.mergeCells(fileName, 'Sheet1', 'BAD', 'COORDS');
      expect(res.success).toBe(false);
      expect(res.error).toMatch(/invalid cell range/i);
    });

    it('adds a table to worksheet', async () => {
      const ws = getWb().getWorksheet('Sheet1')!;
      ws.getCell('A1').value = 'Name';
      ws.getCell('B1').value = 'Age';
      ws.getCell('A2').value = 'Alice';
      ws.getCell('B2').value = 30;

      const res = await structureOps.addTable(fileName, 'Sheet1', 'A1', 'B2', 'PeopleTable', 'TableStyleMedium9');
      expect(res.success).toBe(true);
    });

    it('rejects table creation with empty name or invalid range', async () => {
      const badName = await structureOps.addTable(fileName, 'Sheet1', 'A1', 'B2', '');
      expect(badName.success).toBe(false);
      expect(badName.error).toMatch(/cannot be empty/i);

      const badRange = await structureOps.addTable(fileName, 'Sheet1', 'INVALID', 'RANGE', 'Test');
      expect(badRange.success).toBe(false);
      expect(badRange.error).toMatch(/invalid cell range/i);
    });

    it('adds and removes autoFilter', async () => {
      const addRes = await structureOps.addFilter(fileName, 'Sheet1', 'A1', 'C10');
      expect(addRes.success).toBe(true);
      const ws = getWb().getWorksheet('Sheet1')!;
      expect(ws.autoFilter).toBe('A1:C10');

      const removeRes = await structureOps.removeFilter(fileName, 'Sheet1');
      expect(removeRes.success).toBe(true);
      expect(ws.autoFilter).toBeUndefined();
    });

    it('rejects autoFilter with invalid range', async () => {
      const badFilter = await structureOps.addFilter(fileName, 'Sheet1', '???', '!!!');
      expect(badFilter.success).toBe(false);
      expect(badFilter.error).toMatch(/invalid cell range/i);
    });

    it('executes calculateFormula cleanly', async () => {
      const res = await structureOps.calculateFormula(fileName);
      expect(res.success).toBe(true);
    });
  });

  describe('Worksheet and Cell Protection', () => {
    it('protects worksheet, locks cells, and unprotects worksheet', async () => {
      const lockRes = await structureOps.protectCells(fileName, 'Sheet1', 'A1', 'B5', true);
      expect(lockRes.success).toBe(true);
      const ws = getWb().getWorksheet('Sheet1')!;
      expect(ws.getCell('A1').protection?.locked).toBe(true);

      const protRes = await structureOps.protectWorksheet(fileName, 'Sheet1', 'pass123');
      expect(protRes.success).toBe(true);

      const unprotRes = await structureOps.unprotectWorksheet(fileName, 'Sheet1', 'pass123');
      expect(unprotRes.success).toBe(true);
    });

    it('rejects protectCells on invalid range', async () => {
      const res = await structureOps.protectCells(fileName, 'Sheet1', 'XYZ', '123', true);
      expect(res.success).toBe(false);
      expect(res.error).toMatch(/invalid cell range/i);
    });

    it('protectWorkbook returns transparent engine limitation error', async () => {
      const pRes = await structureOps.protectWorkbook();
      expect(pRes.success).toBe(false);
      expect(pRes.error).toMatch(/not supported by the ExcelJS engine/i);

      const uRes = await structureOps.unprotectWorkbook();
      expect(uRes.success).toBe(false);
      expect(uRes.error).toMatch(/not supported by the ExcelJS engine/i);
    });
  });

  describe('ToolHandler Dispatch Integration', () => {
    it('dispatches excel_list_worksheets and excel_add_worksheet', async () => {
      const addRes = await handler.executeTool('excel_add_worksheet', {
        filename: fileName,
        worksheetName: 'ToolSheet',
      });
      expect(addRes.success).toBe(true);

      const listRes = await handler.executeTool('excel_list_worksheets', {
        filename: fileName,
      });
      expect(listRes.success).toBe(true);
      const sheets = listRes.data as WorksheetInfo[];
      expect(sheets.some((s) => s.name === 'ToolSheet')).toBe(true);
    });

    it('dispatches excel_rename_worksheet and excel_copy_worksheet', async () => {
      const renRes = await handler.executeTool('excel_rename_worksheet', {
        filename: fileName,
        oldName: 'Sheet1',
        newName: 'RenamedViaTool',
      });
      expect(renRes.success).toBe(true);

      const copyRes = await handler.executeTool('excel_copy_worksheet', {
        filename: fileName,
        sourceWorksheet: 'RenamedViaTool',
        targetName: 'ClonedViaTool',
      });
      expect(copyRes.success).toBe(true);
      expect(getWb().getWorksheet('ClonedViaTool')).toBeDefined();
    });

    it('dispatches excel_insert_rows, excel_delete_rows, excel_insert_columns, excel_delete_columns', async () => {
      expect((await handler.executeTool('excel_insert_rows', {
        filename: fileName, worksheet: 'Sheet1', startRow: 1, count: 2,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_delete_rows', {
        filename: fileName, worksheet: 'Sheet1', startRow: 1, count: 2,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_insert_columns', {
        filename: fileName, worksheet: 'Sheet1', startColumn: 'B', count: 1,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_delete_columns', {
        filename: fileName, worksheet: 'Sheet1', startColumn: 'B', count: 1,
      })).success).toBe(true);
    });

    it('dispatches excel_merge_cells and excel_unmerge_cells', async () => {
      expect((await handler.executeTool('excel_merge_cells', {
        filename: fileName, worksheet: 'Sheet1', startCell: 'C1', endCell: 'D2',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_unmerge_cells', {
        filename: fileName, worksheet: 'Sheet1', cellAddress: 'C1',
      })).success).toBe(true);
    });

    it('dispatches excel_add_table, excel_add_filter, excel_remove_filter', async () => {
      const ws = getWb().getWorksheet('Sheet1')!;
      ws.getCell('A1').value = 'H1'; ws.getCell('B1').value = 'H2';
      ws.getCell('A2').value = 1; ws.getCell('B2').value = 2;

      expect((await handler.executeTool('excel_add_table', {
        filename: fileName, worksheet: 'Sheet1', startCell: 'A1', endCell: 'B2', tableName: 'DispatchTbl',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_add_filter', {
        filename: fileName, worksheet: 'Sheet1', startCell: 'A1', endCell: 'B2',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_remove_filter', {
        filename: fileName, worksheet: 'Sheet1',
      })).success).toBe(true);
    });

    it('dispatches excel_delete_worksheet via tool', async () => {
      await handler.executeTool('excel_add_worksheet', { filename: fileName, worksheetName: 'ToDelete' });
      const delRes = await handler.executeTool('excel_delete_worksheet', {
        filename: fileName, worksheet: 'ToDelete',
      });
      expect(delRes.success).toBe(true);
      expect(getWb().getWorksheet('ToDelete')).toBeUndefined();
    });
  });
});
