/**
 * Unit and Integration Tests for AnalysisHandlers
 * Testing data analysis, statistics, filtering, grouping, profiling, and comparison tools:
 * - getColumnStats (value distribution, unique counts, missing cells, case-insensitive columns)
 * - filterData (single & compound filters, operators: equals, contains, greaterThan, lessThan)
 * - groupAggregate (groupByColumn + aggregate operations: count, sum, avg, min, max)
 * - profileData (data typing inference, null counts, stats per column)
 * - search (exact, contains, regex, column range restricted search)
 * - compareRanges (identifying differences between ranges across sheets)
 * - ToolHandler dispatch integration
 */

import ExcelJS from 'exceljs';
import { ExcelService } from '../src/services/excel-service.js';
import { ToolHandler } from '../src/tools/tool-handler.js';
import { AnalysisHandlers } from '../src/tools/handlers/analysis-handlers.js';
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

describe('AnalysisHandlers & Data Analytics Tools', () => {
  let workbooks: Map<string, ExcelJS.Workbook>;
  let analysis: AnalysisHandlers;
  let excelService: ExcelService;
  let handler: ToolHandler;
  const fileName = 'test_analysis.xlsx';
  const sheetName = 'Data';

  beforeEach(() => {
    workbooks = new Map();
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    excelService = new ExcelService(permChecker, logger);
    analysis = new AnalysisHandlers(workbooks);
    handler = new ToolHandler(excelService, permChecker, logger);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);

    // Populate sample dataset
    // Row 1: Header
    ws.getCell('A1').value = 'Department';
    ws.getCell('B1').value = 'Employee';
    ws.getCell('C1').value = 'Salary';
    ws.getCell('D1').value = 'Status';

    // Rows 2-6: Records
    ws.getCell('A2').value = 'Sales';   ws.getCell('B2').value = 'Alice'; ws.getCell('C2').value = 5000; ws.getCell('D2').value = 'Active';
    ws.getCell('A3').value = 'Eng';     ws.getCell('B3').value = 'Bob';   ws.getCell('C3').value = 8000; ws.getCell('D3').value = 'Active';
    ws.getCell('A4').value = 'Sales';   ws.getCell('B4').value = 'Carol'; ws.getCell('C4').value = 6000; ws.getCell('D4').value = 'Inactive';
    ws.getCell('A5').value = 'Eng';     ws.getCell('B5').value = 'Dave';  ws.getCell('C5').value = 9000; ws.getCell('D5').value = 'Active';
    ws.getCell('A6').value = 'HR';      ws.getCell('B6').value = 'Eve';   ws.getCell('C6').value = 4500; ws.getCell('D6').value = 'Active';

    workbooks.set(fileName, wb);
    excelService.getActiveWorkbooks().set(fileName, wb);
  });

  describe('getColumnStats', () => {
    it('computes distribution, uniqueness, and counts for a column', async () => {
      const res = await analysis.getColumnStats(fileName, sheetName, 'A', true);
      expect(res.success).toBe(true);
      const data = res.data as { totalCells: number; nonEmptyCells: number; uniqueCount: number; duplicates: number };
      expect(data.nonEmptyCells).toBe(5);
      expect(data.uniqueCount).toBe(3); // Sales, Eng, HR
      expect(data.duplicates).toBe(2);
    });

    it('works with lowercase column letter', async () => {
      const res = await analysis.getColumnStats(fileName, sheetName, 'c', true);
      expect(res.success).toBe(true);
      const data = res.data as { nonEmptyCells: number };
      expect(data.nonEmptyCells).toBe(5);
    });

    it('returns error when workbook or sheet not found', async () => {
      expect((await analysis.getColumnStats('missing.xlsx', sheetName, 'A')).success).toBe(false);
      expect((await analysis.getColumnStats(fileName, 'MissingSheet', 'A')).success).toBe(false);
    });
  });

  describe('filterData', () => {
    it('filters rows matching conditions (equals and greaterThan)', async () => {
      const res = await analysis.filterData(
        fileName, sheetName, 'A1', 'D6',
        [
          { column: 'A', operator: 'equals', value: 'Eng' },
          { column: 'C', operator: 'greaterThan', value: '8500' },
        ],
        true
      );
      expect(res.success).toBe(true);
      const rows = (res.data as unknown as { filteredRows: Array<Record<string, unknown>> }).filteredRows;
      expect(rows).toHaveLength(1);
      expect(rows[0]['B']).toBe('Dave');
    });

    it('filters using contains operator', async () => {
      const res = await analysis.filterData(
        fileName, sheetName, 'A1', 'D6',
        [{ column: 'B', operator: 'contains', value: 'li' }], // matches Alice
        true
      );
      expect(res.success).toBe(true);
      const rows = (res.data as unknown as { filteredRows: Array<Record<string, unknown>> }).filteredRows;
      expect(rows).toHaveLength(1);
      expect(rows[0]['B']).toBe('Alice');
    });
  });

  describe('groupAggregate', () => {
    it('groups by department and computes sum of salary', async () => {
      const res = await analysis.groupAggregate(
        fileName, sheetName, 'A1', 'C6', 'A', 'C', 'sum', true
      );
      expect(res.success).toBe(true);
      const groups = (res.data as unknown as { groups: Array<Record<string, unknown>> }).groups;
      const sales = groups.find((g) => g['A'] === 'Sales');
      const eng = groups.find((g) => g['A'] === 'Eng');
      const hr = groups.find((g) => g['A'] === 'HR');

      expect(sales?.['sum_C']).toBe(11000); // 5000 + 6000
      expect(eng?.['sum_C']).toBe(17000);   // 8000 + 9000
      expect(hr?.['sum_C']).toBe(4500);
    });

    it('groups by department and computes count', async () => {
      const res = await analysis.groupAggregate(
        fileName, sheetName, 'A1', 'B6', 'A', undefined, 'count', true
      );
      expect(res.success).toBe(true);
      const groups = (res.data as unknown as { groups: Array<Record<string, unknown>> }).groups;
      const sales = groups.find((g) => g['A'] === 'Sales');
      expect(sales?.['count']).toBe(2);
    });
  });

  describe('profileData', () => {
    it('profiles all columns in a range', async () => {
      const res = await analysis.profileData(fileName, sheetName, 'A1', 'D6', true);
      expect(res.success).toBe(true);
      const profiles = (res.data as unknown as { columnProfiles: Array<{ column: string; totalCells: number }> }).columnProfiles;
      expect(profiles).toHaveLength(4);
      expect(profiles[0].column).toBe('A');
      expect(profiles[0].totalCells).toBe(6);
    });
  });

  describe('search', () => {
    it('searches cells with exact and contains match', async () => {
      const exactRes = await analysis.search(fileName, sheetName, 'Alice', 'exact');
      expect(exactRes.success).toBe(true);
      const exactMatches = (exactRes.data as unknown as { matches: Array<{ address: string }> }).matches;
      expect(exactMatches).toHaveLength(1);
      expect(exactMatches[0].address).toBe('B2');

      const containsRes = await analysis.search(fileName, sheetName, 'Act', 'contains');
      expect(containsRes.success).toBe(true);
      const containsMatches = (containsRes.data as unknown as { matches: Array<{ cell: string }> }).matches;
      expect(containsMatches.length).toBeGreaterThanOrEqual(4); // 4 'Active' cells
    });

    it('searches within specific column range', async () => {
      const res = await analysis.search(fileName, sheetName, 'Eng', 'contains', 'A:A');
      expect(res.success).toBe(true);
      const matches = (res.data as unknown as { matches: Array<{ cell: string }> }).matches;
      expect(matches).toHaveLength(2); // Rows 3 and 5 in Col A
    });
  });

  describe('compareRanges', () => {
    it('compares two identical ranges and reports identical true', async () => {
      const res = await analysis.compareRanges(
        fileName, sheetName, 'A1', 'B3',
        undefined, 'A1', 'B3'
      );
      expect(res.success).toBe(true);
      const data = (res.data as unknown as { differentCells: number; differences: unknown[] });
      expect(data.differentCells).toBe(0);
      expect(data.differences).toHaveLength(0);
    });

    it('detects differences between two ranges', async () => {
      const res = await analysis.compareRanges(
        fileName, sheetName, 'A1', 'A3', // Department, Sales, Eng
        undefined, 'D1', 'D3'            // Status, Active, Active
      );
      expect(res.success).toBe(true);
      const data = (res.data as unknown as { differentCells: number; differences: unknown[] });
      expect(data.differentCells).toBeGreaterThan(0);
    });
  });

  describe('ToolHandler Dispatch Integration', () => {
    it('dispatches excel_get_column_stats, excel_filter_data, excel_group_aggregate', async () => {
      const statsRes = await handler.executeTool('excel_get_column_stats', {
        filename: fileName, worksheet: sheetName, column: 'A', hasHeader: true,
      });
      expect(statsRes.success).toBe(true);

      const filterRes = await handler.executeTool('excel_filter_data', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'D6',
        filters: [{ column: 'A', operator: 'equals', value: 'HR' }],
        hasHeader: true,
      });
      expect(filterRes.success).toBe(true);

      const groupRes = await handler.executeTool('excel_group_aggregate', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'C6',
        groupByColumn: 'A', aggregateColumn: 'C', operation: 'avg', hasHeader: true,
      });
      expect(groupRes.success).toBe(true);
    });

    it('dispatches excel_profile_data, excel_search, excel_compare_ranges', async () => {
      const profRes = await handler.executeTool('excel_profile_data', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'D6', hasHeader: true,
      });
      expect(profRes.success).toBe(true);

      const searchRes = await handler.executeTool('excel_search', {
        filename: fileName, worksheet: sheetName, searchQuery: 'Alice', searchType: 'exact',
      });
      expect(searchRes.success).toBe(true);

      const compRes = await handler.executeTool('excel_compare_ranges', {
        filename: fileName, worksheet1: sheetName, range1Start: 'A1', range1End: 'B2',
        range2Start: 'A1', range2End: 'B2',
      });
      expect(compRes.success).toBe(true);
    });
  });
});
