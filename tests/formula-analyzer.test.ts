/**
 * Unit and Integration Tests for ExcelFormulaAnalyzer
 * Testing all tools, edge cases, circular detection, and ToolHandler dispatch:
 * - listFormulas (retrieves all formulas in sheet)
 * - analyzeFormula (extracts functions, precedents, complexity)
 * - getDependencies (precedents and dependents mapping)
 * - explainFormula (structured steps and explanation)
 * - checkCircular (self-referencing and circular loops detection)
 * - auditFormulas (whole-sheet formula statistics and error auditing)
 * - ToolHandler dispatch for all formula analysis tools
 */

import ExcelJS from 'exceljs';
import { ExcelFormulaAnalyzer } from '../src/services/excel-formula-analyzer.js';
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

describe('ExcelFormulaAnalyzer', () => {
  let workbooks: Map<string, ExcelJS.Workbook>;
  let analyzer: ExcelFormulaAnalyzer;
  let excelService: ExcelService;
  let handler: ToolHandler;
  const fileName = 'test_formulas.xlsx';
  const sheetName = 'Formulas';

  beforeEach(() => {
    workbooks = new Map();
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    analyzer = new ExcelFormulaAnalyzer(workbooks);
    excelService = new ExcelService(permChecker, logger);
    handler = new ToolHandler(excelService, permChecker, logger);

    const wb = new ExcelJS.Workbook();
    wb.addWorksheet(sheetName);
    workbooks.set(fileName, wb);
    excelService.getActiveWorkbooks().set(fileName, wb);
  });

  const getWs = (): ExcelJS.Worksheet => workbooks.get(fileName)!.getWorksheet(sheetName)!;

  describe('listFormulas', () => {
    it('lists all formula cells in worksheet and ignores static values', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = 200;
      ws.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 300 };
      ws.getCell('B1').value = { formula: 'AVERAGE(A1:A2)', result: 150 };

      const formulas = await analyzer.listFormulas(fileName, sheetName);
      expect(formulas).toHaveLength(2);
      expect(formulas.map(f => f.address).sort()).toEqual(['A3', 'B1'].sort());
      expect(formulas.find(f => f.address === 'A3')?.formula).toBe('SUM(A1:A2)');
    });

    it('throws error for missing workbook or worksheet', async () => {
      await expect(analyzer.listFormulas('missing.xlsx', sheetName)).rejects.toThrow('not found');
      await expect(analyzer.listFormulas(fileName, 'MissingSheet')).rejects.toThrow('not found');
    });
  });

  describe('analyzeFormula', () => {
    it('extracts functions, cell references, and complexity', async () => {
      const ws = getWs();
      ws.getCell('C1').value = { formula: 'IF(A1>0,SUM(B1:B10),0)', result: 50 };

      const info = await analyzer.analyzeFormula(fileName, sheetName, 'C1');
      expect(info.formula).toBe('IF(A1>0,SUM(B1:B10),0)');
      expect(info.functions).toContain('SUM');
      expect(info.cellReferences).toContain('A1');
      expect(info.complexity.depth).toBeGreaterThanOrEqual(1);
    });
  });

  describe('getDependencies & tracePrecedents/Dependents', () => {
    it('maps precedents and dependents correctly', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 10;
      ws.getCell('B1').value = { formula: 'A1*2', result: 20 };
      ws.getCell('C1').value = { formula: 'B1+5', result: 25 };

      const depsB1 = await analyzer.getDependencies(fileName, sheetName, 'B1');
      expect(depsB1.precedents).toContain('A1');
      expect(depsB1.dependents).toContain('C1');
      expect(depsB1.isCircular).toBe(false);
    });
  });

  describe('explainFormula', () => {
    it('provides clear explanation, steps, and suggestions', async () => {
      const ws = getWs();
      ws.getCell('A10').value = { formula: 'SUM(A1:A9)', result: 500 };

      const exp = await analyzer.explainFormula(fileName, sheetName, 'A10');
      expect(exp.description).toBeDefined();
      expect(exp.steps.length).toBeGreaterThan(0);
      expect(exp.functions.map(f => f.name)).toContain('SUM');
    });
  });

  describe('checkCircularReferences', () => {
    it('detects circular loop between cells (e.g. A1 -> B1 -> A1)', async () => {
      const ws = getWs();
      ws.getCell('A1').value = { formula: 'B1+1' };
      ws.getCell('B1').value = { formula: 'A1*2' };

      const res = await handler.executeTool('excel_check_circular', {
        filename: fileName,
        worksheet: sheetName,
      });

      expect(res.success).toBe(true);
      const data = res.data as { hasCircularReferences: boolean; circularCells: Array<{ cell: string; chain: string[] }> };
      expect(data.hasCircularReferences).toBe(true);
      const cellNames = data.circularCells.map(c => c.cell);
      expect(cellNames).toContain('A1');
      expect(cellNames).toContain('B1');
    });

    it('reports no circular references for clean linear formulas', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 10;
      ws.getCell('A2').value = { formula: 'A1+10' };
      ws.getCell('A3').value = { formula: 'A2*2' };

      const res = await handler.executeTool('excel_check_circular', {
        filename: fileName,
        worksheet: sheetName,
      });

      expect(res.success).toBe(true);
      const data = res.data as { hasCircularReferences: boolean };
      expect(data.hasCircularReferences).toBe(false);
    });
  });

  describe('auditFormulas', () => {
    it('aggregates total formula count, complexity, and function frequencies', async () => {
      const ws = getWs();
      ws.getCell('A1').value = 10;
      ws.getCell('A2').value = 20;
      ws.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 30 };
      ws.getCell('A4').value = { formula: 'SUM(A1:A3)', result: 60 };
      ws.getCell('B1').value = { formula: 'AVERAGE(A1:A2)', result: 15 };

      const audit = await analyzer.auditFormulas(fileName, sheetName);
      expect(audit.totalFormulas).toBe(3);
      expect(audit.statistics.mostUsedFunctions[0].function).toBe('SUM');
      expect(audit.statistics.mostUsedFunctions[0].count).toBe(2);
      expect(audit.errors).toHaveLength(0);
    });
  });

  describe('ToolHandler Integration for Formula Tools', () => {
    it('dispatches excel_list_formulas through ToolHandler', async () => {
      const ws = excelService.getActiveWorkbooks().get(fileName)!.getWorksheet(sheetName)!;
      ws.getCell('A1').value = { formula: '1+1', result: 2 };

      const res = await handler.executeTool('excel_list_formulas', {
        filename: fileName,
        worksheet: sheetName,
      });

      expect(res.success).toBe(true);
      expect((res.data as any).total).toBe(1);
    });

    it('dispatches excel_analyze_formula through ToolHandler', async () => {
      const ws = excelService.getActiveWorkbooks().get(fileName)!.getWorksheet(sheetName)!;
      ws.getCell('D1').value = { formula: 'SUM(A1:B1)', result: 10 };

      const res = await handler.executeTool('excel_analyze_formula', {
        filename: fileName,
        worksheet: sheetName,
        cellAddress: 'D1',
      });

      expect(res.success).toBe(true);
      expect((res.data as any).formula).toBe('SUM(A1:B1)');
      expect((res.data as any).functions).toContain('SUM');
    });

    it('dispatches excel_audit_formulas through ToolHandler', async () => {
      const res = await handler.executeTool('excel_audit_formulas', {
        filename: fileName,
        worksheet: sheetName,
      });

      expect(res.success).toBe(true);
      expect((res.data as any).totalFormulas).toBeDefined();
    });
  });
});
