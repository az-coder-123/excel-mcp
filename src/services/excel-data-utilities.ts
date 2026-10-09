/**
 * Excel Data Utilities Service
 * Single Responsibility: CSV transfer, dedupe, text splitting, lookup formulas,
 * pattern fill, and pivot summaries for the previously-undispatched data tools.
 */

import ExcelJS from 'exceljs';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult } from '../types/index.js';

interface GridRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export class ExcelDataUtilities {
  private permissionChecker: PermissionChecker;
  private logger: Logger;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(
    permissionChecker: PermissionChecker,
    logger: Logger,
    activeWorkbooks: Map<string, ExcelJS.Workbook>
  ) {
    this.permissionChecker = permissionChecker;
    this.logger = logger;
    this.activeWorkbooks = activeWorkbooks;
  }

  /** Import CSV text into a worksheet starting at startCell. */
  public async importCsv(
    filename: string,
    worksheetName: string,
    csvContent: string,
    startCell: string,
    delimiter: string = ','
  ): Promise<OperationResult<{ rowsImported: number }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'write');
    if (!gate.success || !gate.workbook || !gate.worksheet) {
      return { success: false, error: gate.error };
    }
    const { worksheet } = gate;

    try {
      const anchor = worksheet.getCell(startCell);
      const lines = csvContent.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
      while (lines.length > 0 && lines[lines.length - 1].trim() === '') {
        lines.pop();
      }

      let imported = 0;
      lines.forEach((line, r) => {
        if (line === '') return;
        imported++;
        line.split(delimiter).forEach((raw, c) => {
          const cell = worksheet.getCell(anchor.row + r, anchor.col + c);
          const text = raw.trim();
          cell.value = text === '' ? null : this.coerceScalar(text);
        });
      });

      this.logger.info(`Imported ${imported} CSV rows into ${filename}!${worksheetName}`);
      return { success: true, data: { rowsImported: imported } };
    } catch (error) {
      return this.fail('importCsv', error);
    }
  }

  /** Export a worksheet range to CSV text. */
  public async exportCsv(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    delimiter: string = ','
  ): Promise<OperationResult<{ csv: string }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'read');
    if (!gate.success || !gate.worksheet) {
      return { success: false, error: gate.error };
    }

    try {
      const range = this.resolveRange(gate.worksheet, startCell, endCell);
      const rows: string[] = [];
      for (let r = range.startRow; r <= range.endRow; r++) {
        const cols: string[] = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          const value = gate.worksheet.getCell(r, c).value;
          cols.push(this.cellToCsvText(value));
        }
        rows.push(cols.join(delimiter));
      }
      const csv = rows.join('\n');
      return { success: true, data: { csv } };
    } catch (error) {
      return this.fail('exportCsv', error);
    }
  }

  /** Remove duplicate rows in a range (optionally keyed on selected columns). */
  public async removeDuplicates(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    columns?: (string | number)[]
  ): Promise<OperationResult<{ removed: number; remaining: number }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'write');
    if (!gate.success || !gate.worksheet) {
      return { success: false, error: gate.error };
    }
    const { worksheet } = gate;

    try {
      const range = this.resolveRange(worksheet, startCell, endCell);
      const keyCols = (columns ?? []).map((c) =>
        typeof c === 'number' ? c : this.columnToIndex(c)
      );
      const effectiveCols = keyCols.length > 0 ? keyCols : this.range0(range.startCol, range.endCol);

      const seen = new Set<string>();
      const keptRows: unknown[][] = [];
      for (let r = range.startRow; r <= range.endRow; r++) {
        const row: unknown[] = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          row.push(worksheet.getCell(r, c).value);
        }
        const key = JSON.stringify(effectiveCols.map((c) => row[c - range.startCol] ?? null));
        if (!seen.has(key)) {
          seen.add(key);
          keptRows.push(row);
        }
      }

      const removed = (range.endRow - range.startRow + 1) - keptRows.length;
      keptRows.forEach((row, i) => {
        row.forEach((value, j) => {
          worksheet.getCell(range.startRow + i, range.startCol + j).value = value as ExcelJS.CellValue;
        });
      });
      for (let r = range.startRow + keptRows.length; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          worksheet.getCell(r, c).value = null;
        }
      }

      return { success: true, data: { removed, remaining: keptRows.length } };
    } catch (error) {
      return this.fail('removeDuplicates', error);
    }
  }

  /** Split one cell's text across multiple columns at targetCell. */
  public async textToColumns(
    filename: string,
    worksheetName: string,
    sourceCell: string,
    targetCell: string,
    delimiter: string,
    numberOfColumns: number
  ): Promise<OperationResult<{ columnsWritten: number }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'write');
    if (!gate.success || !gate.worksheet) {
      return { success: false, error: gate.error };
    }

    try {
      const source = gate.worksheet.getCell(sourceCell);
      const text = source.value === null || source.value === undefined ? '' : String(source.value);
      const parts = text === '' ? [] : text.split(delimiter).map((p) => p.trim());
      const anchor = gate.worksheet.getCell(targetCell);
      const columns = Math.max(1, Math.min(numberOfColumns, parts.length || 1));

      for (let c = 0; c < columns; c++) {
        gate.worksheet.getCell(anchor.row, anchor.col + c).value = parts[c] ?? null;
      }
      return { success: true, data: { columnsWritten: columns } };
    } catch (error) {
      return this.fail('textToColumns', error);
    }
  }

  /**
   * Flash fill: infer a transform from completed (input, expected) pairs in a
   * two-column source range, then apply it to rows whose target cell is empty.
   * Supported heuristics: identity, uppercase, lowercase, first token, last token.
   */
  public async flashFill(
    filename: string,
    worksheetName: string,
    sourceRange: string,
    targetRange: string
  ): Promise<OperationResult<{ filled: number; transform: string }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'write');
    if (!gate.success || !gate.worksheet) {
      return { success: false, error: gate.error };
    }
    const { worksheet } = gate;

    try {
      const src = this.parseRange(sourceRange);
      const tgt = this.parseRange(targetRange);
      if (!src || !tgt || src.endCol - src.startCol !== 1 || tgt.endCol !== tgt.startCol) {
        return { success: false, error: 'flashFill expects a 2-column sourceRange (input, expected) and a 1-column targetRange' };
      }

      const transforms: Array<{ name: string; fn: (v: string) => string }> = [
        { name: 'identity', fn: (v) => v },
        { name: 'uppercase', fn: (v) => v.toUpperCase() },
        { name: 'lowercase', fn: (v) => v.toLowerCase() },
        { name: 'first-token', fn: (v) => v.split(/[ ,.\-_]/)[0] },
        { name: 'last-token', fn: (v) => { const p = v.split(/[ ,.\-_]/); return p[p.length - 1]; } },
      ];

      const samples: Array<{ input: string; expected: string }> = [];
      const rowOf = (r: number) => ({
        input: String(worksheet.getCell(r, src.startCol).value ?? ''),
        expected: String(worksheet.getCell(r, src.endCol).value ?? ''),
      });
      for (let r = src.startRow; r <= src.endRow; r++) {
        const s = rowOf(r);
        if (s.input !== '' && s.expected !== '') samples.push(s);
      }
      if (samples.length === 0) {
        return { success: false, error: 'flashFill requires at least one completed (input, expected) sample row' };
      }

      const match = transforms.find((t) => samples.every((s) => t.fn(s.input) === s.expected));
      if (!match) {
        return { success: false, error: 'Unable to infer a pattern from the sample rows; supported: identity, uppercase, lowercase, first-token, last-token' };
      }

      let filled = 0;
      for (let r = src.startRow; r <= src.endRow; r++) {
        const input = String(worksheet.getCell(r, src.startCol).value ?? '');
        const target = worksheet.getCell(r, tgt.startCol);
        if (input !== '' && (target.value === null || target.value === undefined || target.value === '')) {
          target.value = match.fn(input);
          filled++;
        }
      }
      return { success: true, data: { filled, transform: match.name } };
    } catch (error) {
      return this.fail('flashFill', error);
    }
  }

  /** Write a VLOOKUP formula (evaluated by Excel when the file is opened). */
  public async vlookup(
    filename: string,
    worksheetName: string,
    targetCell: string,
    lookupValue: string | number,
    tableArray: string,
    colIndex: number,
    rangeLookup: boolean = true
  ): Promise<OperationResult<{ formula: string; note?: string }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'write');
    if (!gate.success || !gate.worksheet) {
      return { success: false, error: gate.error };
    }
    try {
      const lv = typeof lookupValue === 'string' ? `"${lookupValue.replace(/"/g, '""')}"` : String(lookupValue);
      const formula = `VLOOKUP(${lv},${tableArray},${colIndex},${rangeLookup})`;
      gate.worksheet.getCell(targetCell).value = { formula } as ExcelJS.CellFormulaValue;
      return {
        success: true,
        data: { formula, note: 'Stored as a formula; the result is evaluated when opened in Excel' },
      };
    } catch (error) {
      return this.fail('vlookup', error);
    }
  }

  /** Write an INDEX/MATCH formula (evaluated by Excel when the file is opened). */
  public async indexMatch(
    filename: string,
    worksheetName: string,
    targetCell: string,
    returnRange: string,
    lookupRange: string,
    lookupValue: string | number
  ): Promise<OperationResult<{ formula: string }>> {
    const gate = this.requireWorkAccess(filename, worksheetName, 'write');
    if (!gate.success || !gate.worksheet) {
      return { success: false, error: gate.error };
    }
    try {
      const lv = typeof lookupValue === 'string' ? `"${lookupValue.replace(/"/g, '""')}"` : String(lookupValue);
      const formula = `INDEX(${returnRange},MATCH(${lv},${lookupRange},0))`;
      gate.worksheet.getCell(targetCell).value = { formula } as ExcelJS.CellFormulaValue;
      return { success: true, data: { formula } };
    } catch (error) {
      return this.fail('indexMatch', error);
    }
  }

  /** Build a pivot summary (group by one row field, sum numeric value fields). */
  public async createPivotTable(
    filename: string,
    sourceWorksheet: string,
    sourceStartCell: string,
    sourceEndCell: string,
    targetWorksheet: string,
    targetCell: string,
    rowFields?: string[],
    _columnFields?: string[],
    valueFields?: string[]
  ): Promise<OperationResult<{ rows: number; valueFields: string[] }>> {
    const srcGate = this.requireWorkAccess(filename, sourceWorksheet, 'read');
    if (!srcGate.success || !srcGate.workbook || !srcGate.worksheet) {
      return { success: false, error: srcGate.error };
    }
    const { workbook, worksheet } = srcGate;

    try {
      const range = this.resolveRange(worksheet, sourceStartCell, sourceEndCell);
      const headers: string[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        headers.push(String(worksheet.getCell(range.startRow, c).value ?? ''));
      }

      const rowField = rowFields && rowFields.length > 0 ? rowFields[0] : headers[0];
      const rowIdx = headers.indexOf(rowField);
      if (rowIdx === -1) {
        return { success: false, error: `Row field "${rowField}" not found in source headers` };
      }
      const values = (valueFields && valueFields.length > 0
        ? valueFields
        : headers.filter((_h, i) => i !== rowIdx)
      ).filter((h) => headers.indexOf(h) !== -1 && headers.indexOf(h) !== rowIdx);
      if (values.length === 0) {
        return { success: false, error: 'No numeric value fields available to aggregate' };
      }

      const groups = new Map<string, number[]>();
      for (let r = range.startRow + 1; r <= range.endRow; r++) {
        const key = String(worksheet.getCell(r, range.startCol + rowIdx).value ?? '');
        if (key === '') continue;
        if (!groups.has(key)) groups.set(key, values.map(() => 0));
        const sums = groups.get(key)!;
        values.forEach((vf, i) => {
          const raw = worksheet.getCell(r, range.startCol + headers.indexOf(vf)).value;
          if (typeof raw === 'number') sums[i] += raw;
        });
      }

      const target = workbook.getWorksheet(targetWorksheet) ?? workbook.addWorksheet(targetWorksheet);
      const anchor = target.getCell(targetCell);
      target.getCell(anchor.row, anchor.col).value = rowField;
      values.forEach((vf, i) => {
        target.getCell(anchor.row, anchor.col + 1 + i).value = vf;
      });
      let out = 0;
      groups.forEach((sums, key) => {
        out++;
        target.getCell(anchor.row + out, anchor.col).value = key;
        sums.forEach((s, i) => {
          target.getCell(anchor.row + out, anchor.col + 1 + i).value = s;
        });
      });

      return { success: true, data: { rows: groups.size, valueFields: values } };
    } catch (error) {
      return this.fail('createPivotTable', error);
    }
  }

  // ---- shared helpers ----

  private requireWorkAccess(
    filename: string,
    worksheetName: string,
    permission: 'read' | 'write'
  ): OperationResult<{ workbook: ExcelJS.Workbook; worksheet: ExcelJS.Worksheet }> & { workbook?: ExcelJS.Workbook; worksheet?: ExcelJS.Worksheet } {
    const validation = this.permissionChecker.hasPermission(permission);
    if (!validation.success) {
      return { success: false, error: validation.error };
    }
    const workbook = this.activeWorkbooks.get(filename);
    if (!workbook) {
      return { success: false, error: `Workbook "${filename}" not opened` };
    }
    const worksheet = workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      return { success: false, error: `Worksheet "${worksheetName}" not found` };
    }
    // Expose workbook/worksheet at the top level as well as via data so that
    // call sites can read gate.workbook / gate.worksheet directly.
    return { success: true, data: { workbook, worksheet }, workbook, worksheet };
  }

  private resolveRange(worksheet: ExcelJS.Worksheet, startCell: string, endCell: string): GridRange {
    void worksheet;
    const a = this.parseAddress(startCell);
    const b = this.parseAddress(endCell);
    return {
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col),
      endCol: Math.max(a.col, b.col),
    };
  }

  private parseAddress(addr: string): { row: number; col: number } {
    const match = addr.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return { row: 1, col: 1 };
    let col = 0;
    for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
    return { row: parseInt(match[2], 10), col };
  }

  private parseRange(range: string): GridRange | null {
    const match = range.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
    if (!match) return null;
    const parse = (addr: string) => ({
      row: parseInt(addr.match(/\d+/)![0], 10),
      col: this.columnToIndex(addr.match(/^([A-Z]+)/i)![1]),
    });
    const a = parse(match[1]);
    const b = parse(match[2]);
    return {
      startRow: Math.min(a.row, b.row), endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col), endCol: Math.max(a.col, b.col),
    };
  }

  private range0(start: number, end: number): number[] {
    const out: number[] = [];
    for (let c = start; c <= end; c++) out.push(c);
    return out;
  }

  private columnToIndex(column: string): number {
    let n = 0;
    for (const ch of column.toUpperCase()) {
      n = n * 26 + (ch.charCodeAt(0) - 64);
    }
    return n;
  }

  private coerceScalar(text: string): string | number {
    if (/^-?\d+(\.\d+)?$/.test(text)) {
      return Number(text);
    }
    return text;
  }

  private cellToCsvText(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    if (typeof value === 'object') {
      const formula = (value as { formula?: string }).formula;
      return formula ? `=${formula}` : JSON.stringify(value);
    }
    const text = String(value);
    return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  private fail(op: string, error: unknown): { success: false; error: string } {
    const message = error instanceof Error ? error.message : 'Unknown error';
    this.logger.error(`Failed ${op}: ${message}`);
    return { success: false, error: message };
  }
}
