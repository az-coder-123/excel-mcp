/**
 * Base Handler
 * Single Responsibility: Common handler utilities
 */

import { ExcelService } from '../../services/excel-service.js';

export abstract class BaseHandler {
  protected excelService: ExcelService;

  constructor(excelService: ExcelService) {
    this.excelService = excelService;
  }

  protected getStringArg(args: Record<string, unknown>, key: string): string | undefined {
    const value = args[key];
    return typeof value === 'string' ? value : undefined;
  }

  protected getNumberArg(args: Record<string, unknown>, key: string): number | undefined {
    const value = args[key];
    return typeof value === 'number' ? value : undefined;
  }

  protected getBooleanArg(args: Record<string, unknown>, key: string): boolean | undefined {
    const value = args[key];
    return typeof value === 'boolean' ? value : undefined;
  }

  protected parseCellAddress(address: string): { row: number; column: number } | null {
    const match = address.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return null;

    const colStr = match[1].toUpperCase();
    const row = parseInt(match[2], 10);

    let column = 0;
    for (let i = 0; i < colStr.length; i++) {
      column = column * 26 + (colStr.charCodeAt(i) - 64);
    }

    return { row, column };
  }

  protected parseCellRange(startCell: string, endCell: string): { start: { row: number; column: number }; end: { row: number; column: number } } | null {
    const start = this.parseCellAddress(startCell);
    const end = this.parseCellAddress(endCell);

    if (!start || !end) return null;

    return { start, end };
  }
}