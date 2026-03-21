/**
 * Excel Formatting - Handles cell and range formatting operations
 * Single Responsibility: Cell formatting, comments, hyperlinks
 */

import ExcelJS from 'exceljs';
import { OperationResult } from '../types/index.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';

export class ExcelFormatting {
  private permissionChecker: PermissionChecker;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(
    permissionChecker: PermissionChecker,
    _logger: Logger,
    activeWorkbooks: Map<string, ExcelJS.Workbook>
  ) {
    this.permissionChecker = permissionChecker;
    this.activeWorkbooks = activeWorkbooks;
  }

  /**
   * Set cell formatting
   */
  public async setCellFormat(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    format: Record<string, unknown>
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const cell = worksheet.getCell(cellAddress);
      this.applyFormat(cell, format);

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set range formatting
   */
  public async setRangeFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    format: Record<string, unknown>
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const range = this.parseCellRange(startCell, endCell);
      if (!range) {
        return { success: false, error: 'Invalid cell range' };
      }

      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          this.applyFormat(cell, format);
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Auto-fit column widths
   */
  public async autoFitColumns(
    filename: string,
    worksheetName: string,
    startColumn?: string,
    endColumn?: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      // ExcelJS doesn't have autoFit, so we set a reasonable default width
      const start = startColumn ? this.columnLetterToNumber(startColumn) : 1;
      const end = endColumn ? this.columnLetterToNumber(endColumn) : worksheet.columnCount;

      for (let col = start; col <= end; col++) {
        worksheet.getColumn(col).width = 15;
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set column width
   */
  public async setColumnWidth(
    filename: string,
    worksheetName: string,
    column: string,
    width: number
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const colIndex = this.columnLetterToNumber(column);
      worksheet.getColumn(colIndex).width = width;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set row height
   */
  public async setRowHeight(
    filename: string,
    worksheetName: string,
    row: number,
    height: number
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      worksheet.getRow(row).height = height;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add comment to cell
   */
  public async addComment(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    comment: string,
    author?: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const cell = worksheet.getCell(cellAddress);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const note: any = {
        texts: [{ text: comment }],
      };
      if (author) {
        note.author = author;
      }
      cell.note = note;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Remove comment from cell
   */
  public async removeComment(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const cell = worksheet.getCell(cellAddress);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (cell as any).note = undefined;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add hyperlink
   */
  public async addHyperlink(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    url: string,
    display?: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const cell = worksheet.getCell(cellAddress);
      cell.value = {
        text: display || url,
        hyperlink: url,
      };

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add data validation
   */
  public async addDataValidation(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    type: string,
    formula1: string,
    formula2?: string,
    operator?: string,
    errorMessage?: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      const cell = worksheet.getCell(cellAddress);
      
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const validationRule: any = {
        type: type,
        formulae: formula2 ? [formula1, formula2] : [formula1],
      };

      if (operator) {
        validationRule.operator = operator;
      }

      if (errorMessage) {
        validationRule.error = errorMessage;
      }

      cell.dataValidation = validationRule;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Freeze panes
   */
  public async freezePanes(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      worksheet.views = [{
        state: 'frozen',
        xSplit: this.columnLetterToNumber(cellAddress.replace(/[0-9]/g, '')) - 1,
        ySplit: parseInt(cellAddress.replace(/[A-Z]/gi, '')) - 1,
      }];

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set print area
   */
  public async setPrintArea(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      worksheet.pageSetup.printArea = `${startCell}:${endCell}`;

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Add header and footer
   */
  public async addHeaderFooter(
    filename: string,
    worksheetName: string,
    header?: string,
    footer?: string
  ): Promise<OperationResult<void>> {
    try {
      const validation = this.permissionChecker.hasPermission('write');
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

      worksheet.headerFooter = {
        oddHeader: header,
        oddFooter: footer,
      };

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  // Helper methods
  private applyFormat(cell: ExcelJS.Cell, format: Record<string, unknown>): void {
    if (format.bold !== undefined) {
      cell.font = { ...cell.font, bold: format.bold as boolean };
    }
    if (format.italic !== undefined) {
      cell.font = { ...cell.font, italic: format.italic as boolean };
    }
    if (format.underline !== undefined) {
      cell.font = { ...cell.font, underline: format.underline as boolean };
    }
    if (format.fontSize !== undefined) {
      cell.font = { ...cell.font, size: format.fontSize as number };
    }
    if (format.fontColor !== undefined) {
      cell.font = { ...cell.font, color: { argb: format.fontColor as string } };
    }
    if (format.backgroundColor !== undefined) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: format.backgroundColor as string },
      };
    }
    if (format.horizontalAlignment !== undefined) {
      cell.alignment = { ...cell.alignment, horizontal: format.horizontalAlignment as ExcelJS.Alignment['horizontal'] };
    }
    if (format.verticalAlignment !== undefined) {
      cell.alignment = { ...cell.alignment, vertical: format.verticalAlignment as ExcelJS.Alignment['vertical'] };
    }
    if (format.wrapText !== undefined) {
      cell.alignment = { ...cell.alignment, wrapText: format.wrapText as boolean };
    }
    if (format.numberFormat !== undefined) {
      cell.numFmt = format.numberFormat as string;
    }
    if (format.borderColor !== undefined) {
      cell.border = {
        top: { style: 'thin', color: { argb: format.borderColor as string } },
        left: { style: 'thin', color: { argb: format.borderColor as string } },
        bottom: { style: 'thin', color: { argb: format.borderColor as string } },
        right: { style: 'thin', color: { argb: format.borderColor as string } },
      };
    }
  }

  private columnLetterToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result;
  }

  private parseCellRange(startCell: string, endCell: string): { start: { row: number; column: number }; end: { row: number; column: number } } | null {
    const parseAddress = (address: string): { row: number; column: number } | null => {
      const match = address.match(/^([A-Z]+)(\d+)$/i);
      if (!match) return null;

      const column = this.columnLetterToNumber(match[1].toUpperCase());
      const row = parseInt(match[2], 10);

      return { row, column };
    };

    const start = parseAddress(startCell);
    const end = parseAddress(endCell);

    if (!start || !end) return null;

    return { start, end };
  }
}