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

  // New formatting methods
  
  /**
   * Set font style (bold, italic, underline, strikethrough)
   */
  public async setFontStyle(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    options: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean | string;
      strikethrough?: boolean;
    }
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          if (options.bold !== undefined) {
            cell.font = { ...cell.font, bold: options.bold };
          }
          if (options.italic !== undefined) {
            cell.font = { ...cell.font, italic: options.italic };
          }
          if (options.underline !== undefined) {
            cell.font = { ...cell.font, underline: options.underline as any };
          }
          if (options.strikethrough !== undefined) {
            cell.font = { ...cell.font, strike: options.strikethrough };
          }
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set font name and size
   */
  public async setFontNameSize(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    fontName?: string,
    fontSize?: number
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          if (fontName !== undefined) {
            cell.font = { ...cell.font, name: fontName };
          }
          if (fontSize !== undefined) {
            cell.font = { ...cell.font, size: fontSize };
          }
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set alignment
   */
  public async setAlignment(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    options: {
      horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
      vertical?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
      wrapText?: boolean;
      shrinkToFit?: boolean;
      indent?: number;
      textRotation?: number;
    }
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.alignment = {
            ...cell.alignment,
            ...(options.horizontal && { horizontal: options.horizontal }),
            ...(options.vertical && { vertical: options.vertical }),
            ...(options.wrapText !== undefined && { wrapText: options.wrapText }),
            ...(options.shrinkToFit !== undefined && { shrinkToFit: options.shrinkToFit }),
            ...(options.indent !== undefined && { indent: options.indent }),
            ...(options.textRotation !== undefined && { textRotation: options.textRotation }),
          };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Center text
   */
  public async centerText(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined
  ): Promise<OperationResult<void>> {
    return this.setAlignment(filename, worksheetName, startCell, endCell, {
      horizontal: 'center',
      vertical: 'middle',
    });
  }

  /**
   * Set border
   */
  public async setBorder(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    options: {
      borderStyle?: 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed' | 'hair' | 'mediumDashed' | 'mediumDashDot' | 'slantDashDot' | 'dashDot';
      borderColor?: string;
      top?: boolean;
      bottom?: boolean;
      left?: boolean;
      right?: boolean;
      diagonal?: boolean;
    }
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      const style = options.borderStyle || 'thin';
      const color = options.borderColor || 'FF000000';
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.border = {
            ...(cell.border || {}),
            ...(options.top !== false && { top: { style: style, color: { argb: color } } }),
            ...(options.bottom !== false && { bottom: { style: style, color: { argb: color } } }),
            ...(options.left !== false && { left: { style: style, color: { argb: color } } }),
            ...(options.right !== false && { right: { style: style, color: { argb: color } } }),
            ...(options.diagonal && { diagonal: { style: style, color: { argb: color } } }),
          };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Apply all borders
   */
  public async applyAllBorders(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    borderStyle: string = 'thin',
    borderColor: string = 'FF000000'
  ): Promise<OperationResult<void>> {
    return this.setBorder(filename, worksheetName, startCell, endCell, {
      borderStyle: borderStyle as any,
      borderColor,
      top: true,
      bottom: true,
      left: true,
      right: true,
    });
  }

  /**
   * Apply outline border
   */
  public async applyOutlineBorder(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    borderStyle: string = 'thin',
    borderColor: string = 'FF000000'
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      // Apply borders only to outline
      for (let col = range.start.column; col <= range.end.column; col++) {
        const topCell = worksheet.getCell(range.start.row, col);
        const bottomCell = worksheet.getCell(range.end.row, col);
        topCell.border = { ...topCell.border, top: { style: borderStyle as any, color: { argb: borderColor } } };
        bottomCell.border = { ...bottomCell.border, bottom: { style: borderStyle as any, color: { argb: borderColor } } };
      }
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        const leftCell = worksheet.getCell(row, range.start.column);
        const rightCell = worksheet.getCell(row, range.end.column);
        leftCell.border = { ...leftCell.border, left: { style: borderStyle as any, color: { argb: borderColor } } };
        rightCell.border = { ...rightCell.border, right: { style: borderStyle as any, color: { argb: borderColor } } };
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set background color
   */
  public async setBackgroundColor(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    color: string
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color },
          };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set font color
   */
  public async setFontColor(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    color: string
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.font = { ...cell.font, color: { argb: color } };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set number format
   */
  public async setNumberFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    format: string
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      const numFmt = this.getFormatCode(format);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.numFmt = numFmt;
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Apply header style
   */
  public async applyHeaderStyle(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    backgroundColor: string = 'FFD3D3D3',
    fontColor: string = 'FF000000'
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.font = { ...cell.font, bold: true, color: { argb: fontColor } };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: backgroundColor },
          };
          cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Apply title style
   */
  public async applyTitleStyle(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    fontSize: number = 16,
    color: string = 'FF1F4E78'
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

      const range = this.parseCellRangeOrSingle(startCell, endCell);
      
      for (let row = range.start.row; row <= range.end.row; row++) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          cell.font = { ...cell.font, bold: true, size: fontSize, color: { argb: color } };
          cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle' };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Apply currency format
   */
  public async applyCurrencyFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    symbol: string = '$',
    decimalPlaces: number = 2
  ): Promise<OperationResult<void>> {
    const format = symbol === '$' 
      ? `"${symbol}"#,##0.${'0'.repeat(decimalPlaces)}`
      : `"${symbol}" #,##0.${'0'.repeat(decimalPlaces)}`;
    return this.setNumberFormat(filename, worksheetName, startCell, endCell, format);
  }

  /**
   * Apply percentage format
   */
  public async applyPercentageFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    decimalPlaces: number = 2
  ): Promise<OperationResult<void>> {
    const format = `0.${'0'.repeat(decimalPlaces)}%`;
    return this.setNumberFormat(filename, worksheetName, startCell, endCell, format);
  }

  /**
   * Apply date format
   */
  public async applyDateFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    format: string = 'short'
  ): Promise<OperationResult<void>> {
    const formatCodes: Record<string, string> = {
      short: 'mm/dd/yyyy',
      long: 'dddd, mmmm dd, yyyy',
      iso: 'yyyy-mm-dd',
    };
    return this.setNumberFormat(filename, worksheetName, startCell, endCell, formatCodes[format] || format);
  }

  /**
   * Apply table style
   */
  public async applyTableStyle(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    headerBackgroundColor: string = 'FFD3D3D3',
    headerFontColor: string = 'FF000000',
    alternateRowColor: string = 'FFF9F9F9',
    hasHeader: boolean = true
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

      // Apply header style
      if (hasHeader) {
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(range.start.row, col);
          cell.font = { ...cell.font, bold: true, color: { argb: headerFontColor } };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: headerBackgroundColor },
          };
          cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
        }
      }

      // Apply data rows with alternating colors
      let alternate = false;
      for (let row = hasHeader ? range.start.row + 1 : range.start.row; row <= range.end.row; row++) {
        alternate = !alternate;
        for (let col = range.start.column; col <= range.end.column; col++) {
          const cell = worksheet.getCell(row, col);
          if (alternate) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: alternateRowColor },
            };
          }
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
            left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
            bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
            right: { style: 'thin', color: { argb: 'FFD9D9D9' } },
          };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Set rich text
   */
  public async setRichText(
    filename: string,
    worksheetName: string,
    cellAddress: string,
    richText: Array<{ text: string; bold?: boolean; italic?: boolean; underline?: boolean; fontColor?: string; fontSize?: number; fontName?: string }>
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
        richText: richText.map(segment => ({
          text: segment.text,
          font: {
            ...(segment.bold !== undefined && { bold: segment.bold }),
            ...(segment.italic !== undefined && { italic: segment.italic }),
            ...(segment.underline !== undefined && { underline: segment.underline }),
            ...(segment.fontColor && { color: { argb: segment.fontColor } }),
            ...(segment.fontSize && { size: segment.fontSize }),
            ...(segment.fontName && { name: segment.fontName }),
          }
        }))
      };

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  // Existing methods continue...
  
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
  private parseCellRangeOrSingle(startCell: string, endCell: string | undefined): { start: { row: number; column: number }; end: { row: number; column: number } } {
    if (endCell) {
      const range = this.parseCellRange(startCell, endCell);
      if (!range) {
        return { start: { row: 1, column: 1 }, end: { row: 1, column: 1 } };
      }
      return range;
    } else {
      const address = this.parseCellAddress(startCell);
      return { start: address, end: address };
    }
  }

  private parseCellAddress(address: string): { row: number; column: number } {
    const match = address.match(/^([A-Z]+)(\d+)$/i);
    if (!match) {
      return { row: 1, column: 1 };
    }

    const column = this.columnLetterToNumber(match[1].toUpperCase());
    const row = parseInt(match[2], 10);

    return { row, column };
  }

  private getFormatCode(format: string): string {
    const formatCodes: Record<string, string> = {
      currency: '$#,##0.00',
      percentage: '0.00%',
      number: '#,##0.00',
      integer: '#,##0',
      date: 'mm/dd/yyyy',
      time: 'hh:mm:ss',
      datetime: 'mm/dd/yyyy hh:mm:ss',
      fraction: '# ?/?',
      scientific: '0.00E+00',
      text: '@',
    };
    return formatCodes[format] || format;
  }

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