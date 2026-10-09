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
  /** Apply a two-color gradient fill (wires the previously-undispatched excel_set_gradient_fill). */
  public async setGradientFill(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string | undefined,
    color1: string,
    color2: string,
    type: string = 'horizontal'
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

      const degree = type === 'vertical' ? 90 : type === 'diagonal' ? 45 : 0;
      const range = this.parseRangeCells(startCell, endCell ?? startCell);
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          worksheet.getCell(r, c).fill = {
            type: 'gradient',
            gradient: 'angle',
            degree,
            stops: [
              { position: 0, color: { argb: this.toArgb(color1) } },
              { position: 1, color: { argb: this.toArgb(color2) } },
            ],
          };
        }
      }

      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /** Add a conditional formatting rule (cellValue / containsText / blanks / errors). */
  public async addConditionalFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    rule: {
      ruleType: string;
      operator?: string;
      formula1?: string;
      formula2?: string;
      format?: Record<string, unknown>;
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

      const ref = `${startCell}:${endCell}`;
      const style = this.buildCfStyle(rule.format);
      const ruleType = rule.ruleType;

      if (ruleType === 'colorScale' || ruleType === 'dataBar' || ruleType === 'iconSet') {
        return { success: false, error: `Use the dedicated ${ruleType} action for rule type "${ruleType}"` };
      }

      let cfRule: Record<string, unknown>;
      if (ruleType === 'cellValue') {
        const formulae = [rule.formula1 ?? '0'];
        if (rule.formula2 !== undefined && rule.operator === 'between') {
          formulae.push(rule.formula2);
        }
        cfRule = { type: 'cellIs', operator: rule.operator ?? 'equal', formulae, style };
      } else if (ruleType === 'containsText') {
        cfRule = { type: 'containsText', operator: 'containsText', text: rule.formula1 ?? '', style };
      } else if (ruleType === 'blanks') {
        cfRule = { type: 'expression', formulae: [`ISBLANK(${startCell})`], style };
      } else if (ruleType === 'errors') {
        cfRule = { type: 'expression', formulae: [`ISERROR(${startCell})`], style };
      } else {
        return { success: false, error: `Unsupported conditional format rule type: ${ruleType}` };
      }

      worksheet.addConditionalFormatting({ ref, rules: [cfRule] as unknown as ExcelJS.ConditionalFormattingOptions['rules'] });
      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /** Add a data-bar conditional format over a range. */
  public async addDataBar(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    color?: string
  ): Promise<OperationResult<void>> {
    const rule: Record<string, unknown> = {
      type: 'dataBar',
      cfvo: [{ type: 'min' }, { type: 'max' }],
    };
    if (color) rule.color = this.toArgb(color);
    return this.addVisualCf(filename, worksheetName, startCell, endCell, rule);
  }

  /** Add a color-scale conditional format over a range. */
  public async addColorScale(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    minColor?: string,
    midColor?: string,
    maxColor?: string
  ): Promise<OperationResult<void>> {
    const cfvo: Array<{ type: string; value?: number }> = [{ type: 'min' }];
    const colors: string[] = [];
    if (minColor) colors.push(this.toArgb(minColor));
    if (midColor) cfvo.push({ type: 'percentile', value: 50 });
    if (maxColor) colors.push(this.toArgb(maxColor));
    cfvo.push({ type: 'max' });
    return this.addVisualCf(filename, worksheetName, startCell, endCell, {
      type: 'colorScale',
      cfvo,
      color: colors,
    });
  }

  /** Add an icon-set conditional format over a range. */
  public async addIconSet(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    iconSet: string
  ): Promise<OperationResult<void>> {
    return this.addVisualCf(filename, worksheetName, startCell, endCell, {
      type: 'iconSet',
      iconSet,
      cfvo: [
        { type: 'percent', value: 0 },
        { type: 'percent', value: 33 },
        { type: 'percent', value: 67 },
      ],
      showValue: true,
    });
  }

  /** Remove conditional formatting rules attached to a specific range ref. */
  public async removeConditionalFormat(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string
  ): Promise<OperationResult<{ removed: number }>> {
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

      const ref = `${startCell}:${endCell}`;
      const store = worksheet as unknown as { conditionalFormattings: Array<{ ref: string }> };
      if (!Array.isArray(store.conditionalFormattings)) {
        return { success: true, data: { removed: 0 } };
      }
      const before = store.conditionalFormattings.length;
      store.conditionalFormattings = store.conditionalFormattings.filter((cf) => cf.ref !== ref);
      return { success: true, data: { removed: before - store.conditionalFormattings.length } };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  private async addVisualCf(
    filename: string,
    worksheetName: string,
    startCell: string,
    endCell: string,
    rule: Record<string, unknown>
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
      worksheet.addConditionalFormatting({
        ref: `${startCell}:${endCell}`,
        rules: [rule] as unknown as ExcelJS.ConditionalFormattingOptions['rules'],
      });
      return { success: true };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  private buildCfStyle(format?: Record<string, unknown>): Partial<ExcelJS.Style> {
    const style: Partial<ExcelJS.Style> = {};
    if (!format) return style;
    const fill = typeof format.fill === 'string' ? format.fill : undefined;
    const fontColor = typeof format.fontColor === 'string' ? format.fontColor : undefined;
    const bold = format.bold === true;
    if (fill) {
      style.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: this.toArgb(fill) } };
    }
    if (fontColor || bold) {
      style.font = {
        ...(fontColor ? { color: { argb: this.toArgb(fontColor) } } : {}),
        ...(bold ? { bold } : {}),
      };
    }
    return style;
  }

  private toArgb(color: string): string {
    const hex = color.replace('#', '').toUpperCase();
    return hex.length === 6 ? `FF${hex}` : hex;
  }

  private parseRangeCells(startCell: string, endCell: string): { startRow: number; endRow: number; startCol: number; endCol: number } {
    const parse = (addr: string) => {
      const match = addr.match(/^([A-Z]+)(\d+)$/i);
      if (!match) return { row: 1, col: 1 };
      let col = 0;
      for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
      return { row: parseInt(match[2], 10), col };
    };
    const a = parse(startCell);
    const b = parse(endCell);
    return {
      startRow: Math.min(a.row, b.row), endRow: Math.max(a.row, b.row),
      startCol: Math.min(a.col, b.col), endCol: Math.max(a.col, b.col),
    };
  }

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