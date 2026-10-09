/**
 * Formatting Handlers
 * Single Responsibility: Cell and range formatting operations
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class FormattingHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleSetCellFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const format = args['format'] as Record<string, unknown>;

    if (!filename || !worksheet || !cellAddress || !format) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setCellFormat(filename, worksheet, cellAddress, format);
  }

  async handleSetRangeFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const format = args['format'] as Record<string, unknown>;

    if (!filename || !worksheet || !startCell || !endCell || !format) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setRangeFormat(filename, worksheet, startCell, endCell, format);
  }

  async handleAutoFitColumns(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startColumn = this.getStringArg(args, 'startColumn');
    const endColumn = this.getStringArg(args, 'endColumn');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.autoFitColumns(filename, worksheet, startColumn, endColumn);
  }

  async handleSetColumnWidth(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const column = this.getStringArg(args, 'column');
    const width = this.getNumberArg(args, 'width');

    if (!filename || !worksheet || !column || !width) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setColumnWidth(filename, worksheet, column, width);
  }

  async handleSetRowHeight(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const row = this.getNumberArg(args, 'row');
    const height = this.getNumberArg(args, 'height');

    if (!filename || !worksheet || !row || !height) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setRowHeight(filename, worksheet, row, height);
  }

  async handleFreezePanes(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.freezePanes(filename, worksheet, cellAddress);
  }

  async handleSetPrintArea(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setPrintArea(filename, worksheet, startCell, endCell);
  }

  async handleAddHeaderFooter(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const header = this.getStringArg(args, 'header');
    const footer = this.getStringArg(args, 'footer');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addHeaderFooter(filename, worksheet, header, footer);
  }

  async handleSetFontStyle(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const options = args['options'] as {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean | string;
      strikethrough?: boolean;
    };

    if (!filename || !worksheet || !startCell || !options) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setFontStyle(filename, worksheet, startCell, endCell, options);
  }

  async handleSetFontNameSize(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const fontName = this.getStringArg(args, 'fontName');
    const fontSize = this.getNumberArg(args, 'fontSize');

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setFontNameSize(filename, worksheet, startCell, endCell, fontName, fontSize);
  }

  async handleSetAlignment(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const options = args['options'] as {
      horizontal?: string;
      vertical?: string;
      wrapText?: boolean;
      shrinkToFit?: boolean;
      indent?: number;
      textRotation?: number;
    };

    if (!filename || !worksheet || !startCell || !options) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setAlignment(filename, worksheet, startCell, endCell, options);
  }

  async handleCenterText(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.centerText(filename, worksheet, startCell, endCell);
  }

  async handleSetBorder(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const options = args['options'] as {
      borderStyle?: string;
      borderColor?: string;
      top?: boolean;
      bottom?: boolean;
      left?: boolean;
      right?: boolean;
      diagonal?: boolean;
    };

    if (!filename || !worksheet || !startCell || !options) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setBorder(filename, worksheet, startCell, endCell, options);
  }

  async handleApplyAllBorders(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const borderStyle = this.getStringArg(args, 'borderStyle') || 'thin';
    const borderColor = this.getStringArg(args, 'borderColor') || 'FF000000';

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyAllBorders(filename, worksheet, startCell, endCell, borderStyle, borderColor);
  }

  async handleApplyOutlineBorder(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const borderStyle = this.getStringArg(args, 'borderStyle') || 'thin';
    const borderColor = this.getStringArg(args, 'borderColor') || 'FF000000';

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyOutlineBorder(filename, worksheet, startCell, endCell, borderStyle, borderColor);
  }

  async handleSetBackgroundColor(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const color = this.getStringArg(args, 'color');

    if (!filename || !worksheet || !startCell || !color) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setBackgroundColor(filename, worksheet, startCell, endCell, color);
  }

  async handleSetFontColor(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const color = this.getStringArg(args, 'color');

    if (!filename || !worksheet || !startCell || !color) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setFontColor(filename, worksheet, startCell, endCell, color);
  }

  async handleSetNumberFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const format = this.getStringArg(args, 'format');

    if (!filename || !worksheet || !startCell || !format) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setNumberFormat(filename, worksheet, startCell, endCell, format);
  }

  async handleApplyHeaderStyle(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const backgroundColor = this.getStringArg(args, 'backgroundColor') || 'FFD3D3D3';
    const fontColor = this.getStringArg(args, 'fontColor') || 'FF000000';

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyHeaderStyle(filename, worksheet, startCell, endCell, backgroundColor, fontColor);
  }

  async handleApplyTitleStyle(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const fontSize = this.getNumberArg(args, 'fontSize') || 16;
    const color = this.getStringArg(args, 'color') || 'FF1F4E78';

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyTitleStyle(filename, worksheet, startCell, endCell, fontSize, color);
  }

  async handleApplyCurrencyFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const symbol = this.getStringArg(args, 'symbol') || '$';
    const decimalPlaces = this.getNumberArg(args, 'decimalPlaces') || 2;

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyCurrencyFormat(filename, worksheet, startCell, endCell, symbol, decimalPlaces);
  }

  async handleApplyPercentageFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const decimalPlaces = this.getNumberArg(args, 'decimalPlaces') || 2;

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyPercentageFormat(filename, worksheet, startCell, endCell, decimalPlaces);
  }

  async handleApplyDateFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const format = this.getStringArg(args, 'format') || 'short';

    if (!filename || !worksheet || !startCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyDateFormat(filename, worksheet, startCell, endCell, format);
  }

  async handleApplyTableStyle(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const headerBackgroundColor = this.getStringArg(args, 'headerBackgroundColor') || 'FFD3D3D3';
    const headerFontColor = this.getStringArg(args, 'headerFontColor') || 'FF000000';
    const alternateRowColor = this.getStringArg(args, 'alternateRowColor') || 'FFF9F9F9';
    const hasHeader = this.getBooleanArg(args, 'hasHeader') || true;

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.applyTableStyle(
      filename,
      worksheet,
      startCell,
      endCell,
      headerBackgroundColor,
      headerFontColor,
      alternateRowColor,
      hasHeader
    );
  }

  async handleSetRichText(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const richText = args['richText'] as Array<{
      text: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontColor?: string;
      fontSize?: number;
      fontName?: string;
    }>;

    if (!filename || !worksheet || !cellAddress || !richText) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setRichText(filename, worksheet, cellAddress, richText);
  }
}
