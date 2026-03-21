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
}