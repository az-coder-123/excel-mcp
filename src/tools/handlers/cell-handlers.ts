/**
 * Cell Handlers
 * Single Responsibility: Cell and range data operations
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class CellHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleReadCell(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.readCell(filename, worksheet, cellAddress);
  }

  async handleReadRange(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    const range = this.parseCellRange(startCell, endCell);
    if (!range) {
      return { success: false, error: 'Invalid cell range format' };
    }

    return this.excelService.readRange(filename, worksheet, range);
  }

  async handleWriteCell(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const value = args['value'];

    if (!filename || !worksheet || !cellAddress || value === undefined) {
      return { success: false, error: 'Missing required parameters' };
    }

    // Try to parse as number if it looks like one
    let parsedValue: unknown = value;
    if (typeof value === 'string') {
      // Check if it's a formula
      if (value.startsWith('=')) {
        parsedValue = { formula: value };
      } else if (!isNaN(Number(value)) && value.trim() !== '') {
        parsedValue = Number(value);
      }
    }

    return this.excelService.writeCell(filename, worksheet, cellAddress, parsedValue);
  }

  async handleWriteBatch(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const data = args['data'] as Array<{ cellAddress: string; value: string }>;

    if (!filename || !worksheet || !data) {
      return { success: false, error: 'Missing required parameters' };
    }

    for (const item of data) {
      const result = await this.handleWriteCell({
        filename,
        worksheet,
        cellAddress: item.cellAddress,
        value: item.value
      });
      if (!result.success) return result;
    }

    return { success: true, data: { count: data.length } };
  }

  async handleGetCellInfo(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.readCell(filename, worksheet, cellAddress);
  }

  async handleCopyRange(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const sourceStart = this.getStringArg(args, 'sourceStart');
    const sourceEnd = this.getStringArg(args, 'sourceEnd');
    const targetStart = this.getStringArg(args, 'targetStart');
    const targetWorksheet = this.getStringArg(args, 'targetWorksheet');

    if (!filename || !worksheet || !sourceStart || !sourceEnd || !targetStart) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.copyRange(
      filename,
      worksheet,
      sourceStart,
      sourceEnd,
      targetStart,
      targetWorksheet
    );
  }

  async handleFindReplace(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const findText = this.getStringArg(args, 'findText');
    const replaceText = this.getStringArg(args, 'replaceText');
    const matchCase = this.getBooleanArg(args, 'matchCase');
    const matchEntireCell = this.getBooleanArg(args, 'matchEntireCell');

    if (!filename || !worksheet || !findText || replaceText === undefined) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.findReplace(
      filename,
      worksheet,
      findText,
      replaceText,
      matchCase,
      matchEntireCell
    );
  }

  async handleSortRange(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const sortColumn = this.getNumberArg(args, 'sortColumn');
    const ascending = this.getBooleanArg(args, 'ascending') ?? true;

    if (!filename || !worksheet || !startCell || !endCell || !sortColumn) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.sortRange(filename, worksheet, startCell, endCell, sortColumn, ascending);
  }

  async handleGetNamedRanges(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.getNamedRanges(filename);
  }

  async handleAddNamedRange(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const name = this.getStringArg(args, 'name');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !name || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addNamedRange(filename, worksheet, name, startCell, endCell);
  }

  async handleCalculateFormula(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.calculateFormula(filename);
  }
}