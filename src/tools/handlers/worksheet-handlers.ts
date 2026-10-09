/**
 * Worksheet Handlers
 * Single Responsibility: Worksheet and structure operations
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class WorksheetHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleListWorksheets(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.getWorksheets(filename);
  }

  async handleAddWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheetName = this.getStringArg(args, 'worksheetName');

    if (!filename || !worksheetName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addWorksheet(filename, worksheetName);
  }

  async handleDeleteWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteWorksheet(filename, worksheet);
  }

  async handleRenameWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const oldName = this.getStringArg(args, 'oldName');
    const newName = this.getStringArg(args, 'newName');

    if (!filename || !oldName || !newName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.renameWorksheet(filename, oldName, newName);
  }

  async handleCopyWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const sourceWorksheet = this.getStringArg(args, 'sourceWorksheet');
    const targetName = this.getStringArg(args, 'targetName');

    if (!filename || !sourceWorksheet || !targetName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.copyWorksheet(filename, sourceWorksheet, targetName);
  }

  async handleInsertRows(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startRow = this.getNumberArg(args, 'startRow');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startRow) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.insertRows(filename, worksheet, startRow, count);
  }

  async handleInsertColumns(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startColumn = this.getStringArg(args, 'startColumn');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startColumn) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.insertColumns(filename, worksheet, startColumn, count);
  }

  async handleDeleteRows(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startRow = this.getNumberArg(args, 'startRow');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startRow) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteRows(filename, worksheet, startRow, count);
  }

  async handleDeleteColumns(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startColumn = this.getStringArg(args, 'startColumn');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startColumn) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteColumns(filename, worksheet, startColumn, count);
  }

  async handleMergeCells(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.mergeCells(filename, worksheet, startCell, endCell);
  }

  async handleUnmergeCells(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.unmergeCells(filename, worksheet, cellAddress);
  }

  async handleAddTable(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const tableName = this.getStringArg(args, 'tableName');
    const style = this.getStringArg(args, 'style');

    if (!filename || !worksheet || !startCell || !endCell || !tableName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addTable(filename, worksheet, startCell, endCell, tableName, style);
  }

  async handleAddFilter(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addFilter(filename, worksheet, startCell, endCell);
  }

  async handleRemoveFilter(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.removeFilter(filename, worksheet);
  }
}