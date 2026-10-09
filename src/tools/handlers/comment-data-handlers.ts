/**
 * Comment and Data Handlers
 * Single Responsibility: Comments, hyperlinks, and data validation operations
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class CommentDataHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleAddComment(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const comment = this.getStringArg(args, 'comment');
    const author = this.getStringArg(args, 'author');

    if (!filename || !worksheet || !cellAddress || !comment) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addComment(filename, worksheet, cellAddress, comment, author);
  }

  async handleRemoveComment(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.removeComment(filename, worksheet, cellAddress);
  }

  async handleAddHyperlink(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const url = this.getStringArg(args, 'url');
    const display = this.getStringArg(args, 'display');

    if (!filename || !worksheet || !cellAddress || !url) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addHyperlink(filename, worksheet, cellAddress, url, display);
  }

  async handleAddDataValidation(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const type = this.getStringArg(args, 'type');
    const formula1 = this.getStringArg(args, 'formula1');
    const formula2 = this.getStringArg(args, 'formula2');
    const operator = this.getStringArg(args, 'operator');
    const errorMessage = this.getStringArg(args, 'errorMessage');

    if (!filename || !worksheet || !cellAddress || !type || !formula1) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addDataValidation(
      filename,
      worksheet,
      cellAddress,
      type,
      formula1,
      formula2,
      operator,
      errorMessage
    );
  }
}