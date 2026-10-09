/**
 * Protection Handlers
 * Single Responsibility: Worksheet/cell protection tool handlers
 * (wires the previously-undispatched protection tools)
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class ProtectionHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleProtectWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const password = this.getStringArg(args, 'password');
    const allowLocked = this.getBooleanArg(args, 'allowSelectLockedCells') ?? true;
    const allowUnlocked = this.getBooleanArg(args, 'allowSelectUnlockedCells') ?? true;
    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.protectWorksheet(filename, worksheet, password, allowLocked, allowUnlocked);
  }

  async handleUnprotectWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const password = this.getStringArg(args, 'password');
    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.unprotectWorksheet(filename, worksheet, password);
  }

  async handleProtectCells(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const locked = this.getBooleanArg(args, 'locked') ?? true;
    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.protectCells(filename, worksheet, startCell, endCell, locked);
  }

  async handleProtectWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.protectWorkbook(this.getStringArg(args, 'password'));
  }

  async handleUnprotectWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.unprotectWorkbook(this.getStringArg(args, 'password'));
  }
}
