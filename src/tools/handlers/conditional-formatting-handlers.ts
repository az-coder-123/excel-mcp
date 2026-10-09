/**
 * Conditional Formatting Handlers
 * Single Responsibility: Conditional formatting tool handlers
 * (wires the previously-undispatched conditional formatting tools)
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class ConditionalFormattingHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleAddConditionalFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const ruleType = this.getStringArg(args, 'ruleType');
    if (!filename || !worksheet || !startCell || !endCell || !ruleType) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.addConditionalFormat(filename, worksheet, startCell, endCell, {
      ruleType,
      operator: this.getStringArg(args, 'operator'),
      formula1: this.getStringArg(args, 'formula1'),
      formula2: this.getStringArg(args, 'formula2'),
      format: this.getObjectArg(args, 'format'),
    });
  }

  async handleRemoveConditionalFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.removeConditionalFormat(filename, worksheet, startCell, endCell);
  }

  async handleAddDataBar(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const color = this.getStringArg(args, 'color');
    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.addDataBar(filename, worksheet, startCell, endCell, color);
  }

  async handleAddColorScale(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const minColor = this.getStringArg(args, 'minColor');
    const midColor = this.getStringArg(args, 'midColor');
    const maxColor = this.getStringArg(args, 'maxColor');
    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.addColorScale(filename, worksheet, startCell, endCell, minColor, midColor, maxColor);
  }

  async handleAddIconSet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const iconSet = this.getStringArg(args, 'iconSet') ?? '3Arrows';
    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }
    return this.excelService.addIconSet(filename, worksheet, startCell, endCell, iconSet);
  }
}
