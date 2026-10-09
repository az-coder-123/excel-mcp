/**
 * Chart Handlers
 * Single Responsibility: Chart-related tool handlers
 */

import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import { BaseHandler } from './base-handler.js';

export class ChartHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  public async handleAddChart(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const dataRange = this.getStringArg(args, 'dataRange');
    const chartType = this.getStringArg(args, 'chartType');
    const targetCell = this.getStringArg(args, 'targetCell');
    const title = this.getStringArg(args, 'title');
    const width = this.getNumberArg(args, 'width');
    const height = this.getNumberArg(args, 'height');

    if (!filename || !worksheet || !dataRange || !chartType || !targetCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addChart(
      filename,
      worksheet,
      dataRange,
      chartType,
      targetCell,
      title,
      width,
      height
    );
  }

  public async handleUpdateChart(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const chartIndex = this.getNumberArg(args, 'chartIndex');
    const dataRange = this.getStringArg(args, 'dataRange');

    if (!filename || !worksheet || chartIndex === undefined || !dataRange) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.updateChart(filename, worksheet, chartIndex, dataRange);
  }

  public async handleDeleteChart(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const chartIndex = this.getNumberArg(args, 'chartIndex');

    if (!filename || !worksheet || chartIndex === undefined) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteChart(filename, worksheet, chartIndex);
  }

  public async handleListCharts(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.listCharts(filename, worksheet);
  }
}
