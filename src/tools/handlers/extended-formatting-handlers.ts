/**
 * Extended Formatting Handlers
 * All-in-one formatting operations
 */

import { ExcelService } from '../../services/excel-service.js';
import { OperationResult } from '../../types/index.js';
import { BaseHandler } from './base-handler.js';

export class ExtendedFormattingHandlers extends BaseHandler {
  constructor(excelService: ExcelService) {
    super(excelService);
  }

  async handleFormatWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const titleStyle = args['titleStyle'] as Record<string, unknown> | undefined;
    const headerStyle = args['headerStyle'] as Record<string, unknown> | undefined;
    const dataStyle = args['dataStyle'] as Record<string, unknown> | undefined;
    const columnWidthsStr = this.getStringArg(args, 'columnWidths');
    const autoFitColumns = this.getBooleanArg(args, 'autoFitColumns') ?? true;

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    try {
      // Apply title style
      if (titleStyle) {
        const cellAddress = this.getStringArg(titleStyle, 'cellAddress');
        if (cellAddress) {
          const fontSize = this.getNumberArg(titleStyle, 'fontSize') || 16;
          const color = this.getStringArg(titleStyle, 'color') || 'FF1F4E78';
          const bold = this.getBooleanArg(titleStyle, 'bold') ?? true;

          await this.excelService.applyTitleStyle(filename, worksheet, cellAddress, undefined, fontSize, color);
          if (bold) {
            await this.excelService.setFontStyle(filename, worksheet, cellAddress, undefined, { bold: true });
          }
        }
      }

      // Apply header style
      if (headerStyle) {
        const startCell = this.getStringArg(headerStyle, 'startCell');
        const endCell = this.getStringArg(headerStyle, 'endCell');
        if (startCell && endCell) {
          const backgroundColor = this.getStringArg(headerStyle, 'backgroundColor') || 'FFD3D3D3';
          const fontColor = this.getStringArg(headerStyle, 'fontColor') || 'FF000000';
          const bold = this.getBooleanArg(headerStyle, 'bold') ?? true;

          await this.excelService.applyHeaderStyle(filename, worksheet, startCell, endCell, backgroundColor, fontColor);
          if (bold) {
            await this.excelService.setFontStyle(filename, worksheet, startCell, endCell, { bold: true });
          }
        }
      }

      // Apply data style
      if (dataStyle) {
        const startCell = this.getStringArg(dataStyle, 'startCell');
        const endCell = this.getStringArg(dataStyle, 'endCell');
        if (startCell && endCell) {
          const borders = this.getBooleanArg(dataStyle, 'borders') ?? true;
          const borderStyle = this.getStringArg(dataStyle, 'borderStyle') || 'thin';
          const borderColor = this.getStringArg(dataStyle, 'borderColor') || 'FF000000';
          const alternateRows = this.getBooleanArg(dataStyle, 'alternateRows') ?? false;
          const alternateRowColor = this.getStringArg(dataStyle, 'alternateRowColor') || 'FFF9F9F9';

          if (borders) {
            await this.excelService.applyAllBorders(filename, worksheet, startCell, endCell, borderStyle, borderColor);
          }

          if (alternateRows) {
            await this.excelService.applyTableStyle(filename, worksheet, startCell, endCell, 'FFD3D3D3', 'FF000000', alternateRowColor, true);
          }
        }
      }

      // Apply column widths
      if (columnWidthsStr) {
        const columnWidths = JSON.parse(columnWidthsStr) as Array<{ column: string; width: number }>;
        for (const colWidth of columnWidths) {
          await this.excelService.setColumnWidth(filename, worksheet, colWidth.column, colWidth.width);
        }
      }

      // Auto-fit columns
      if (autoFitColumns) {
        await this.excelService.autoFitColumns(filename, worksheet);
      }

      return { success: true, data: { message: 'Worksheet formatted successfully' } };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  async handleBatchFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const operationsStr = this.getStringArg(args, 'operations');

    if (!filename || !worksheet || !operationsStr) {
      return { success: false, error: 'Missing required parameters' };
    }

    try {
      const operations = JSON.parse(operationsStr) as Array<Record<string, unknown>>;

      for (const op of operations) {
        const type = this.getStringArg(op, 'type');
        const startCell = this.getStringArg(op, 'startCell');
        const endCell = this.getStringArg(op, 'endCell');
        const cellAddress = this.getStringArg(op, 'cellAddress');
        const style = op['style'] as Record<string, unknown>;

        switch (type) {
          case 'titleStyle':
            if (cellAddress) {
              const fontSize = this.getNumberArg(style, 'fontSize') || 16;
              const color = this.getStringArg(style, 'color') || 'FF1F4E78';
              await this.excelService.applyTitleStyle(filename, worksheet, cellAddress, undefined, fontSize, color);
            }
            break;

          case 'headerStyle':
            if (startCell && endCell) {
              const backgroundColor = this.getStringArg(style, 'backgroundColor') || 'FFD3D3D3';
              const fontColor = this.getStringArg(style, 'fontColor') || 'FF000000';
              await this.excelService.applyHeaderStyle(filename, worksheet, startCell, endCell, backgroundColor, fontColor);
            }
            break;

          case 'borders':
            if (startCell && endCell) {
              const borderStyle = this.getStringArg(style, 'borderStyle') || 'thin';
              const borderColor = this.getStringArg(style, 'borderColor') || 'FF000000';
              await this.excelService.applyAllBorders(filename, worksheet, startCell, endCell, borderStyle, borderColor);
            }
            break;

          case 'autoFit':
            await this.excelService.autoFitColumns(filename, worksheet);
            break;

          case 'columnWidth':
            if (style && style['column'] && style['width']) {
              await this.excelService.setColumnWidth(filename, worksheet, style['column'] as string, style['width'] as number);
            }
            break;

          case 'rowHeight':
            if (startCell && style && style['height']) {
              const row = parseInt(startCell.replace(/[A-Z]/g, ''));
              await this.excelService.setRowHeight(filename, worksheet, row, style['height'] as number);
            }
            break;

          case 'backgroundColor':
            if (startCell && style && style['color']) {
              await this.excelService.setBackgroundColor(filename, worksheet, startCell, endCell, style['color'] as string);
            }
            break;

          case 'fontColor':
            if (startCell && style && style['color']) {
              await this.excelService.setFontColor(filename, worksheet, startCell, endCell, style['color'] as string);
            }
            break;

          case 'alignment':
            if (startCell) {
              await this.excelService.setAlignment(filename, worksheet, startCell, endCell, style as any);
            }
            break;

          default:
            break;
        }
      }

      return { success: true, data: { message: 'Batch formatting completed', count: operations.length } };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }
}