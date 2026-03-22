/**
 * Excel Service - Main orchestrator for Excel operations
 * Single Responsibility: Delegates to specialized modules
 */

import ExcelJS from 'exceljs';
import { PermissionChecker } from '../security/permission-checker.js';
import {
  CellRange,
  CellValue,
  OperationResult,
  WorkbookInfo,
  WorksheetInfo
} from '../types/index.js';
import { Logger } from '../utils/logger.js';
import { ExcelAccounting } from './excel-accounting.js';
import { ExcelAdvancedAccounting } from './excel-advanced-accounting.js';
import { ExcelCellOperations } from './excel-cell-operations.js';
import { ExcelFormatting } from './excel-formatting.js';
import { ExcelFormulaAnalyzer } from './excel-formula-analyzer.js';
import { ExcelStructureOperations } from './excel-structure-operations.js';
import { ExcelWorkbookManager } from './excel-workbook-manager.js';

export class ExcelService {
  private workbookManager: ExcelWorkbookManager;
  private cellOperations: ExcelCellOperations;
  private structureOperations: ExcelStructureOperations;
  private formatting: ExcelFormatting;
  public accounting: ExcelAccounting;
  public advancedAccounting: ExcelAdvancedAccounting;
  public formulaAnalyzer: ExcelFormulaAnalyzer;
  public activeWorkbooks: Map<string, ExcelJS.Workbook>;
  private logger: Logger;

  constructor(permissionChecker: PermissionChecker, logger: Logger) {
    this.activeWorkbooks = new Map();
    this.logger = logger;
    this.workbookManager = new ExcelWorkbookManager(permissionChecker, logger, this.activeWorkbooks);
    this.cellOperations = new ExcelCellOperations(permissionChecker, logger, this.activeWorkbooks);
    this.structureOperations = new ExcelStructureOperations(permissionChecker, logger, this.activeWorkbooks);
    this.formatting = new ExcelFormatting(permissionChecker, logger, this.activeWorkbooks);
    this.accounting = new ExcelAccounting(permissionChecker, logger, this.activeWorkbooks);
    this.advancedAccounting = new ExcelAdvancedAccounting(permissionChecker, logger, this.activeWorkbooks);
    this.formulaAnalyzer = new ExcelFormulaAnalyzer(this.activeWorkbooks);
  }

  // Workbook operations
  public async openWorkbook(filePath: string): Promise<OperationResult<WorkbookInfo>> {
    return this.workbookManager.openWorkbook(filePath);
  }

  public async createWorkbook(filename: string): Promise<OperationResult<WorkbookInfo>> {
    return this.workbookManager.createWorkbook(filename);
  }

  public async saveWorkbook(filename: string, outputPath?: string): Promise<OperationResult<void>> {
    return this.workbookManager.saveWorkbook(filename, outputPath);
  }

  public closeWorkbook(filename: string): OperationResult<void> {
    return this.workbookManager.closeWorkbook(filename);
  }

  public async exportWorksheetToNewFile(
    filename: string,
    worksheetName: string,
    newFilePath: string
  ): Promise<OperationResult<void>> {
    return this.workbookManager.exportWorksheetToNewFile(filename, worksheetName, newFilePath);
  }

  public getWorkbookContext(filename: string): OperationResult<{
    isOpen: boolean;
    filePath: string;
    worksheets: WorksheetInfo[];
    currentWorksheet: string | null;
  }> {
    return this.workbookManager.getWorkbookContext(filename);
  }

  // Cell operations
  public async readCell(filename: string, worksheetName: string, cellAddress: string): Promise<OperationResult<CellValue>> {
    return this.cellOperations.readCell(filename, worksheetName, cellAddress);
  }

  public async readRange(filename: string, worksheetName: string, range: CellRange): Promise<OperationResult<CellValue[][]>> {
    return this.cellOperations.readRange(filename, worksheetName, range);
  }

  public async writeCell(filename: string, worksheetName: string, cellAddress: string, value: unknown): Promise<OperationResult<void>> {
    return this.cellOperations.writeCell(filename, worksheetName, cellAddress, value);
  }

  public async writeBatch(filename: string, worksheetName: string, data: Array<{ cellAddress: string; value: unknown }>): Promise<OperationResult<{ count: number }>> {
    return this.cellOperations.writeBatch(filename, worksheetName, data);
  }

  public async copyRange(filename: string, worksheetName: string, sourceStart: string, sourceEnd: string, targetStart: string, targetWorksheet?: string): Promise<OperationResult<void>> {
    return this.cellOperations.copyRange(filename, worksheetName, sourceStart, sourceEnd, targetStart, targetWorksheet);
  }

  public async findReplace(filename: string, worksheetName: string, findText: string, replaceText: string, matchCase?: boolean, matchEntireCell?: boolean): Promise<OperationResult<{ count: number }>> {
    return this.cellOperations.findReplace(filename, worksheetName, findText, replaceText, matchCase, matchEntireCell);
  }

  public async sortRange(filename: string, worksheetName: string, startCell: string, endCell: string, sortColumn: number, ascending: boolean): Promise<OperationResult<void>> {
    return this.cellOperations.sortRange(filename, worksheetName, startCell, endCell, sortColumn, ascending);
  }

  public async getNamedRanges(filename: string): Promise<OperationResult<Array<{ name: string; ref: string }>>> {
    return this.cellOperations.getNamedRanges(filename);
  }

  public async addNamedRange(filename: string, worksheetName: string, name: string, startCell: string, endCell: string): Promise<OperationResult<void>> {
    return this.cellOperations.addNamedRange(filename, worksheetName, name, startCell, endCell);
  }

  // Structure operations
  public getWorksheets(filename: string): OperationResult<WorksheetInfo[]> {
    return this.structureOperations.getWorksheets(filename);
  }

  public async addWorksheet(filename: string, worksheetName: string): Promise<OperationResult<WorksheetInfo>> {
    return this.structureOperations.addWorksheet(filename, worksheetName);
  }

  public async deleteWorksheet(filename: string, worksheetName: string): Promise<OperationResult<void>> {
    return this.structureOperations.deleteWorksheet(filename, worksheetName);
  }

  public async renameWorksheet(filename: string, oldName: string, newName: string): Promise<OperationResult<void>> {
    return this.structureOperations.renameWorksheet(filename, oldName, newName);
  }

  public async copyWorksheet(filename: string, sourceWorksheet: string, targetName: string): Promise<OperationResult<void>> {
    return this.structureOperations.copyWorksheet(filename, sourceWorksheet, targetName);
  }

  public async insertRows(filename: string, worksheetName: string, startRow: number, count: number): Promise<OperationResult<void>> {
    return this.structureOperations.insertRows(filename, worksheetName, startRow, count);
  }

  public async insertColumns(filename: string, worksheetName: string, startColumn: string, count: number): Promise<OperationResult<void>> {
    return this.structureOperations.insertColumns(filename, worksheetName, startColumn, count);
  }

  public async deleteRows(filename: string, worksheetName: string, startRow: number, count: number): Promise<OperationResult<void>> {
    return this.structureOperations.deleteRows(filename, worksheetName, startRow, count);
  }

  public async deleteColumns(filename: string, worksheetName: string, startColumn: string, count: number): Promise<OperationResult<void>> {
    return this.structureOperations.deleteColumns(filename, worksheetName, startColumn, count);
  }

  public async mergeCells(filename: string, worksheetName: string, startCell: string, endCell: string): Promise<OperationResult<void>> {
    return this.structureOperations.mergeCells(filename, worksheetName, startCell, endCell);
  }

  public async unmergeCells(filename: string, worksheetName: string, cellAddress: string): Promise<OperationResult<void>> {
    return this.structureOperations.unmergeCells(filename, worksheetName, cellAddress);
  }

  public async addTable(filename: string, worksheetName: string, startCell: string, endCell: string, tableName: string, style?: string): Promise<OperationResult<void>> {
    return this.structureOperations.addTable(filename, worksheetName, startCell, endCell, tableName, style);
  }

  public async addFilter(filename: string, worksheetName: string, startCell: string, endCell: string): Promise<OperationResult<void>> {
    return this.structureOperations.addFilter(filename, worksheetName, startCell, endCell);
  }

  public async removeFilter(filename: string, worksheetName: string): Promise<OperationResult<void>> {
    return this.structureOperations.removeFilter(filename, worksheetName);
  }

  public async calculateFormula(filename: string): Promise<OperationResult<void>> {
    return this.structureOperations.calculateFormula(filename);
  }

  // Formatting operations
  public async setCellFormat(filename: string, worksheetName: string, cellAddress: string, format: Record<string, unknown>): Promise<OperationResult<void>> {
    return this.formatting.setCellFormat(filename, worksheetName, cellAddress, format);
  }

  public async setRangeFormat(filename: string, worksheetName: string, startCell: string, endCell: string, format: Record<string, unknown>): Promise<OperationResult<void>> {
    return this.formatting.setRangeFormat(filename, worksheetName, startCell, endCell, format);
  }

  public async autoFitColumns(filename: string, worksheetName: string, startColumn?: string, endColumn?: string): Promise<OperationResult<void>> {
    return this.formatting.autoFitColumns(filename, worksheetName, startColumn, endColumn);
  }

  public async setColumnWidth(filename: string, worksheetName: string, column: string, width: number): Promise<OperationResult<void>> {
    return this.formatting.setColumnWidth(filename, worksheetName, column, width);
  }

  public async setRowHeight(filename: string, worksheetName: string, row: number, height: number): Promise<OperationResult<void>> {
    return this.formatting.setRowHeight(filename, worksheetName, row, height);
  }

  public async addComment(filename: string, worksheetName: string, cellAddress: string, comment: string, author?: string): Promise<OperationResult<void>> {
    return this.formatting.addComment(filename, worksheetName, cellAddress, comment, author);
  }

  public async removeComment(filename: string, worksheetName: string, cellAddress: string): Promise<OperationResult<void>> {
    return this.formatting.removeComment(filename, worksheetName, cellAddress);
  }

  public async addHyperlink(filename: string, worksheetName: string, cellAddress: string, url: string, display?: string): Promise<OperationResult<void>> {
    return this.formatting.addHyperlink(filename, worksheetName, cellAddress, url, display);
  }

  public async addDataValidation(filename: string, worksheetName: string, cellAddress: string, type: string, formula1: string, formula2?: string, operator?: string, errorMessage?: string): Promise<OperationResult<void>> {
    return this.formatting.addDataValidation(filename, worksheetName, cellAddress, type, formula1, formula2, operator, errorMessage);
  }

  public async freezePanes(filename: string, worksheetName: string, cellAddress: string): Promise<OperationResult<void>> {
    return this.formatting.freezePanes(filename, worksheetName, cellAddress);
  }

  public async setPrintArea(filename: string, worksheetName: string, startCell: string, endCell: string): Promise<OperationResult<void>> {
    return this.formatting.setPrintArea(filename, worksheetName, startCell, endCell);
  }

  public async addHeaderFooter(filename: string, worksheetName: string, header?: string, footer?: string): Promise<OperationResult<void>> {
    return this.formatting.addHeaderFooter(filename, worksheetName, header, footer);
  }

  public async setFontStyle(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, options: { bold?: boolean; italic?: boolean; underline?: boolean | string; strikethrough?: boolean }): Promise<OperationResult<void>> {
    return this.formatting.setFontStyle(filename, worksheetName, startCell, endCell, options);
  }

  public async setFontNameSize(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, fontName?: string, fontSize?: number): Promise<OperationResult<void>> {
    return this.formatting.setFontNameSize(filename, worksheetName, startCell, endCell, fontName, fontSize);
  }

  public async setAlignment(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, options: { horizontal?: string; vertical?: string; wrapText?: boolean; shrinkToFit?: boolean; indent?: number; textRotation?: number }): Promise<OperationResult<void>> {
    return this.formatting.setAlignment(filename, worksheetName, startCell, endCell, options as any);
  }

  public async centerText(filename: string, worksheetName: string, startCell: string, endCell: string | undefined): Promise<OperationResult<void>> {
    return this.formatting.centerText(filename, worksheetName, startCell, endCell);
  }

  public async setBorder(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, options: { borderStyle?: string; borderColor?: string; top?: boolean; bottom?: boolean; left?: boolean; right?: boolean; diagonal?: boolean }): Promise<OperationResult<void>> {
    return this.formatting.setBorder(filename, worksheetName, startCell, endCell, options as any);
  }

  public async applyAllBorders(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, borderStyle?: string, borderColor?: string): Promise<OperationResult<void>> {
    return this.formatting.applyAllBorders(filename, worksheetName, startCell, endCell, borderStyle, borderColor);
  }

  public async applyOutlineBorder(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, borderStyle?: string, borderColor?: string): Promise<OperationResult<void>> {
    return this.formatting.applyOutlineBorder(filename, worksheetName, startCell, endCell, borderStyle, borderColor);
  }

  public async setBackgroundColor(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, color: string): Promise<OperationResult<void>> {
    return this.formatting.setBackgroundColor(filename, worksheetName, startCell, endCell, color);
  }

  public async setFontColor(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, color: string): Promise<OperationResult<void>> {
    return this.formatting.setFontColor(filename, worksheetName, startCell, endCell, color);
  }

  public async setNumberFormat(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, format: string): Promise<OperationResult<void>> {
    return this.formatting.setNumberFormat(filename, worksheetName, startCell, endCell, format);
  }

  public async applyHeaderStyle(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, backgroundColor?: string, fontColor?: string): Promise<OperationResult<void>> {
    return this.formatting.applyHeaderStyle(filename, worksheetName, startCell, endCell, backgroundColor, fontColor);
  }

  public async applyTitleStyle(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, fontSize?: number, color?: string): Promise<OperationResult<void>> {
    return this.formatting.applyTitleStyle(filename, worksheetName, startCell, endCell, fontSize, color);
  }

  public async applyCurrencyFormat(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, symbol?: string, decimalPlaces?: number): Promise<OperationResult<void>> {
    return this.formatting.applyCurrencyFormat(filename, worksheetName, startCell, endCell, symbol, decimalPlaces);
  }

  public async applyPercentageFormat(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, decimalPlaces?: number): Promise<OperationResult<void>> {
    return this.formatting.applyPercentageFormat(filename, worksheetName, startCell, endCell, decimalPlaces);
  }

  public async applyDateFormat(filename: string, worksheetName: string, startCell: string, endCell: string | undefined, format?: string): Promise<OperationResult<void>> {
    return this.formatting.applyDateFormat(filename, worksheetName, startCell, endCell, format);
  }

  public async applyTableStyle(filename: string, worksheetName: string, startCell: string, endCell: string, headerBackgroundColor?: string, headerFontColor?: string, alternateRowColor?: string, hasHeader?: boolean): Promise<OperationResult<void>> {
    return this.formatting.applyTableStyle(filename, worksheetName, startCell, endCell, headerBackgroundColor, headerFontColor, alternateRowColor, hasHeader);
  }

  public async setRichText(filename: string, worksheetName: string, cellAddress: string, richText: Array<{ text: string; bold?: boolean; italic?: boolean; underline?: boolean; fontColor?: string; fontSize?: number; fontName?: string }>): Promise<OperationResult<void>> {
    return this.formatting.setRichText(filename, worksheetName, cellAddress, richText);
  }

  // Chart operations
  // Note: ExcelJS has limited chart support. These methods store chart metadata for reference.
  // Full chart creation and manipulation requires ExcelJS's native chart API or external libraries.
  
  public async addChart(
    filename: string,
    worksheetName: string,
    dataRange: string,
    chartType: string,
    targetCell: string,
    title?: string,
    width?: number,
    height?: number
  ): Promise<OperationResult<{ chartIndex: number }>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // Validate data range format
      const rangeMatch = dataRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (!rangeMatch) {
        return { success: false, error: `Invalid data range format: ${dataRange}. Expected format: A1:D10` };
      }

      // Validate target cell format
      const targetMatch = targetCell.match(/^([A-Z]+)(\d+)$/);
      if (!targetMatch) {
        return { success: false, error: `Invalid target cell format: ${targetCell}. Expected format: A1` };
      }

      // Store chart metadata for reference
      // Note: ExcelJS chart creation is limited, metadata is stored for future implementation
      const charts = (worksheet as any).__charts || [];
      const chartIndex = charts.push({
        type: chartType,
        dataRange,
        targetCell,
        title: title || 'Chart',
        width: width || 400,
        height: height || 300,
      });

      (worksheet as any).__charts = charts;

      this.logger.info(`Chart metadata stored for ${filename}!${worksheetName} at index ${chartIndex - 1}`);
      return { success: true, data: { chartIndex: chartIndex - 1 } };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to add chart: ${errorMessage}`);
      return { success: false, error: `Failed to add chart: ${errorMessage}` };
    }
  }

  public async updateChart(
    filename: string,
    worksheetName: string,
    chartIndex: number,
    dataRange: string
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // Get stored chart metadata
      const charts = (worksheet as any).__charts;
      if (!charts || !Array.isArray(charts)) {
        return { success: false, error: 'No charts found in worksheet' };
      }

      if (chartIndex < 0 || chartIndex >= charts.length) {
        return { success: false, error: `Chart index ${chartIndex} out of range` };
      }

      // Validate data range format
      const rangeMatch = dataRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (!rangeMatch) {
        return { success: false, error: `Invalid data range format: ${dataRange}. Expected format: A1:D10` };
      }

      // Update chart metadata
      charts[chartIndex].dataRange = dataRange;

      this.logger.info(`Updated chart ${chartIndex} in ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to update chart: ${errorMessage}`);
      return { success: false, error: `Failed to update chart: ${errorMessage}` };
    }
  }

  public async deleteChart(
    filename: string,
    worksheetName: string,
    chartIndex: number
  ): Promise<OperationResult<void>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // Get stored chart metadata
      const charts = (worksheet as any).__charts;
      if (!charts || !Array.isArray(charts)) {
        return { success: false, error: 'No charts found in worksheet' };
      }

      if (chartIndex < 0 || chartIndex >= charts.length) {
        return { success: false, error: `Chart index ${chartIndex} out of range` };
      }

      // Remove chart metadata
      charts.splice(chartIndex, 1);

      this.logger.info(`Deleted chart ${chartIndex} from ${filename}!${worksheetName}`);
      return { success: true };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to delete chart: ${errorMessage}`);
      return { success: false, error: `Failed to delete chart: ${errorMessage}` };
    }
  }

  public async listCharts(
    filename: string,
    worksheetName: string
  ): Promise<OperationResult<Array<{ index: number; type: string; title: string }>>> {
    try {
      const workbook = this.activeWorkbooks.get(filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(worksheetName);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${worksheetName}" not found` };
      }

      // Get stored chart metadata
      const charts = (worksheet as any).__charts;
      if (!charts || !Array.isArray(charts)) {
        return { success: true, data: [] };
      }

      const chartList = charts.map((chart: any, index: number) => ({
        index,
        type: chart.type || 'Unknown',
        title: chart.title || 'Untitled',
      }));

      return { success: true, data: chartList };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Failed to list charts: ${errorMessage}`);
      return { success: false, error: `Failed to list charts: ${errorMessage}` };
    }
  }
}