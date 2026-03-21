/**
 * Excel Service - Main orchestrator for Excel operations
 * Single Responsibility: Delegates to specialized modules
 */

import ExcelJS from 'exceljs';
import { 
  CellValue, 
  WorksheetInfo, 
  WorkbookInfo, 
  OperationResult,
  CellRange 
} from '../types/index.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { ExcelWorkbookManager } from './excel-workbook-manager.js';
import { ExcelCellOperations } from './excel-cell-operations.js';
import { ExcelStructureOperations } from './excel-structure-operations.js';
import { ExcelFormatting } from './excel-formatting.js';

export class ExcelService {
  private workbookManager: ExcelWorkbookManager;
  private cellOperations: ExcelCellOperations;
  private structureOperations: ExcelStructureOperations;
  private formatting: ExcelFormatting;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(permissionChecker: PermissionChecker, logger: Logger) {
    this.activeWorkbooks = new Map();
    this.workbookManager = new ExcelWorkbookManager(permissionChecker, logger, this.activeWorkbooks);
    this.cellOperations = new ExcelCellOperations(permissionChecker, logger, this.activeWorkbooks);
    this.structureOperations = new ExcelStructureOperations(permissionChecker, logger, this.activeWorkbooks);
    this.formatting = new ExcelFormatting(permissionChecker, logger, this.activeWorkbooks);
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
}