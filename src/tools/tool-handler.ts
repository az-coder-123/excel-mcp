/**
 * Tool Handler - Processes MCP tool calls
 * Single Responsibility: Handles tool execution and validation
 */

import { ExcelService } from '../services/excel-service.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult, CellRange } from '../types/index.js';
import { getToolDefinition } from './tool-definitions.js';

export class ToolHandler {
  private excelService: ExcelService;
  private permissionChecker: PermissionChecker;
  private logger: Logger;

  constructor(
    excelService: ExcelService,
    permissionChecker: PermissionChecker,
    logger: Logger
  ) {
    this.excelService = excelService;
    this.permissionChecker = permissionChecker;
    this.logger = logger;
  }

  /**
   * Execute a tool call
   */
  public async executeTool(
    toolName: string,
    args: Record<string, unknown>
  ): Promise<OperationResult> {
    this.logger.info(`Executing tool: ${toolName}`, { args });

    // Validate tool exists
    const toolDef = getToolDefinition(toolName);
    if (!toolDef) {
      return {
        success: false,
        error: `Unknown tool: ${toolName}`,
      };
    }

    // Validate permissions
    for (const permission of toolDef.requiredPermissions) {
      const permResult = this.permissionChecker.hasPermission(permission);
      if (!permResult.success) {
        return permResult;
      }
    }

    // Execute tool
    try {
      switch (toolName) {
        case 'excel_open_workbook':
          return await this.handleOpenWorkbook(args);
        case 'excel_create_workbook':
          return await this.handleCreateWorkbook(args);
        case 'excel_read_cell':
          return await this.handleReadCell(args);
        case 'excel_read_range':
          return await this.handleReadRange(args);
        case 'excel_write_cell':
          return await this.handleWriteCell(args);
        case 'excel_list_worksheets':
          return await this.handleListWorksheets(args);
        case 'excel_add_worksheet':
          return await this.handleAddWorksheet(args);
        case 'excel_save_workbook':
          return await this.handleSaveWorkbook(args);
        case 'excel_close_workbook':
          return await this.handleCloseWorkbook(args);
        case 'excel_get_cell_info':
          return await this.handleGetCellInfo(args);
        case 'excel_write_batch':
          return await this.handleWriteBatch(args);
        case 'excel_insert_rows':
          return await this.handleInsertRows(args);
        case 'excel_insert_columns':
          return await this.handleInsertColumns(args);
        case 'excel_delete_rows':
          return await this.handleDeleteRows(args);
        case 'excel_delete_columns':
          return await this.handleDeleteColumns(args);
        case 'excel_merge_cells':
          return await this.handleMergeCells(args);
        case 'excel_unmerge_cells':
          return await this.handleUnmergeCells(args);
        case 'excel_copy_range':
          return await this.handleCopyRange(args);
        case 'excel_set_cell_format':
          return await this.handleSetCellFormat(args);
        case 'excel_set_range_format':
          return await this.handleSetRangeFormat(args);
        case 'excel_auto_fit_columns':
          return await this.handleAutoFitColumns(args);
        case 'excel_set_column_width':
          return await this.handleSetColumnWidth(args);
        case 'excel_set_row_height':
          return await this.handleSetRowHeight(args);
        case 'excel_add_data_validation':
          return await this.handleAddDataValidation(args);
        case 'excel_add_table':
          return await this.handleAddTable(args);
        case 'excel_freeze_panes':
          return await this.handleFreezePanes(args);
        case 'excel_add_filter':
          return await this.handleAddFilter(args);
        case 'excel_remove_filter':
          return await this.handleRemoveFilter(args);
        case 'excel_delete_worksheet':
          return await this.handleDeleteWorksheet(args);
        case 'excel_rename_worksheet':
          return await this.handleRenameWorksheet(args);
        case 'excel_copy_worksheet':
          return await this.handleCopyWorksheet(args);
        case 'excel_set_print_area':
          return await this.handleSetPrintArea(args);
        case 'excel_add_header_footer':
          return await this.handleAddHeaderFooter(args);
        case 'excel_add_hyperlink':
          return await this.handleAddHyperlink(args);
        case 'excel_add_comment':
          return await this.handleAddComment(args);
        case 'excel_remove_comment':
          return await this.handleRemoveComment(args);
        case 'excel_find_replace':
          return await this.handleFindReplace(args);
        case 'excel_sort_range':
          return await this.handleSortRange(args);
        case 'excel_calculate_formula':
          return await this.handleCalculateFormula(args);
        case 'excel_get_named_ranges':
          return await this.handleGetNamedRanges(args);
        case 'excel_add_named_range':
          return await this.handleAddNamedRange(args);
        default:
          return {
            success: false,
            error: `Tool not implemented: ${toolName}`,
          };
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      this.logger.error(`Tool execution failed: ${message}`);
      return { success: false, error: message };
    }
  }

  private async handleOpenWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filePath = this.getStringArg(args, 'filePath');
    if (!filePath) {
      return { success: false, error: 'Missing required parameter: filePath' };
    }
    return this.excelService.openWorkbook(filePath);
  }

  private async handleCreateWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const outputPath = this.getStringArg(args, 'outputPath');
    
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    if (!outputPath) {
      return { success: false, error: 'Missing required parameter: outputPath' };
    }

    const fullPath = `${outputPath.replace(/\/$/, '')}/${filename}.xlsx`;
    return this.excelService.createWorkbook(fullPath);
  }

  private async handleReadCell(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.readCell(filename, worksheet, cellAddress);
  }

  private async handleReadRange(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleWriteCell(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleListWorksheets(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.getWorksheets(filename);
  }

  private async handleAddWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheetName = this.getStringArg(args, 'worksheetName');

    if (!filename || !worksheetName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addWorksheet(filename, worksheetName);
  }

  private async handleSaveWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const outputPath = this.getStringArg(args, 'outputPath');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.saveWorkbook(filename, outputPath || undefined);
  }

  private async handleCloseWorkbook(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }
    return this.excelService.closeWorkbook(filename);
  }

  private async handleGetCellInfo(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.readCell(filename, worksheet, cellAddress);
  }

  private async handleWriteBatch(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleInsertRows(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startRow = this.getNumberArg(args, 'startRow');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startRow) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.insertRows(filename, worksheet, startRow, count);
  }

  private async handleInsertColumns(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startColumn = this.getStringArg(args, 'startColumn');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startColumn) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.insertColumns(filename, worksheet, startColumn, count);
  }

  private async handleDeleteRows(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startRow = this.getNumberArg(args, 'startRow');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startRow) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteRows(filename, worksheet, startRow, count);
  }

  private async handleDeleteColumns(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startColumn = this.getStringArg(args, 'startColumn');
    const count = this.getNumberArg(args, 'count') || 1;

    if (!filename || !worksheet || !startColumn) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteColumns(filename, worksheet, startColumn, count);
  }

  private async handleMergeCells(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.mergeCells(filename, worksheet, startCell, endCell);
  }

  private async handleUnmergeCells(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.unmergeCells(filename, worksheet, cellAddress);
  }

  private async handleCopyRange(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleSetCellFormat(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');
    const format = args['format'] as Record<string, unknown>;

    if (!filename || !worksheet || !cellAddress || !format) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setCellFormat(filename, worksheet, cellAddress, format);
  }

  private async handleSetRangeFormat(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleAutoFitColumns(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startColumn = this.getStringArg(args, 'startColumn');
    const endColumn = this.getStringArg(args, 'endColumn');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.autoFitColumns(filename, worksheet, startColumn, endColumn);
  }

  private async handleSetColumnWidth(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const column = this.getStringArg(args, 'column');
    const width = this.getNumberArg(args, 'width');

    if (!filename || !worksheet || !column || !width) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setColumnWidth(filename, worksheet, column, width);
  }

  private async handleSetRowHeight(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const row = this.getNumberArg(args, 'row');
    const height = this.getNumberArg(args, 'height');

    if (!filename || !worksheet || !row || !height) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setRowHeight(filename, worksheet, row, height);
  }

  private async handleAddDataValidation(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleAddTable(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleFreezePanes(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.freezePanes(filename, worksheet, cellAddress);
  }

  private async handleAddFilter(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addFilter(filename, worksheet, startCell, endCell);
  }

  private async handleRemoveFilter(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.removeFilter(filename, worksheet);
  }

  private async handleDeleteWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.deleteWorksheet(filename, worksheet);
  }

  private async handleRenameWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const oldName = this.getStringArg(args, 'oldName');
    const newName = this.getStringArg(args, 'newName');

    if (!filename || !oldName || !newName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.renameWorksheet(filename, oldName, newName);
  }

  private async handleCopyWorksheet(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const sourceWorksheet = this.getStringArg(args, 'sourceWorksheet');
    const targetName = this.getStringArg(args, 'targetName');

    if (!filename || !sourceWorksheet || !targetName) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.copyWorksheet(filename, sourceWorksheet, targetName);
  }

  private async handleSetPrintArea(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');

    if (!filename || !worksheet || !startCell || !endCell) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.setPrintArea(filename, worksheet, startCell, endCell);
  }

  private async handleAddHeaderFooter(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const header = this.getStringArg(args, 'header');
    const footer = this.getStringArg(args, 'footer');

    if (!filename || !worksheet) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.addHeaderFooter(filename, worksheet, header, footer);
  }

  private async handleAddHyperlink(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleAddComment(args: Record<string, unknown>): Promise<OperationResult> {
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

  private async handleRemoveComment(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const cellAddress = this.getStringArg(args, 'cellAddress');

    if (!filename || !worksheet || !cellAddress) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.removeComment(filename, worksheet, cellAddress);
  }

  private async handleFindReplace(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const findText = this.getStringArg(args, 'findText');
    const replaceText = this.getStringArg(args, 'replaceText');
    const matchCase = args['matchCase'] as boolean;
    const matchEntireCell = args['matchEntireCell'] as boolean;

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

  private async handleSortRange(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');
    const worksheet = this.getStringArg(args, 'worksheet');
    const startCell = this.getStringArg(args, 'startCell');
    const endCell = this.getStringArg(args, 'endCell');
    const sortColumn = this.getNumberArg(args, 'sortColumn');
    const ascending = args['ascending'] as boolean ?? true;

    if (!filename || !worksheet || !startCell || !endCell || !sortColumn) {
      return { success: false, error: 'Missing required parameters' };
    }

    return this.excelService.sortRange(filename, worksheet, startCell, endCell, sortColumn, ascending);
  }

  private async handleCalculateFormula(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.calculateFormula(filename);
  }

  private async handleGetNamedRanges(args: Record<string, unknown>): Promise<OperationResult> {
    const filename = this.getStringArg(args, 'filename');

    if (!filename) {
      return { success: false, error: 'Missing required parameter: filename' };
    }

    return this.excelService.getNamedRanges(filename);
  }

  private async handleAddNamedRange(args: Record<string, unknown>): Promise<OperationResult> {
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

  // Helper methods
  private getStringArg(args: Record<string, unknown>, key: string): string | undefined {
    const value = args[key];
    return typeof value === 'string' ? value : undefined;
  }

  private getNumberArg(args: Record<string, unknown>, key: string): number | undefined {
    const value = args[key];
    return typeof value === 'number' ? value : undefined;
  }

  private parseCellRange(startCell: string, endCell: string): CellRange | null {
    const parseAddress = (addr: string): { row: number; column: number } | null => {
      const match = addr.match(/^([A-Z]+)(\d+)$/i);
      if (!match) return null;

      const colStr = match[1].toUpperCase();
      const row = parseInt(match[2], 10);

      let column = 0;
      for (let i = 0; i < colStr.length; i++) {
        column = column * 26 + (colStr.charCodeAt(i) - 64);
      }

      return { row, column };
    };

    const start = parseAddress(startCell);
    const end = parseAddress(endCell);

    if (!start || !end) return null;

    return { start, end };
  }
}