/**
 * Tool Handler - Processes MCP tool calls
 * Single Responsibility: Delegates to specialized handler modules
 */

import { ExcelService } from '../services/excel-service.js';
import { PermissionChecker } from '../security/permission-checker.js';
import { Logger } from '../utils/logger.js';
import { OperationResult } from '../types/index.js';
import { getToolDefinition } from './tool-definitions.js';
import {
  WorkbookHandlers,
  CellHandlers,
  WorksheetHandlers,
  FormattingHandlers,
  CommentDataHandlers,
  AdvancedDataHandlers,
  AnalysisHandlers,
  SystemHandlers
} from './handlers/index.js';

export class ToolHandler {
  private permissionChecker: PermissionChecker;
  private logger: Logger;
  private workbookHandlers: WorkbookHandlers;
  private cellHandlers: CellHandlers;
  private worksheetHandlers: WorksheetHandlers;
  private formattingHandlers: FormattingHandlers;
  private commentDataHandlers: CommentDataHandlers;
  private advancedDataHandlers: AdvancedDataHandlers;
  private analysisHandlers: AnalysisHandlers;
  private systemHandlers: SystemHandlers;

  constructor(
    excelService: ExcelService,
    permissionChecker: PermissionChecker,
    logger: Logger
  ) {
    this.permissionChecker = permissionChecker;
    this.logger = logger;
    this.workbookHandlers = new WorkbookHandlers(excelService);
    this.cellHandlers = new CellHandlers(excelService);
    this.worksheetHandlers = new WorksheetHandlers(excelService);
    this.formattingHandlers = new FormattingHandlers(excelService);
    this.commentDataHandlers = new CommentDataHandlers(excelService);
    this.advancedDataHandlers = new AdvancedDataHandlers(excelService);
    this.analysisHandlers = new AnalysisHandlers(excelService['activeWorkbooks']);
    this.systemHandlers = new SystemHandlers(excelService, permissionChecker, logger);
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

    // Execute tool by delegating to appropriate handler
    try {
      switch (toolName) {
        // Workbook operations
        case 'excel_open_workbook':
          return await this.workbookHandlers.handleOpenWorkbook(args);
        case 'excel_create_workbook':
          return await this.workbookHandlers.handleCreateWorkbook(args);
        case 'excel_save_workbook':
          return await this.workbookHandlers.handleSaveWorkbook(args);
        case 'excel_close_workbook':
          return await this.workbookHandlers.handleCloseWorkbook(args);

        // Cell operations
        case 'excel_read_cell':
          return await this.cellHandlers.handleReadCell(args);
        case 'excel_read_range':
          return await this.cellHandlers.handleReadRange(args);
        case 'excel_write_cell':
          return await this.cellHandlers.handleWriteCell(args);
        case 'excel_write_batch':
          return await this.cellHandlers.handleWriteBatch(args);
        case 'excel_get_cell_info':
          return await this.cellHandlers.handleGetCellInfo(args);
        case 'excel_copy_range':
          return await this.cellHandlers.handleCopyRange(args);
        case 'excel_find_replace':
          return await this.cellHandlers.handleFindReplace(args);
        case 'excel_sort_range':
          return await this.cellHandlers.handleSortRange(args);
        case 'excel_get_named_ranges':
          return await this.cellHandlers.handleGetNamedRanges(args);
        case 'excel_add_named_range':
          return await this.cellHandlers.handleAddNamedRange(args);
        case 'excel_calculate_formula':
          return await this.cellHandlers.handleCalculateFormula(args);

        // Worksheet operations
        case 'excel_list_worksheets':
          return await this.worksheetHandlers.handleListWorksheets(args);
        case 'excel_add_worksheet':
          return await this.worksheetHandlers.handleAddWorksheet(args);
        case 'excel_delete_worksheet':
          return await this.worksheetHandlers.handleDeleteWorksheet(args);
        case 'excel_rename_worksheet':
          return await this.worksheetHandlers.handleRenameWorksheet(args);
        case 'excel_copy_worksheet':
          return await this.worksheetHandlers.handleCopyWorksheet(args);
        case 'excel_insert_rows':
          return await this.worksheetHandlers.handleInsertRows(args);
        case 'excel_insert_columns':
          return await this.worksheetHandlers.handleInsertColumns(args);
        case 'excel_delete_rows':
          return await this.worksheetHandlers.handleDeleteRows(args);
        case 'excel_delete_columns':
          return await this.worksheetHandlers.handleDeleteColumns(args);
        case 'excel_merge_cells':
          return await this.worksheetHandlers.handleMergeCells(args);
        case 'excel_unmerge_cells':
          return await this.worksheetHandlers.handleUnmergeCells(args);
        case 'excel_add_table':
          return await this.worksheetHandlers.handleAddTable(args);
        case 'excel_add_filter':
          return await this.worksheetHandlers.handleAddFilter(args);
        case 'excel_remove_filter':
          return await this.worksheetHandlers.handleRemoveFilter(args);

        // Formatting operations
        case 'excel_set_cell_format':
          return await this.formattingHandlers.handleSetCellFormat(args);
        case 'excel_set_range_format':
          return await this.formattingHandlers.handleSetRangeFormat(args);
        case 'excel_auto_fit_columns':
          return await this.formattingHandlers.handleAutoFitColumns(args);
        case 'excel_set_column_width':
          return await this.formattingHandlers.handleSetColumnWidth(args);
        case 'excel_set_row_height':
          return await this.formattingHandlers.handleSetRowHeight(args);
        case 'excel_freeze_panes':
          return await this.formattingHandlers.handleFreezePanes(args);
        case 'excel_set_print_area':
          return await this.formattingHandlers.handleSetPrintArea(args);
        case 'excel_add_header_footer':
          return await this.formattingHandlers.handleAddHeaderFooter(args);

        // Comment and data operations
        case 'excel_add_comment':
          return await this.commentDataHandlers.handleAddComment(args);
        case 'excel_remove_comment':
          return await this.commentDataHandlers.handleRemoveComment(args);
        case 'excel_add_hyperlink':
          return await this.commentDataHandlers.handleAddHyperlink(args);
        case 'excel_add_data_validation':
          return await this.commentDataHandlers.handleAddDataValidation(args);

        // Advanced data operations
        case 'excel_find_duplicates':
          return await this.advancedDataHandlers.handleFindDuplicates(args);
        case 'excel_count_unique_values':
          return await this.advancedDataHandlers.handleCountUniqueValues(args);
        case 'excel_highlight_duplicates':
          return await this.advancedDataHandlers.handleHighlightDuplicates(args);
        case 'excel_get_duplicate_info':
          return await this.advancedDataHandlers.handleGetDuplicateInfo(args);

        // Analysis operations
        case 'excel_get_column_stats':
          return await this.analysisHandlers.getColumnStats(
            args.filename as string,
            args.worksheet as string,
            args.column as string,
            args.hasHeader as boolean ?? true
          );
        case 'excel_filter_data':
          return await this.analysisHandlers.filterData(
            args.filename as string,
            args.worksheet as string,
            args.startCell as string,
            args.endCell as string,
            args.filters as any[],
            args.hasHeader as boolean ?? true
          );
        case 'excel_group_aggregate':
          return await this.analysisHandlers.groupAggregate(
            args.filename as string,
            args.worksheet as string,
            args.startCell as string,
            args.endCell as string,
            args.groupByColumn as string,
            args.aggregateColumn as string,
            args.operation as string,
            args.hasHeader as boolean ?? true
          );
        case 'excel_profile_data':
          return await this.analysisHandlers.profileData(
            args.filename as string,
            args.worksheet as string,
            args.startCell as string,
            args.endCell as string,
            args.hasHeader as boolean ?? true
          );
        case 'excel_search':
          return await this.analysisHandlers.search(
            args.filename as string,
            args.worksheet as string,
            args.searchQuery as string,
            args.searchType as string ?? 'contains',
            args.columnRange as string ?? '',
            args.matchCase as boolean ?? false,
            args.maxResults as number ?? 100
          );
        case 'excel_compare_ranges':
          return await this.analysisHandlers.compareRanges(
            args.filename as string,
            args.worksheet1 as string,
            args.range1Start as string,
            args.range1End as string,
            args.worksheet2 as string ?? '',
            args.range2Start as string,
            args.range2End as string,
            args.compareType as string ?? 'values'
          );

        // System operations
        case 'excel_health_check':
          return await this.systemHandlers.healthCheck();

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
}