/**
 * Unit and Integration Tests for ExcelFormatting & Formatting Handlers
 * Testing all styling, formatting, layout, validation tools, and ToolHandler dispatch:
 * - Font: setFontStyle (bold, italic, underline, strike), setFontNameSize, setRichText
 * - Alignment: setAlignment (horizontal, vertical, wrapText), centerText
 * - Borders: setBorder, applyAllBorders, applyOutlineBorder
 * - Colors: setBackgroundColor, setFontColor, setGradientFill
 * - Number Formats: setNumberFormat, applyCurrencyFormat, applyPercentageFormat, applyDateFormat
 * - Presets: applyHeaderStyle, applyTitleStyle, applyTableStyle
 * - Dimensions & View: setColumnWidth, setRowHeight, autoFitColumns, freezePanes, setPrintArea, addHeaderFooter
 * - Interactive: addComment, removeComment, addHyperlink, addDataValidation
 */

import ExcelJS from 'exceljs';
import { ExcelFormatting } from '../src/services/excel-formatting.js';
import { ExcelService } from '../src/services/excel-service.js';
import { ToolHandler } from '../src/tools/tool-handler.js';
import { PermissionChecker } from '../src/security/permission-checker.js';
import { Logger } from '../src/utils/logger.js';
import { PermissionConfig } from '../src/types/index.js';

const config = (overrides: Partial<PermissionConfig> = {}): PermissionConfig => ({
  allowedPaths: [],
  deniedPaths: [],
  maxFileSize: 50 * 1024 * 1024,
  allowedExtensions: ['.xlsx', '.xls', '.xlsm', '.xlsb'],
  permissions: ['read', 'write', 'delete'],
  ...overrides,
});

describe('ExcelFormatting & Formatting Tools', () => {
  let workbooks: Map<string, ExcelJS.Workbook>;
  let formatting: ExcelFormatting;
  let excelService: ExcelService;
  let handler: ToolHandler;
  const fileName = 'test_format.xlsx';
  const sheetName = 'StyleSheet';

  beforeEach(() => {
    workbooks = new Map();
    const permChecker = new PermissionChecker(config());
    const logger = new Logger('silent');
    formatting = new ExcelFormatting(permChecker, logger, workbooks);
    excelService = new ExcelService(permChecker, logger);
    handler = new ToolHandler(excelService, permChecker, logger);

    const wb = new ExcelJS.Workbook();
    wb.addWorksheet(sheetName);
    workbooks.set(fileName, wb);
    excelService.getActiveWorkbooks().set(fileName, wb);
  });

  const getWs = (): ExcelJS.Worksheet => workbooks.get(fileName)!.getWorksheet(sheetName)!;

  describe('Font and Typography Styling', () => {
    it('sets font styles (bold, italic, underline, strike)', async () => {
      const res = await formatting.setFontStyle(fileName, sheetName, 'A1', 'A1', {
        bold: true,
        italic: true,
        underline: true,
        strikethrough: true,
      });
      expect(res.success).toBe(true);

      const font = getWs().getCell('A1').font;
      expect(font?.bold).toBe(true);
      expect(font?.italic).toBe(true);
      expect(font?.underline).toBe(true);
      expect(font?.strike).toBe(true);
    });

    it('sets font name and size', async () => {
      const res = await formatting.setFontNameSize(fileName, sheetName, 'B2', undefined, 'Arial', 18);
      expect(res.success).toBe(true);

      const font = getWs().getCell('B2').font;
      expect(font?.name).toBe('Arial');
      expect(font?.size).toBe(18);
    });

    it('sets rich text in a cell', async () => {
      const res = await formatting.setRichText(fileName, sheetName, 'C3', [
        { text: 'Prefix: ', bold: true },
        { text: 'Highlight', fontColor: 'FFFF0000', italic: true },
      ]);
      expect(res.success).toBe(true);

      const cellValue = getWs().getCell('C3').value as { richText: Array<{ text: string }> };
      expect(cellValue.richText).toHaveLength(2);
      expect(cellValue.richText[0].text).toBe('Prefix: ');
      expect(cellValue.richText[1].text).toBe('Highlight');
    });
  });

  describe('Alignment and Text Orientation', () => {
    it('sets horizontal, vertical alignment and wrapText', async () => {
      const res = await formatting.setAlignment(fileName, sheetName, 'A1', 'B2', {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      });
      expect(res.success).toBe(true);

      const alignment = getWs().getCell('A1').alignment;
      expect(alignment?.horizontal).toBe('center');
      expect(alignment?.vertical).toBe('middle');
      expect(alignment?.wrapText).toBe(true);
    });

    it('centers text in a cell or range', async () => {
      const res = await formatting.centerText(fileName, sheetName, 'D4');
      expect(res.success).toBe(true);

      const alignment = getWs().getCell('D4').alignment;
      expect(alignment?.horizontal).toBe('center');
      expect(alignment?.vertical).toBe('middle');
    });
  });

  describe('Borders and Fills', () => {
    it('sets specific borders on a range', async () => {
      const res = await formatting.setBorder(fileName, sheetName, 'B2', 'C3', {
        borderStyle: 'thick',
        borderColor: 'FF000000',
        top: true,
        bottom: true,
      });
      expect(res.success).toBe(true);

      const border = getWs().getCell('B2').border;
      expect(border?.top?.style).toBe('thick');
      expect(border?.bottom?.style).toBe('thick');
    });

    it('applies all borders across cells', async () => {
      const res = await formatting.applyAllBorders(fileName, sheetName, 'A1', 'B2', 'thin', 'FF333333');
      expect(res.success).toBe(true);

      const b = getWs().getCell('A1').border;
      expect(b?.top?.style).toBe('thin');
      expect(b?.bottom?.style).toBe('thin');
      expect(b?.left?.style).toBe('thin');
      expect(b?.right?.style).toBe('thin');
    });

    it('applies outline border around a multi-cell box', async () => {
      const res = await formatting.applyOutlineBorder(fileName, sheetName, 'A1', 'C3', 'medium', 'FF0000FF');
      expect(res.success).toBe(true);

      const topLeft = getWs().getCell('A1').border;
      expect(topLeft?.top?.style).toBe('medium');
      expect(topLeft?.left?.style).toBe('medium');

      const bottomRight = getWs().getCell('C3').border;
      expect(bottomRight?.bottom?.style).toBe('medium');
      expect(bottomRight?.right?.style).toBe('medium');
    });

    it('sets background color and font color', async () => {
      await formatting.setBackgroundColor(fileName, sheetName, 'E5', undefined, 'FFFFCC00');
      await formatting.setFontColor(fileName, sheetName, 'E5', undefined, 'FF000000');

      const fill = getWs().getCell('E5').fill as ExcelJS.FillPattern;
      expect(fill.type).toBe('pattern');
      expect(fill.fgColor?.argb).toBe('FFFFCC00');

      const font = getWs().getCell('E5').font;
      expect(font?.color?.argb).toBe('FF000000');
    });
  });

  describe('Number, Currency, Date, and Percentage Formats', () => {
    it('sets custom number format', async () => {
      const res = await formatting.setNumberFormat(fileName, sheetName, 'A1', 'A5', '#,##0.00');
      expect(res.success).toBe(true);
      expect(getWs().getCell('A1').numFmt).toBe('#,##0.00');
    });

    it('applies currency format with custom symbol', async () => {
      const res = await formatting.applyCurrencyFormat(fileName, sheetName, 'B1', 'B5', '€', 2);
      expect(res.success).toBe(true);
      expect(getWs().getCell('B1').numFmt).toBe('"€" #,##0.00');
    });

    it('applies percentage format', async () => {
      const res = await formatting.applyPercentageFormat(fileName, sheetName, 'C1', 'C5', 1);
      expect(res.success).toBe(true);
      expect(getWs().getCell('C1').numFmt).toBe('0.0%');
    });

    it('applies date formats (short, long, iso)', async () => {
      await formatting.applyDateFormat(fileName, sheetName, 'D1', undefined, 'short');
      expect(getWs().getCell('D1').numFmt).toBe('mm/dd/yyyy');

      await formatting.applyDateFormat(fileName, sheetName, 'D2', undefined, 'long');
      expect(getWs().getCell('D2').numFmt).toBe('dddd, mmmm dd, yyyy');

      await formatting.applyDateFormat(fileName, sheetName, 'D3', undefined, 'iso');
      expect(getWs().getCell('D3').numFmt).toBe('yyyy-mm-dd');
    });
  });

  describe('Style Presets (header, title, table)', () => {
    it('applies header style preset', async () => {
      const res = await formatting.applyHeaderStyle(fileName, sheetName, 'A1', 'E1', 'FF1F4E78', 'FFFFFFFF');
      expect(res.success).toBe(true);

      const cell = getWs().getCell('A1');
      expect(cell.font?.bold).toBe(true);
      expect(cell.font?.color?.argb).toBe('FFFFFFFF');
      expect((cell.fill as ExcelJS.FillPattern)?.fgColor?.argb).toBe('FF1F4E78');
    });

    it('applies title style preset', async () => {
      const res = await formatting.applyTitleStyle(fileName, sheetName, 'A1', 'E1', 20, 'FF2E75B6');
      expect(res.success).toBe(true);

      const cell = getWs().getCell('A1');
      expect(cell.font?.size).toBe(20);
      expect(cell.font?.bold).toBe(true);
      expect(cell.font?.color?.argb).toBe('FF2E75B6');
    });

    it('applies table style preset with alternating rows', async () => {
      const res = await formatting.applyTableStyle(
        fileName, sheetName, 'A1', 'C4',
        'FF4472C4', 'FFFFFFFF', 'FFF2F2F2', true
      );
      expect(res.success).toBe(true);

      // Header row
      expect(getWs().getCell('A1').font?.bold).toBe(true);
      // Alternate row (row 3 is even index after header row 1 -> row 2 vs row 3)
      const r2Fill = getWs().getCell('A2').fill as ExcelJS.FillPattern;
      const r3Fill = getWs().getCell('A3').fill as ExcelJS.FillPattern;
      expect(r2Fill || r3Fill).toBeDefined();
    });
  });

  describe('Sheet Dimensions, Freeze Panes, and Print Area', () => {
    it('sets column width and row height', async () => {
      await formatting.setColumnWidth(fileName, sheetName, 'B', 25);
      expect(getWs().getColumn('B').width).toBe(25);

      await formatting.setRowHeight(fileName, sheetName, 3, 40);
      expect(getWs().getRow(3).height).toBe(40);
    });

    it('auto-fits columns based on cell content length', async () => {
      getWs().getCell('A1').value = 'Very Long Header Text Exceeding Default';
      const res = await formatting.autoFitColumns(fileName, sheetName, 'A', 'A');
      expect(res.success).toBe(true);
      expect(getWs().getColumn('A').width).toBeGreaterThan(20);
    });

    it('freezes panes at a specified cell address', async () => {
      const res = await formatting.freezePanes(fileName, sheetName, 'B2');
      expect(res.success).toBe(true);

      const views = getWs().views;
      expect(views).toBeDefined();
      expect(views.length).toBeGreaterThan(0);
      expect(views[0].state).toBe('frozen');
    });

    it('sets print area and adds header/footer', async () => {
      const printRes = await formatting.setPrintArea(fileName, sheetName, 'A1', 'G50');
      expect(printRes.success).toBe(true);
      expect(getWs().pageSetup.printArea).toBe('A1:G50');

      const hfRes = await formatting.addHeaderFooter(fileName, sheetName, 'Confidential', 'Page 1');
      expect(hfRes.success).toBe(true);
      expect(getWs().headerFooter.oddHeader).toBe('Confidential');
      expect(getWs().headerFooter.oddFooter).toBe('Page 1');
    });
  });

  describe('Comments, Hyperlinks, and Data Validation', () => {
    it('adds and removes comments', async () => {
      const addRes = await formatting.addComment(fileName, sheetName, 'A1', 'Review needed', 'Auditor');
      expect(addRes.success).toBe(true);

      const cell = getWs().getCell('A1');
      expect(cell.note).toBeDefined();

      const removeRes = await formatting.removeComment(fileName, sheetName, 'A1');
      expect(removeRes.success).toBe(true);
      expect(cell.note).toBeUndefined();
    });

    it('adds a hyperlink to a cell', async () => {
      const res = await formatting.addHyperlink(fileName, sheetName, 'B2', 'https://example.com', 'Example Portal');
      expect(res.success).toBe(true);

      const cell = getWs().getCell('B2');
      const val = cell.value as { text: string; hyperlink: string };
      expect(val.text).toBe('Example Portal');
      expect(val.hyperlink).toBe('https://example.com');
    });

    it('adds list and range data validation', async () => {
      const listRes = await formatting.addDataValidation(
        fileName, sheetName, 'C3', 'list', '"Active,Inactive,Pending"', undefined, undefined, 'Must choose from list'
      );
      expect(listRes.success).toBe(true);

      const cell = getWs().getCell('C3');
      expect(cell.dataValidation).toBeDefined();
      expect(cell.dataValidation?.type).toBe('list');
    });
  });

  describe('ToolHandler Dispatch Integration', () => {
    it('dispatches font and alignment tools', async () => {
      const fontRes = await handler.executeTool('excel_set_font_style', {
        filename: fileName, worksheet: sheetName, startCell: 'A1',
        options: { bold: true },
      });
      expect(fontRes.success).toBe(true);

      const alignRes = await handler.executeTool('excel_set_alignment', {
        filename: fileName, worksheet: sheetName, startCell: 'A1',
        options: { horizontal: 'center' },
      });
      expect(alignRes.success).toBe(true);

      const centerRes = await handler.executeTool('excel_center_text', {
        filename: fileName, worksheet: sheetName, startCell: 'B2',
      });
      expect(centerRes.success).toBe(true);
    });

    it('dispatches border and color tools', async () => {
      expect((await handler.executeTool('excel_set_border', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'A2',
        options: { borderStyle: 'thin', top: true },
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_all_borders', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'B2',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_outline_border', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'B2',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_set_background_color', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', color: 'FFFFEEAA',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_set_font_color', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', color: 'FF333333',
      })).success).toBe(true);
    });

    it('dispatches format presets and number format tools', async () => {
      expect((await handler.executeTool('excel_set_number_format', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', format: '#,##0',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_header_style', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'D1',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_title_style', {
        filename: fileName, worksheet: sheetName, startCell: 'A1',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_currency_format', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', symbol: '$', decimalPlaces: 2,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_percentage_format', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', decimalPlaces: 1,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_apply_date_format', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', format: 'short',
      })).success).toBe(true);
    });

    it('dispatches comments, hyperlinks, and validation tools', async () => {
      expect((await handler.executeTool('excel_add_comment', {
        filename: fileName, worksheet: sheetName, cellAddress: 'F1', comment: 'Audit Check',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_remove_comment', {
        filename: fileName, worksheet: sheetName, cellAddress: 'F1',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_add_hyperlink', {
        filename: fileName, worksheet: sheetName, cellAddress: 'F2', url: 'https://example.com', display: 'Link',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_add_data_validation', {
        filename: fileName, worksheet: sheetName, cellAddress: 'F3', type: 'whole', formula1: '1', formula2: '100', operator: 'between',
      })).success).toBe(true);
    });

    it('dispatches dimension and view layout tools', async () => {
      expect((await handler.executeTool('excel_set_column_width', {
        filename: fileName, worksheet: sheetName, column: 'C', width: 20,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_set_row_height', {
        filename: fileName, worksheet: sheetName, row: 5, height: 35,
      })).success).toBe(true);

      expect((await handler.executeTool('excel_freeze_panes', {
        filename: fileName, worksheet: sheetName, cellAddress: 'B2',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_set_print_area', {
        filename: fileName, worksheet: sheetName, startCell: 'A1', endCell: 'E20',
      })).success).toBe(true);

      expect((await handler.executeTool('excel_add_header_footer', {
        filename: fileName, worksheet: sheetName, header: 'Top Title', footer: 'Bottom Note',
      })).success).toBe(true);
    });
  });
});
