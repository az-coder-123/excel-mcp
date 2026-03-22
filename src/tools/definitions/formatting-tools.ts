  /**
   * Formatting Tool Definitions
   * Single Responsibility: Cell and range formatting operations
   */

  import { ToolDefinition } from '../../types/index.js';

  export const FORMATTING_TOOLS: ToolDefinition[] = [
  // Font Style Tools
  {
    name: 'excel_set_font_style',
    description: 'Apply font styling (bold, italic, underline, strikethrough) to a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'bold',
        type: 'boolean',
        description: 'Bold text',
        required: false,
      },
      {
        name: 'italic',
        type: 'boolean',
        description: 'Italic text',
        required: false,
      },
      {
        name: 'underline',
        type: 'string',
        description: 'Underline style (true for single, "double", "singleAccounting", "doubleAccounting")',
        required: false,
      },
      {
        name: 'strikethrough',
        type: 'boolean',
        description: 'Strikethrough text',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_font_name_size',
    description: 'Set font family and size for a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'fontName',
        type: 'string',
        description: 'Font name (e.g., Arial, Calibri, Times New Roman)',
        required: false,
      },
      {
        name: 'fontSize',
        type: 'number',
        description: 'Font size in points',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Alignment Tools
  {
    name: 'excel_set_alignment',
    description: 'Set horizontal and vertical alignment for a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'horizontal',
        type: 'string',
        description: 'Horizontal alignment (left, center, right, fill, justify, distributed)',
        required: false,
      },
      {
        name: 'vertical',
        type: 'string',
        description: 'Vertical alignment (top, middle, bottom, justify, distributed)',
        required: false,
      },
      {
        name: 'wrapText',
        type: 'boolean',
        description: 'Wrap text in cells',
        required: false,
      },
      {
        name: 'shrinkToFit',
        type: 'boolean',
        description: 'Shrink text to fit cell',
        required: false,
      },
      {
        name: 'indent',
        type: 'number',
        description: 'Text indent level (0-255)',
        required: false,
      },
      {
        name: 'textRotation',
        type: 'number',
        description: 'Text rotation angle in degrees (0-180)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_center_text',
    description: 'Quickly center text horizontally and vertically in a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Border Tools
  {
    name: 'excel_set_border',
    description: 'Apply borders to a cell or range with customizable style and color',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'borderStyle',
        type: 'string',
        description: 'Border style (thin, medium, thick, double, dotted, dashed, hair, mediumDashed, dashDot, mediumDashDot, dashDotDot, mediumDashDotDot, slantDashDot)',
        required: false,
      },
      {
        name: 'borderColor',
        type: 'string',
        description: 'Border color in hex format (e.g., FF000000 for black)',
        required: false,
      },
      {
        name: 'top',
        type: 'boolean',
        description: 'Apply top border',
        required: false,
      },
      {
        name: 'bottom',
        type: 'boolean',
        description: 'Apply bottom border',
        required: false,
      },
      {
        name: 'left',
        type: 'boolean',
        description: 'Apply left border',
        required: false,
      },
      {
        name: 'right',
        type: 'boolean',
        description: 'Apply right border',
        required: false,
      },
      {
        name: 'diagonal',
        type: 'boolean',
        description: 'Apply diagonal borders',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_all_borders',
    description: 'Apply borders to all sides of a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'borderStyle',
        type: 'string',
        description: 'Border style (default: thin)',
        required: false,
      },
      {
        name: 'borderColor',
        type: 'string',
        description: 'Border color in hex (default: FF000000)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_outline_border',
    description: 'Apply border only to the outline of a range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'borderStyle',
        type: 'string',
        description: 'Border style (default: thin)',
        required: false,
      },
      {
        name: 'borderColor',
        type: 'string',
        description: 'Border color in hex (default: FF000000)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Fill/Background Tools
  {
    name: 'excel_set_background_color',
    description: 'Set background color for a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Background color in hex format (e.g., FFFF0000 for red)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_gradient_fill',
    description: 'Apply gradient fill to a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'color1',
        type: 'string',
        description: 'Start color in hex format',
        required: true,
      },
      {
        name: 'color2',
        type: 'string',
        description: 'End color in hex format',
        required: true,
      },
      {
        name: 'type',
        type: 'string',
        description: 'Gradient type (horizontal, vertical, diagonal)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Font Color Tools
  {
    name: 'excel_set_font_color',
    description: 'Set font color for a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Font color in hex format (e.g., FF000000 for black)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Number Format Tools
  {
    name: 'excel_set_number_format',
    description: 'Set number format for a cell or range',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'format',
        type: 'string',
        description: 'Number format code or preset (currency, percentage, number, date, time, datetime, fraction, scientific, text, or custom format)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Style Presets
  {
    name: 'excel_apply_header_style',
    description: 'Apply professional header style (bold, centered, gray background, bottom border)',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range',
        required: false,
      },
      {
        name: 'backgroundColor',
        type: 'string',
        description: 'Background color in hex (default: FFD3D3D3)',
        required: false,
      },
      {
        name: 'fontColor',
        type: 'string',
        description: 'Font color in hex (default: FF000000)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_title_style',
    description: 'Apply professional title style (large, bold, centered, dark blue text)',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'fontSize',
        type: 'number',
        description: 'Font size in points (default: 16)',
        required: false,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Font color in hex (default: FF1F4E78)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_currency_format',
    description: 'Apply currency formatting with 2 decimal places',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'symbol',
        type: 'string',
        description: 'Currency symbol (default: $)',
        required: false,
      },
      {
        name: 'decimalPlaces',
        type: 'number',
        description: 'Number of decimal places (default: 2)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_percentage_format',
    description: 'Apply percentage formatting',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'decimalPlaces',
        type: 'number',
        description: 'Number of decimal places (default: 2)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_date_format',
    description: 'Apply date formatting',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell for range (omit for single cell)',
        required: false,
      },
      {
        name: 'format',
        type: 'string',
        description: 'Date format preset (short, long, iso, or custom format like dd/mm/yyyy)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_table_style',
    description: 'Apply professional table formatting with headers and alternating row colors',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell of table (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of table (e.g., D10)',
        required: true,
      },
      {
        name: 'headerBackgroundColor',
        type: 'string',
        description: 'Header background color in hex (default: FFD3D3D3)',
        required: false,
      },
      {
        name: 'headerFontColor',
        type: 'string',
        description: 'Header font color in hex (default: FF000000)',
        required: false,
      },
      {
        name: 'alternateRowColor',
        type: 'string',
        description: 'Alternate row background color in hex (default: FFF9F9F9)',
        required: false,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Whether first row is header (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Rich Text
  {
    name: 'excel_set_rich_text',
    description: 'Apply rich text formatting with multiple styles in a single cell',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1)',
        required: true,
      },
      {
        name: 'richText',
        type: 'array',
        description: 'Array of text segments with formatting',
        required: true,
        properties: {
          text: { type: 'string', description: 'Text segment' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'boolean' },
          fontColor: { type: 'string', description: 'Font color in hex' },
          fontSize: { type: 'number', description: 'Font size' },
          fontName: { type: 'string', description: 'Font name' }
        }
      },
    ],
    requiredPermissions: ['write'],
  },
  // Existing tools continue...
  {
    name: 'excel_set_cell_format',
    description: 'Apply formatting to a cell (font, color, borders, alignment)',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell address (e.g., A1)',
        required: true,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Formatting options object',
        required: true,
        properties: {
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'boolean' },
          fontSize: { type: 'number' },
          fontColor: { type: 'string' },
          backgroundColor: { type: 'string' },
          borderColor: { type: 'string' },
          borderStyle: { type: 'string' },
          horizontalAlignment: { type: 'string' },
          verticalAlignment: { type: 'string' },
          wrapText: { type: 'boolean' },
          numberFormat: { type: 'string' }
        }
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_range_format',
    description: 'Apply formatting to a range of cells',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell of range (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of range (e.g., C10)',
        required: true,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Formatting options object',
        required: true,
        properties: {
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          fontSize: { type: 'number' },
          fontColor: { type: 'string' },
          backgroundColor: { type: 'string' },
          numberFormat: { type: 'string' }
        }
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_auto_fit_columns',
    description: 'Auto-fit column widths to content',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startColumn',
        type: 'string',
        description: 'Start column (e.g., A)',
        required: false,
      },
      {
        name: 'endColumn',
        type: 'string',
        description: 'End column (e.g., Z)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_column_width',
    description: 'Set width for a specific column',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'column',
        type: 'string',
        description: 'Column letter (e.g., A)',
        required: true,
      },
      {
        name: 'width',
        type: 'number',
        description: 'Column width',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_row_height',
    description: 'Set height for a specific row',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'row',
        type: 'number',
        description: 'Row number (1-based)',
        required: true,
      },
      {
        name: 'height',
        type: 'number',
        description: 'Row height',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_freeze_panes',
    description: 'Freeze panes for easier navigation',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Cell to freeze at (e.g., B2 freezes row 1 and column A)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_print_area',
    description: 'Set the print area for a worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell of print area (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of print area (e.g., G50)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_header_footer',
    description: 'Add header and footer to worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'header',
        type: 'string',
        description: 'Header text (use &[Page], &[Pages], &[Date], &[Time], &[File])',
        required: false,
      },
      {
        name: 'footer',
        type: 'string',
        description: 'Footer text',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
];