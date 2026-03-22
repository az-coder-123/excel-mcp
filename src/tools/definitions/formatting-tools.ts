  /**
   * Formatting Tool Definitions
   * Single Responsibility: Cell and range formatting operations
   */

  import { ToolDefinition } from '../../types/index.js';

  export const FORMATTING_TOOLS: ToolDefinition[] = [
  // Font Style Tools
  {
    name: 'excel_set_font_style',
    description: 'Apply font style (bold, italic, underline, strikethrough)',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'bold',
        type: 'boolean',
        description: 'Bold',
        required: false,
      },
      {
        name: 'italic',
        type: 'boolean',
        description: 'Italic',
        required: false,
      },
      {
        name: 'underline',
        type: 'string',
        description: 'Underline (true/double/singleAccounting/doubleAccounting)',
        required: false,
      },
      {
        name: 'strikethrough',
        type: 'boolean',
        description: 'Strikethrough',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_font_name_size',
    description: 'Set font name and size',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'fontName',
        type: 'string',
        description: 'Font name (e.g., Arial, Calibri)',
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
    description: 'Set cell alignment',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'horizontal',
        type: 'string',
        description: 'Horizontal (left/center/right/fill/justify/distributed)',
        required: false,
      },
      {
        name: 'vertical',
        type: 'string',
        description: 'Vertical (top/middle/bottom/justify/distributed)',
        required: false,
      },
      {
        name: 'wrapText',
        type: 'boolean',
        description: 'Wrap text',
        required: false,
      },
      {
        name: 'shrinkToFit',
        type: 'boolean',
        description: 'Shrink to fit',
        required: false,
      },
      {
        name: 'indent',
        type: 'number',
        description: 'Indent level (0-255)',
        required: false,
      },
      {
        name: 'textRotation',
        type: 'number',
        description: 'Rotation angle (0-180)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_center_text',
    description: 'Center text horizontally and vertically',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Border Tools
  {
    name: 'excel_set_border',
    description: 'Apply custom borders',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'borderStyle',
        type: 'string',
        description: 'Style (thin/medium/thick/double/dotted/dashed)',
        required: false,
      },
      {
        name: 'borderColor',
        type: 'string',
        description: 'Color hex (e.g., FF000000)',
        required: false,
      },
      {
        name: 'top',
        type: 'boolean',
        description: 'Top border',
        required: false,
      },
      {
        name: 'bottom',
        type: 'boolean',
        description: 'Bottom border',
        required: false,
      },
      {
        name: 'left',
        type: 'boolean',
        description: 'Left border',
        required: false,
      },
      {
        name: 'right',
        type: 'boolean',
        description: 'Right border',
        required: false,
      },
      {
        name: 'diagonal',
        type: 'boolean',
        description: 'Diagonal borders',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_all_borders',
    description: 'Apply borders all sides',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'borderStyle',
        type: 'string',
        description: 'Style (default: thin)',
        required: false,
      },
      {
        name: 'borderColor',
        type: 'string',
        description: 'Color hex (default: FF000000)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_outline_border',
    description: 'Apply outline border only',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'borderStyle',
        type: 'string',
        description: 'Style (default: thin)',
        required: false,
      },
      {
        name: 'borderColor',
        type: 'string',
        description: 'Color hex (default: FF000000)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Fill/Background Tools
  {
    name: 'excel_set_background_color',
    description: 'Set background color',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Color hex (e.g., FFFF0000)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_gradient_fill',
    description: 'Apply gradient fill',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'color1',
        type: 'string',
        description: 'Start color hex',
        required: true,
      },
      {
        name: 'color2',
        type: 'string',
        description: 'End color hex',
        required: true,
      },
      {
        name: 'type',
        type: 'string',
        description: 'Type (horizontal/vertical/diagonal)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Font Color Tools
  {
    name: 'excel_set_font_color',
    description: 'Set font color',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Color hex (e.g., FF000000)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Number Format Tools
  {
    name: 'excel_set_number_format',
    description: 'Set number format',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'format',
        type: 'string',
        description: 'Format (currency/percentage/number/date/time or custom)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Style Presets
  {
    name: 'excel_apply_header_style',
    description: 'Apply header style preset',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'backgroundColor',
        type: 'string',
        description: 'Background hex (default: FFD3D3D3)',
        required: false,
      },
      {
        name: 'fontColor',
        type: 'string',
        description: 'Font color hex (default: FF000000)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_title_style',
    description: 'Apply title style preset',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'fontSize',
        type: 'number',
        description: 'Size (default: 16)',
        required: false,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Color hex (default: FF1F4E78)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_currency_format',
    description: 'Apply currency format',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'symbol',
        type: 'string',
        description: 'Symbol (default: $)',
        required: false,
      },
      {
        name: 'decimalPlaces',
        type: 'number',
        description: 'Decimals (default: 2)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_percentage_format',
    description: 'Apply percentage format',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'decimalPlaces',
        type: 'number',
        description: 'Decimals (default: 2)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_date_format',
    description: 'Apply date format',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'End cell (optional)',
        required: false,
      },
      {
        name: 'format',
        type: 'string',
        description: 'Format (short/long/iso or custom)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_apply_table_style',
    description: 'Apply table style preset',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Table start (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Table end (e.g., D10)',
        required: true,
      },
      {
        name: 'headerBackgroundColor',
        type: 'string',
        description: 'Header bg hex (default: FFD3D3D3)',
        required: false,
      },
      {
        name: 'headerFontColor',
        type: 'string',
        description: 'Header font hex (default: FF000000)',
        required: false,
      },
      {
        name: 'alternateRowColor',
        type: 'string',
        description: 'Alt row hex (default: FFF9F9F9)',
        required: false,
      },
      {
        name: 'hasHeader',
        type: 'boolean',
        description: 'Has header (default: true)',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  // Rich Text
  {
    name: 'excel_set_rich_text',
    description: 'Apply rich text formatting',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'Text segments with formatting',
        required: true,
        properties: {
          text: { type: 'string', description: 'Text' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'boolean' },
          fontColor: { type: 'string', description: 'Color hex' },
          fontSize: { type: 'number', description: 'Size' },
          fontName: { type: 'string', description: 'Font' }
        }
      },
    ],
    requiredPermissions: ['write'],
  },
  // Existing tools continue...
  {
    name: 'excel_set_cell_format',
    description: 'Apply cell formatting',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'Format options',
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
    description: 'Apply range formatting',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Range start (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end (e.g., C10)',
        required: true,
      },
      {
        name: 'format',
        type: 'object',
        description: 'Format options',
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
    description: 'Auto-fit columns',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
    description: 'Set column width',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'column',
        type: 'string',
        description: 'Column (e.g., A)',
        required: true,
      },
      {
        name: 'width',
        type: 'number',
        description: 'Width',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_row_height',
    description: 'Set row height',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
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
        description: 'Height',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_freeze_panes',
    description: 'Freeze panes',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'cellAddress',
        type: 'string',
        description: 'Freeze at cell (e.g., B2)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_set_print_area',
    description: 'Set print area',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Area start (e.g., A1)',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Area end (e.g., G50)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_add_header_footer',
    description: 'Add header/footer',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Opened workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'header',
        type: 'string',
        description: 'Header (use &[Page], &[Date], etc.)',
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