# New Excel Formatting Tools

## Overview
This document describes the new formatting tools added to the Excel MCP server to make creating professional Excel files simpler and more efficient.

## New Tools Added

### Font Style Tools

#### 1. `excel_set_font_style`
Apply font styling (bold, italic, underline, strikethrough) to a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `bold` (optional): Bold text
- `italic` (optional): Italic text
- `underline` (optional): Underline style (true, "double", "singleAccounting", "doubleAccounting")
- `strikethrough` (optional): Strikethrough text

#### 2. `excel_set_font_name_size`
Set font family and size for a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `fontName` (optional): Font name (e.g., Arial, Calibri, Times New Roman)
- `fontSize` (optional): Font size in points

### Alignment Tools

#### 3. `excel_set_alignment`
Set horizontal and vertical alignment for a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `horizontal` (optional): Horizontal alignment (left, center, right, fill, justify, distributed)
- `vertical` (optional): Vertical alignment (top, middle, bottom, justify, distributed)
- `wrapText` (optional): Wrap text in cells
- `shrinkToFit` (optional): Shrink text to fit cell
- `indent` (optional): Text indent level (0-255)
- `textRotation` (optional): Text rotation angle in degrees (0-180)

#### 4. `excel_center_text`
Quickly center text horizontally and vertically in a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range

### Border Tools

#### 5. `excel_set_border`
Apply borders to a cell or range with customizable style and color.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `borderStyle` (optional): Border style (thin, medium, thick, double, dotted, dashed, etc.)
- `borderColor` (optional): Border color in hex format (e.g., FF000000 for black)
- `top` (optional): Apply top border
- `bottom` (optional): Apply bottom border
- `left` (optional): Apply left border
- `right` (optional): Apply right border
- `diagonal` (optional): Apply diagonal borders

#### 6. `excel_apply_all_borders`
Apply borders to all sides of a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `borderStyle` (optional): Border style (default: thin)
- `borderColor` (optional): Border color in hex (default: FF000000)

#### 7. `excel_apply_outline_border`
Apply border only to the outline of a range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `borderStyle` (optional): Border style (default: thin)
- `borderColor` (optional): Border color in hex (default: FF000000)

### Color Tools

#### 8. `excel_set_background_color`
Set background color for a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `color` (required): Background color in hex format (e.g., FFFF0000 for red)

#### 9. `excel_set_font_color`
Set font color for a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `color` (required): Font color in hex format (e.g., FF000000 for black)

### Number Format Tools

#### 10. `excel_set_number_format`
Set number format for a cell or range.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `format` (required): Number format code or preset (currency, percentage, number, date, time, datetime, fraction, scientific, text, or custom format)

### Style Presets

#### 11. `excel_apply_header_style`
Apply professional header style (bold, centered, gray background, bottom border).

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `backgroundColor` (optional): Background color in hex (default: FFD3D3D3)
- `fontColor` (optional): Font color in hex (default: FF000000)

#### 12. `excel_apply_title_style`
Apply professional title style (large, bold, centered, dark blue text).

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `fontSize` (optional): Font size in points (default: 16)
- `color` (optional): Font color in hex (default: FF1F4E78)

#### 13. `excel_apply_currency_format`
Apply currency formatting with 2 decimal places.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `symbol` (optional): Currency symbol (default: $)
- `decimalPlaces` (optional): Number of decimal places (default: 2)

#### 14. `excel_apply_percentage_format`
Apply percentage formatting.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `decimalPlaces` (optional): Number of decimal places (default: 2)

#### 15. `excel_apply_date_format`
Apply date formatting.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell (e.g., A1)
- `endCell` (optional): End cell for range
- `format` (optional): Date format preset (short, long, iso, or custom format like dd/mm/yyyy)

#### 16. `excel_apply_table_style`
Apply professional table formatting with headers and alternating row colors.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `startCell` (required): Start cell of table (e.g., A1)
- `endCell` (required): End cell of table (e.g., D10)
- `headerBackgroundColor` (optional): Header background color in hex (default: FFD3D3D3)
- `headerFontColor` (optional): Header font color in hex (default: FF000000)
- `alternateRowColor` (optional): Alternate row background color in hex (default: FFF9F9F9)
- `hasHeader` (optional): Whether first row is header (default: true)

### Rich Text

#### 17. `excel_set_rich_text`
Apply rich text formatting with multiple styles in a single cell.

**Parameters:**
- `filename` (required): Name of the opened workbook
- `worksheet` (required): Name of the worksheet
- `cellAddress` (required): Cell address (e.g., A1)
- `richText` (required): Array of text segments with formatting
  - `text` (required): Text segment
  - `bold` (optional): Bold text
  - `italic` (optional): Italic text
  - `underline` (optional): Underline text
  - `fontColor` (optional): Font color in hex
  - `fontSize` (optional): Font size
  - `fontName` (optional): Font name

## Usage Examples

### Example 1: Create a Professional Table
```typescript
// Create workbook and add data
await excel_create_workbook("report", "/path/to/output");
await excel_write_batch("report", "Sheet1", [
  { cellAddress: "A1", value: "Name" },
  { cellAddress: "B1", value: "Age" },
  { cellAddress: "C1", value: "Salary" },
  { cellAddress: "A2", value: "John" },
  { cellAddress: "B2", value: "30" },
  { cellAddress: "C2", value: "50000" },
]);

// Apply professional table style
await excel_apply_table_style("report", "Sheet1", "A1", "C2");
await excel_apply_all_borders("report", "Sheet1", "A1", "C2");
await excel_set_column_width("report", "Sheet1", "A", 15);
await excel_set_column_width("report", "Sheet1", "B", 10);
await excel_set_column_width("report", "Sheet1", "C", 15);
```

### Example 2: Apply Currency Formatting
```typescript
// Apply currency format to salary column
await excel_apply_currency_format("report", "Sheet1", "C2", "C2", "$", 2);
```

### Example 3: Create a Header Row
```typescript
// Apply header style
await excel_apply_header_style("report", "Sheet1", "A1", "C1", "FF4472C4", "FFFFFFFF");
```

### Example 4: Add a Title
```typescript
// Add title and style it
await excel_write_cell("report", "Sheet1", "A1", "Employee Report");
await excel_apply_title_style("report", "Sheet1", "A1", "C1", 18, "FF1F4E78");
await excel_merge_cells("report", "Sheet1", "A1", "C1");
await excel_center_text("report", "Sheet1", "A1", "C1");
```

### Example 5: Apply Rich Text
```typescript
await excel_set_rich_text("report", "Sheet1", "A1", [
  { text: "Important: ", bold: true, fontColor: "FFFF0000" },
  { text: "This is critical data", fontColor: "FF000000" }
]);
```

## Color Format Reference

Colors are specified in 8-character hex format: **AARRGGBB**
- **AA**: Alpha/Opacity (FF = opaque)
- **RR**: Red component (00-FF)
- **GG**: Green component (00-FF)
- **BB**: Blue component (00-FF)

Common colors:
- Black: FF000000
- White: FFFFFFFF
- Red: FFFF0000
- Green: FF00FF00
- Blue: FF0000FF
- Gray: FFD3D3D3
- Light Gray: FFF9F9F9
- Dark Blue: FF1F4E78

## Benefits

These new tools provide:

1. **Simplicity**: Easy-to-use presets for common formatting needs
2. **Consistency**: Professional-looking tables and reports
3. **Efficiency**: Apply multiple formatting options in single calls
4. **Flexibility**: Both granular control and quick presets available
5. **Professional Results**: Create publication-quality Excel files

## Migration Notes

The existing `excel_set_cell_format` and `excel_set_range_format` tools continue to work for backward compatibility. The new tools provide more focused and easier-to-use alternatives for common formatting tasks.