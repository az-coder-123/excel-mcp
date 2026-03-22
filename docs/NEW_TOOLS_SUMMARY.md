# New Tools Added Summary

This document summarizes the new formatting and accounting tools added to the Excel MCP server.

## Formatting Tools

### Font & Text Styling
- `excel_set_font_style` - Apply bold, italic, underline, strikethrough to cells/range
- `excel_set_font_name_size` - Change font family and size
- `excel_set_rich_text` - Apply multiple styles within a single cell

### Alignment & Text Control
- `excel_set_alignment` - Set horizontal, vertical alignment, wrap text, indent, text rotation
- `excel_center_text` - Quick centering (horizontal & vertical)

### Borders & Borders Styling
- `excel_set_border` - Apply borders to specific sides (top, bottom, left, right, diagonal)
- `excel_apply_all_borders` - Apply borders to all sides of a range
- `excel_apply_outline_border` - Apply borders only to outer edges of a range

### Colors & Backgrounds
- `excel_set_background_color` - Set fill color for cells/range
- `excel_set_font_color` - Set text color

### Number Formats
- `excel_set_number_format` - Apply custom number format
- `excel_apply_currency_format` - Format as currency with symbol
- `excel_apply_percentage_format` - Format as percentage
- `excel_apply_date_format` - Format as date

### Pre-built Styles
- `excel_apply_header_style` - Apply professional header styling with customizable colors
- `excel_apply_title_style` - Apply bold, large title styling
- `excel_apply_table_style` - Apply alternating row colors and header styling

## Accounting Tools

### Financial Calculations
- `excel_financial_sum` - Calculate sum with optional criteria (SUMIF-like)
- `excel_financial_average` - Calculate average of values in range
- `excel_running_total` - Calculate cumulative running totals
- `excel_percentage_of_total` - Calculate each value as percentage of total
- `excel_year_to_date` - Calculate year-to-date cumulative values

### Accounting Formats
- `excel_accounting_format` - Apply professional accounting number format with options for:
  - Decimal places
  - Thousand separators
  - Negative in parentheses
  - Negative in red
  - Currency symbol
- `excel_vnd_currency_format` - Format as Vietnamese Dong (đ) with optional decimals
- `excel_negative_red_format` - Display negative numbers in red color
- `excel_show_zeros_instead_of_empty` - Replace empty cells with zeros

### Financial Analysis
- `excel_period_comparison` - Compare current vs previous period values
- `excel_variance_analysis` - Calculate budget vs actual variance
- `excel_check_balance` - Verify debits equal credits (accounting balance check)
- `excel_find_anomalies` - Identify outliers using standard deviation

## Usage Examples

### Example 1: Create Professional Report

```javascript
// Create workbook
await excel_create_workbook({
  filename: "report",
  outputPath: "/path/to/output"
});

// Add headers with style
await excel_apply_header_style({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "D1",
  backgroundColor: "4472C4",
  fontColor: "FFFFFF"
});

// Add title
await excel_apply_title_style({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "A1"
});

// Add table with alternating colors
await excel_apply_table_style({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A2",
  endCell: "D20",
  headerBackgroundColor: "4472C4",
  headerFontColor: "FFFFFF",
  alternateRowColor: "E7E6E6",
  hasHeader: true
});

// Add all borders
await excel_apply_all_borders({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A2",
  endCell: "D20",
  borderStyle: "thin",
  borderColor: "000000"
});

// Save
await excel_save_workbook({ filename: "report" });
```

### Example 2: Financial Report with VND Format

```javascript
// Calculate running total
await excel_running_total({
  filename: "financial_report",
  worksheet: "Data",
  valueStartCell: "C2",
  valueEndCell: "C100",
  outputStartCell: "D2"
});

// Apply VND currency format
await excel_vnd_currency_format({
  filename: "financial_report",
  worksheet: "Data",
  startCell: "C2",
  endCell: "D100",
  decimalPlaces: 0
});

// Apply negative red format
await excel_negative_red_format({
  filename: "financial_report",
  worksheet: "Data",
  startCell: "E2",
  endCell: "E100",
  decimalPlaces: 2
});

// Calculate period comparison
await excel_period_comparison({
  filename: "financial_report",
  worksheet: "Data",
  currentValueRange: "C2:C100",
  previousValueRange: "B2:B100",
  outputRange: "F2",
  showPercentage: true
});
```

### Example 3: Budget vs Actual Analysis

```javascript
// Calculate variance
await excel_variance_analysis({
  filename: "budget_report",
  worksheet: "Summary",
  budgetRange: "B2:B50",
  actualRange: "C2:C50",
  outputRange: "D2",
  showPercentage: true
});

// Apply accounting format
await excel_accounting_format({
  filename: "budget_report",
  worksheet: "Summary",
  startCell: "B2",
  endCell: "D50",
  decimalPlaces: 2,
  useSeparator: true,
  showNegativeInParentheses: true,
  showNegativeInRed: true,
  currencySymbol: "$"
});

// Check if balanced
const balanceCheck = await excel_check_balance({
  filename: "budget_report",
  worksheet: "Summary",
  debitRange: "B2:B50",
  creditRange: "C2:C50"
});

if (balanceCheck.data?.isBalanced) {
  console.log("Report is balanced!");
} else {
  console.log(`Difference: ${balanceCheck.data?.difference}`);
}
```

### Example 4: Rich Text Formatting

```javascript
// Create rich text cell with multiple styles
await excel_set_rich_text({
  filename: "document",
  worksheet: "Sheet1",
  cellAddress: "A1",
  richText: [
    { 
      text: "Important Notice: ", 
      bold: true, 
      fontColor: "FF0000" 
    },
    { 
      text: "This is urgent.", 
      italic: true,
      underline: true
    },
    {
      text: " Contact support for help.",
      fontColor: "0000FF"
    }
  ]
});
```

## Color Codes

Colors should be specified in hex format (e.g., "FF0000" for red, "00FF00" for green):

- **Blue**: "4472C4", "1F4E78"
- **Red**: "FF0000", "C00000"
- **Green**: "00B050", "006100"
- **Yellow**: "FFC000", "FFC000"
- **Orange**: "FF9900", "ED7D31"
- **Gray**: "E7E6E6", "D9D9D9"
- **Black**: "000000"
- **White**: "FFFFFF"

## Border Styles

- `thin` - Thin line
- `medium` - Medium line
- `thick` - Thick line
- `double` - Double line
- `dotted` - Dotted line
- `dashed` - Dashed line
- `hair` - Hairline

## Font Styles

Common font names:
- Arial
- Calibri
- Times New Roman
- Verdana
- Georgia

Font sizes are specified in points (e.g., 10, 11, 12, 14, 16, 18)

## Integration Notes

All new tools:
- Follow the existing MCP tool architecture
- Include proper error handling and validation
- Support both single cells and ranges
- Return consistent OperationResult objects
- Include detailed logging for debugging

The tools are now registered and ready to use through the MCP server!