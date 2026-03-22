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

## Advanced Accounting Tools

### Financial Investment Analysis
- `excel_calculate_npv` - Calculate Net Present Value (NPV) for investment evaluation
  - Discount future cash flows to present value
  - Essential for investment decision-making
- `excel_calculate_irr` - Calculate Internal Rate of Return (IRR)
  - Find discount rate that makes NPV equal to zero
  - Expected annual return of an investment

### Financial Ratios
- `excel_calculate_financial_ratio` - Calculate key financial ratios:
  - **Current Ratio**: Measures ability to pay short-term obligations
  - **Quick Ratio**: Liquidity without inventory
  - **Debt-to-Equity**: Financial leverage
  - **Return on Equity (ROE)**: Profitability relative to equity
  - **Profit Margin**: Net profit as percentage of revenue

### Loan & Debt Management
- `excel_create_amortization_schedule` - Generate loan amortization table:
  - Payment breakdown (principal vs interest)
  - Remaining balance after each payment
  - Useful for mortgages, car loans, business loans
- `excel_create_aging_report` - Accounts receivable/payable aging:
  - Categorize by days past due: 0-30, 31-60, 61-90, 91-120, 120+ days
  - Track overdue invoices and payments

### Tax & Currency
- `excel_calculate_tax` - Calculate tax amounts at specified rates
  - Apply tax percentage to a range of values
  - Support VAT, GST, sales tax calculations
- `excel_convert_currency` - Convert currency amounts:
  - Apply exchange rate to convert values
  - Useful for multi-currency financial reporting

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

### Example 5: Investment Analysis (NPV & IRR)

```javascript
// Cash flow data in A1:A6: -10000, 3000, 4000, 3000, 2000, 1000
const npvResult = await excel_calculate_npv({
  filename: "investment_analysis",
  worksheet: "Project A",
  rate: 0.1, // 10% discount rate
  valuesRange: "A1:A6"
});

console.log(`NPV: $${npvResult.data?.npv.toFixed(2)}`);

// Calculate IRR
const irrResult = await excel_calculate_irr({
  filename: "investment_analysis",
  worksheet: "Project A",
  valuesRange: "A1:A6",
  guess: 0.1 // Initial guess
});

console.log(`IRR: ${(irrResult.data?.irr * 100).toFixed(2)}%`);
```

### Example 6: Loan Amortization Schedule

```javascript
// Create amortization schedule for a mortgage
await excel_create_amortization_schedule({
  filename: "loan",
  worksheet: "Mortgage",
  startCell: "A1",
  principal: 500000, // $500,000 loan
  annualRate: 0.05, // 5% annual interest
  numberOfPeriods: 360 // 30 years (360 months)
});

// Result: Table with columns:
// - Period: 1, 2, 3, ...
// - Payment: Monthly payment amount
// - Principal: Principal portion
// - Interest: Interest portion
// - Balance: Remaining balance
```

### Example 7: Financial Ratios Analysis

```javascript
// Calculate Current Ratio (Current Assets / Current Liabilities)
await excel_calculate_financial_ratio({
  filename: "financials",
  worksheet: "Balance Sheet",
  ratioType: "current",
  numeratorRange: "B2:B10", // Current Assets
  denominatorRange: "B15:B20" // Current Liabilities
});

// Calculate ROE (Net Income / Shareholder Equity)
await excel_calculate_financial_ratio({
  filename: "financials",
  worksheet: "Income Statement",
  ratioType: "return-on-equity",
  numeratorRange: "B30", // Net Income
  denominatorRange: "B45" // Shareholder Equity
});

// Calculate Profit Margin (Net Income / Revenue)
await excel_calculate_financial_ratio({
  filename: "financials",
  worksheet: "Income Statement",
  ratioType: "profit-margin",
  numeratorRange: "B30", // Net Income
  denominatorRange: "B5" // Revenue
});
```

### Example 8: Accounts Receivable Aging

```javascript
// Create aging report
await excel_create_aging_report({
  filename: "receivables",
  worksheet: "Invoices",
  invoiceDateColumn: "A", // Invoice dates
  amountColumn: "B", // Invoice amounts
  asOfDate: "2024-12-31", // As of December 31, 2024
  outputStartCell: "E1"
});

// Result: Aging summary table
// - 0-30 days: Total amount
// - 31-60 days: Total amount
// - 61-90 days: Total amount
// - 91-120 days: Total amount
// - 120+ days: Total amount
// - Total Outstanding: Sum of all categories
```

### Example 9: Tax Calculation

```javascript
// Calculate VAT (10%) for a range of prices
await excel_calculate_tax({
  filename: "sales",
  worksheet: "Products",
  amountRange: "B2:B100", // Prices
  taxRate: 10, // 10% VAT
  outputRange: "C2:C100" // Tax amounts
});

// Prices in B2:B100: 100, 200, 300...
// Tax in C2:C100: 10, 20, 30...
```

### Example 10: Currency Conversion

```javascript
// Convert USD to VND
await excel_convert_currency({
  filename: "financial_report",
  worksheet: "Summary",
  amountRange: "B2:B50", // USD amounts
  exchangeRate: 24000, // 1 USD = 24,000 VND
  outputRange: "C2:C50" // VND amounts
});

// USD in B2:B50: 100, 200, 300...
// VND in C2:C50: 2,400,000, 4,800,000, 7,200,000...
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