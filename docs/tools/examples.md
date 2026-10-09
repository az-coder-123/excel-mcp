# Usage Examples

Practical examples covering formatting, accounting, and financial analysis tools.

## Example 1: Professional Report

Create a workbook with styled headers, a title, and a table with alternating row colors and borders.

```javascript
// Create workbook
await excel_create_workbook({
  filename: "report",
  outputPath: "/path/to/output"
});

// Add title
await excel_apply_title_style({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "A1"
});

// Add headers with style
await excel_apply_header_style({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A2",
  endCell: "D2",
  backgroundColor: "4472C4",
  fontColor: "FFFFFF"
});

// Add table with alternating colors
await excel_apply_table_style({
  filename: "report",
  worksheet: "Sheet1",
  startCell: "A3",
  endCell: "D20",
  headerBackgroundColor: "4472C4",
  headerFontColor: "FFFFFF",
  alternateRowColor: "E7E6E6",
  hasHeader: false
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

## Example 2: Financial Report with VND Format

Calculate running totals and apply VND currency + negative-in-red formatting.

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

## Example 3: Budget vs Actual Analysis

Calculate variance, apply accounting format, and verify balance.

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

## Example 4: Rich Text Formatting

Create a cell with multiple text styles in a single cell.

```javascript
await excel_set_rich_text({
  filename: "document",
  worksheet: "Sheet1",
  cellAddress: "A1",
  richText: [
    { text: "Important Notice: ", bold: true, fontColor: "FF0000" },
    { text: "This is urgent.", italic: true, underline: true },
    { text: " Contact support for help.", fontColor: "0000FF" }
  ]
});
```

## Example 5: Investment Analysis (NPV & IRR)

Evaluate an investment using NPV and IRR.

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
  guess: 0.1
});

console.log(`IRR: ${(irrResult.data?.irr * 100).toFixed(2)}%`);
```

## Example 6: Loan Amortization Schedule

Generate a mortgage amortization table.

```javascript
await excel_create_amortization_schedule({
  filename: "loan",
  worksheet: "Mortgage",
  startCell: "A1",
  principal: 500000,     // $500,000 loan
  annualRate: 0.05,      // 5% annual interest
  numberOfPeriods: 360   // 30 years (360 months)
});

// Result: Table with columns:
// Period | Payment | Principal | Interest | Balance
```

## Example 7: Financial Ratios Analysis

Calculate key financial ratios from balance sheet and income statement data.

```javascript
// Current Ratio (Current Assets / Current Liabilities)
await excel_calculate_financial_ratio({
  filename: "financials",
  worksheet: "Balance Sheet",
  ratioType: "current",
  numeratorRange: "B2:B10",   // Current Assets
  denominatorRange: "B15:B20" // Current Liabilities
});

// ROE (Net Income / Shareholder Equity)
await excel_calculate_financial_ratio({
  filename: "financials",
  worksheet: "Income Statement",
  ratioType: "return-on-equity",
  numeratorRange: "B30", // Net Income
  denominatorRange: "B45" // Shareholder Equity
});

// Profit Margin (Net Income / Revenue)
await excel_calculate_financial_ratio({
  filename: "financials",
  worksheet: "Income Statement",
  ratioType: "profit-margin",
  numeratorRange: "B30", // Net Income
  denominatorRange: "B5"  // Revenue
});
```

## Example 8: Accounts Receivable Aging

Categorize outstanding invoices by age.

```javascript
await excel_create_aging_report({
  filename: "receivables",
  worksheet: "Invoices",
  invoiceDateColumn: "A",  // Invoice dates
  amountColumn: "B",       // Invoice amounts
  asOfDate: "2024-12-31",  // As of December 31, 2024
  outputStartCell: "E1"
});

// Result: Aging summary table
// 0-30 days | 31-60 days | 61-90 days | 91-120 days | 120+ days | Total
```

## Example 9: Tax Calculation

Calculate VAT for a range of prices.

```javascript
await excel_calculate_tax({
  filename: "sales",
  worksheet: "Products",
  amountRange: "B2:B100", // Prices
  taxRate: 10,            // 10% VAT
  outputRange: "C2:C100"  // Tax amounts
});

// Prices: 100, 200, 300... → Tax: 10, 20, 30...
```

## Example 10: Currency Conversion

Convert USD to VND.

```javascript
await excel_convert_currency({
  filename: "financial_report",
  worksheet: "Summary",
  amountRange: "B2:B50",   // USD amounts
  exchangeRate: 24000,     // 1 USD = 24,000 VND
  outputRange: "C2:C50"    // VND amounts
});

// USD: 100, 200, 300... → VND: 2,400,000, 4,800,000, 7,200,000...
```

---

← [Back to index](README.md)