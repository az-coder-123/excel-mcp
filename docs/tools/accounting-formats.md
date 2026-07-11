# Accounting Formats

Tools for professional accounting number formatting.

## `excel_accounting_format`

Apply professional accounting number format with configurable options.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `decimalPlaces` | ❌ | Number of decimal places |
| `useSeparator` | ❌ | Use thousand separators |
| `showNegativeInRed` | ❌ | Display negative numbers in red |
| `showNegativeInParentheses` | ❌ | Display negative numbers in parentheses |
| `currencySymbol` | ❌ | Currency symbol |

## `excel_vnd_currency_format`

Format values as Vietnamese Dong (₫) with optional decimal places.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `decimalPlaces` | ❌ | Number of decimal places |

## `excel_negative_red_format`

Display negative numbers in red.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `decimalPlaces` | ❌ | Number of decimal places |

## `excel_show_zeros_instead_of_empty`

Replace empty cells with zeros.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |

---

← [Back to index](README.md)