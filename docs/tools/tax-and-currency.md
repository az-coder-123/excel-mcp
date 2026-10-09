# Tax & Currency

Tools for calculating tax amounts and converting currencies.

## `excel_calculate_tax`

Calculate tax amounts at a specified rate. Supports VAT, GST, and sales tax calculations.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `amountRange` | ✅ | Range of amounts to tax |
| `taxRate` | ✅ | Tax rate in percentage (e.g., `10` for 10%) |
| `outputRange` | ✅ | Output range for tax amounts |

## `excel_convert_currency`

Convert currency amounts using an exchange rate. Useful for multi-currency financial reporting.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `amountRange` | ✅ | Range of amounts to convert |
| `exchangeRate` | ✅ | Exchange rate to apply |
| `outputRange` | ✅ | Output range for converted amounts |

---

← [Back to index](README.md)