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

## `excel_calculate_progressive_tax`

Calculate progressive (graduated) taxes, such as Personal Income Tax (PIT) or graduated Corporate Income Tax brackets. Applies each rate tier only to the portion of income exceeding the tier threshold.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `amountRange` | ✅ | Range containing taxable income amounts (e.g., `A2:A10`) |
| `brackets` | ✅ | Array of tax brackets `[{threshold: number, rate: number}]`, sorted ascending by threshold |
| `outputRange` | ✅ | Range where computed tax amounts will be written (e.g., `B2`) |

### Example Tax Brackets (Vietnam PIT Biểu thuế luỹ tiến từng phần)

```json
[
  { "threshold": 0, "rate": 5 },
  { "threshold": 5000000, "rate": 10 },
  { "threshold": 10000000, "rate": 15 },
  { "threshold": 18000000, "rate": 20 },
  { "threshold": 32000000, "rate": 25 },
  { "threshold": 52000000, "rate": 30 },
  { "threshold": 80000000, "rate": 35 }
]
```

---

← [Back to index](README.md)