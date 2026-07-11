# Investment Analysis

Tools for investment evaluation: NPV, IRR, and financial ratios.

## `excel_calculate_npv`

Calculate Net Present Value (NPV) for investment evaluation. Discounts future cash flows to present value — essential for investment decision-making.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `rate` | ✅ | Discount rate (decimal, e.g., `0.1` for 10%) |
| `valuesRange` | ✅ | Cash flow values range |

## `excel_calculate_irr`

Calculate Internal Rate of Return (IRR) — the discount rate that makes NPV equal to zero. Represents the expected annual return of an investment.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `valuesRange` | ✅ | Cash flow values range |
| `guess` | ❌ | Initial guess for the rate |

## `excel_calculate_financial_ratio`

Calculate key financial ratios.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `ratioType` | ✅ | Ratio type (see below) |
| `numeratorRange` | ✅ | Numerator range |
| `denominatorRange` | ✅ | Denominator range |

### Supported `ratioType` values

| Ratio | Description |
|-------|-------------|
| `current` | Current Ratio — ability to pay short-term obligations |
| `quick` | Quick Ratio — liquidity without inventory |
| `debt-to-equity` | Debt-to-Equity — financial leverage |
| `return-on-equity` | Return on Equity (ROE) — profitability relative to equity |
| `profit-margin` | Profit Margin — net profit as percentage of revenue |

---

← [Back to index](README.md)