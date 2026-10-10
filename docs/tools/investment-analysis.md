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

## `excel_calculate_xirr`

Calculate Internal Rate of Return for irregular or non-periodic cash flows (XIRR). Unlike standard periodic IRR which assumes equidistant intervals, XIRR factors in exact transaction dates using day-fraction discounting ($d_i - d_0)/365$.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `dateRange` | ✅ | Range containing transaction dates (e.g., `A2:A6`) |
| `valuesRange` | ✅ | Range containing cash flow values (e.g., `B2:B6`) |
| `guess` | ❌ | Initial discount rate guess (default: `0.1` for 10%) |

**Requirements:**
- Cash flows must contain at least one positive and at least one negative amount.
- `dateRange` and `valuesRange` must have identical row dimensions.
- Returns annualized rate in percentage (e.g., `14.25` for 14.25%).

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