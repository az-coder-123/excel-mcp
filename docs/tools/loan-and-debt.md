# Loan & Debt Management

Tools for generating loan amortization schedules and accounts receivable/payable aging reports.

## `excel_create_amortization_schedule`

Generate a loan amortization table with payment breakdown (principal vs interest) and remaining balance after each payment. Useful for mortgages, car loans, and business loans.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell for the schedule table |
| `principal` | ✅ | Principal amount |
| `annualRate` | ✅ | Annual interest rate (decimal, e.g., `0.05` for 5%) |
| `numberOfPeriods` | ✅ | Number of periods (months) |

**Output columns:** Period | Payment | Principal | Interest | Balance

## `excel_create_aging_report`

Create an accounts receivable/payable aging report. Categorizes invoices by days past due: 0–30, 31–60, 61–90, 91–120, 120+ days.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `invoiceDateColumn` | ✅ | Column containing invoice dates |
| `amountColumn` | ✅ | Column containing invoice amounts |
| `asOfDate` | ✅ | As-of date for aging calculation (ISO format) |
| `outputStartCell` | ✅ | Start cell for the aging summary table |

**Output:** Aging summary table with totals for each category and total outstanding.

---

← [Back to index](README.md)