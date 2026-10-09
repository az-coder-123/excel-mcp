# Financial Calculations

Tools for basic financial calculations: sum, average, running totals, percentages, and year-to-date values.

## `excel_financial_sum`

Calculate sum with optional criteria (SUMIF-like).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `rangeStart` | ✅ | Range start cell |
| `rangeEnd` | ✅ | Range end cell |
| `criteriaColumn` | ❌ | Column to filter by |
| `criteriaValue` | ❌ | Filter value |

## `excel_financial_average`

Calculate the average of values in a range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `rangeStart` | ✅ | Range start cell |
| `rangeEnd` | ✅ | Range end cell |

## `excel_running_total`

Calculate cumulative running totals.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `valueStartCell` | ✅ | Values range start |
| `valueEndCell` | ✅ | Values range end |
| `outputStartCell` | ✅ | Output range start |

## `excel_percentage_of_total`

Calculate each value as a percentage of the total.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `valueStartCell` | ✅ | Values range start |
| `valueEndCell` | ✅ | Values range end |
| `outputStartCell` | ✅ | Output range start |

## `excel_year_to_date`

Calculate year-to-date cumulative values.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `dateColumn` | ✅ | Date column |
| `valueColumn` | ✅ | Value column |
| `rangeStart` | ✅ | Data range start |
| `rangeEnd` | ✅ | Data range end |
| `outputColumn` | ✅ | Output column |

---

← [Back to index](README.md)