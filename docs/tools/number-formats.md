# Number Formats

Tools for formatting numbers, currency, percentages, and dates.

## `excel_set_number_format`

Apply a custom number format to a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `format` | ✅ | Format code or preset: `currency`, `percentage`, `number`, `date`, `time`, `datetime`, `fraction`, `scientific`, `text`, or a custom format |

## `excel_apply_currency_format`

Apply currency formatting with configurable symbol and decimal places.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `symbol` | ❌ | Currency symbol (default: `$`) |
| `decimalPlaces` | ❌ | Number of decimal places (default: `2`) |

## `excel_apply_percentage_format`

Apply percentage formatting.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `decimalPlaces` | ❌ | Number of decimal places (default: `2`) |

## `excel_apply_date_format`

Apply date formatting.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `format` | ❌ | Date format preset: `short`, `long`, `iso`, or a custom format like `dd/mm/yyyy` |

---

← [Back to index](README.md)