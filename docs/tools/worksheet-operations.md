# Worksheet & Structure Operations

Tools for managing worksheet lifecycles, inserting/deleting rows and columns, merging cells, tables, auto-filtering, freeze panes, print setup, and headers/footers.

## `excel_list_worksheets`

List all worksheets in an opened workbook, including visibility, index, row count, and column count.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |

**Response data:** Array of `WorksheetInfo` `{ name, index, rowCount, columnCount, hidden }`

---

## `excel_add_worksheet`

Create a new worksheet in the workbook.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheetName` | ✅ | Name for the new worksheet |

---

## `excel_delete_worksheet`

Delete an existing worksheet. (Requires `delete` permission).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheetName` | ✅ | Name of the worksheet to remove |

---

## `excel_rename_worksheet`

Rename an existing worksheet.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `oldName` | ✅ | Current worksheet name |
| `newName` | ✅ | New worksheet name |

---

## `excel_copy_worksheet`

Duplicate an existing worksheet within the workbook.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `sourceWorksheet` | ✅ | Name of the worksheet to duplicate |
| `newWorksheetName` | ✅ | Name for the copy |

---

## Row & Column Mutations

Insert or remove rows and columns dynamically:

### `excel_insert_rows`
- `filename`, `worksheet`: Target sheet
- `startRow`: 1-based row index to insert at
- `count`: Number of blank rows to insert

### `excel_insert_columns`
- `filename`, `worksheet`: Target sheet
- `startColumn`: 1-based column index to insert at
- `count`: Number of blank columns to insert

### `excel_delete_rows`
- `filename`, `worksheet`: Target sheet
- `startRow`: 1-based start row index
- `count`: Number of rows to delete

### `excel_delete_columns`
- `filename`, `worksheet`: Target sheet
- `startColumn`: 1-based start column index
- `count`: Number of columns to delete

---

## Cell Merging

### `excel_merge_cells`
Merge a range of cells into a single unified block:
- `filename`, `worksheet`: Target sheet
- `startCell`: Top-left cell (e.g., `A1`)
- `endCell`: Bottom-right cell (e.g., `D1`)

### `excel_unmerge_cells`
Split a merged range back into individual cells:
- `filename`, `worksheet`: Target sheet
- `startCell`, `endCell`: Merged range coordinates

---

## Tables & Filters

### `excel_add_table`
Convert a rectangular range into an official Excel Table object with banded rows and sort/filter dropdowns:
- `filename`, `worksheet`: Target sheet
- `tableName`: Unique table name (e.g., `SalesTable`)
- `range`: Range coordinates (e.g., `A1:F50`)
- `hasHeaders`: Whether the first row contains column headers (default: `true`)

### `excel_add_filter` & `excel_remove_filter`
Add or remove auto-filter dropdown arrows on header rows without creating a full table object.

---

## View & Page Setup

### `excel_freeze_panes`
Freeze header rows and/or columns when scrolling:
- `filename`, `worksheet`: Target sheet
- `cellAddress`: Top-left unfrozen cell (e.g., `A2` freezes row 1, `B2` freezes row 1 and column A)

### `excel_set_print_area`
Specify the printable cell boundary for exporting or printing:
- `filename`, `worksheet`: Target sheet
- `startCell`, `endCell`: Range bounds

### `excel_add_header_footer`
Configure headers and footers for printed sheets:
- `filename`, `worksheet`: Target sheet
- `header`: Header text / placeholders
- `footer`: Footer text (e.g., `Page &P of &N`)

---

← [Back to index](README.md)
