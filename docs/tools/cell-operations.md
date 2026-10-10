# Cell Operations

Tools for reading, writing, copying, searching, sorting, and managing Excel cells and named ranges.

## `excel_read_cell`

Read the value, type, and formula of an individual cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Cell coordinate (e.g., `A1`, `B12`) |

**Response data:** `{ address: string, value: unknown, type: string, formula?: string }`

---

## `excel_read_range`

Read a 2D rectangular grid of cell values.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `range` | ✅ | Range specification object with `start: { row, column }` and `end: { row, column }` |

**Response data:** 2D array of `CellValue[][]`

---

## `excel_write_cell`

Write a value or formula into an individual cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Cell coordinate (e.g., `A1`) |
| `value` | ✅ | Value to write: string, number, boolean, or formula (`=SUM(A1:A5)`) |

---

## `excel_write_batch`

Write multiple cells in a single batch operation for optimal performance.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `data` | ✅ | Array of `{ cellAddress: string, value: unknown }` objects |

**Response data:** `{ count: number }`

---

## `excel_copy_range`

Copy a rectangular range of cells from a source location to a destination start cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `sourceStart` | ✅ | Source range top-left cell (e.g., `A1`) |
| `sourceEnd` | ✅ | Source range bottom-right cell (e.g., `C10`) |
| `targetStart` | ✅ | Destination top-left cell (e.g., `E1`) |
| `targetWorksheet` | ❌ | Target worksheet name (defaults to source worksheet) |

---

## `excel_find_replace`

Search and replace text within all cells of a worksheet.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `findText` | ✅ | Text to search for |
| `replaceText` | ✅ | Replacement text |
| `matchCase` | ❌ | Case-sensitive search (`true`/`false`, default: `false`) |
| `matchEntireCell` | ❌ | Match whole cell content only (`true`/`false`, default: `false`) |

**Response data:** `{ count: number }`

---

## `excel_sort_range`

Sort rows within a rectangular range by a specified column.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Top-left cell of the range to sort (e.g., `A2`) |
| `endCell` | ✅ | Bottom-right cell of the range to sort (e.g., `D20`) |
| `sortColumn` | ✅ | 1-based column index to sort by (1 for first column in range) |
| `ascending` | ✅ | `true` for ascending (A-Z, 0-9), `false` for descending |

---

## `excel_get_named_ranges` & `excel_add_named_range`

Manage workbook-level scoped named ranges for cleaner formulas.

### `excel_add_named_range`

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Worksheet where range resides |
| `name` | ✅ | Identifier name (e.g., `TaxRates`, `RevenueQ1`) |
| `startCell` | ✅ | Range start cell (e.g., `B2`) |
| `endCell` | ✅ | Range end cell (e.g., `B13`) |

### `excel_get_named_ranges`

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |

**Response data:** Array of `{ name: string, ref: string }`

---

← [Back to index](README.md)
