# Data Utilities & Advanced Analytics

Tools for data transformation, CSV transfer, duplicate management, lookups, grouping, profiling, and search.

---

## 1. CSV Transfer

### `excel_import_csv`
Imports CSV content into a worksheet at the specified start cell, parsing numbers and text.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Target worksheet name |
| `csvContent` | ✅ | CSV raw string |
| `startCell` | ✅ | Destination starting cell (e.g., `A1`) |
| `delimiter` | ❌ | Column delimiter (default: `,`) |

### `excel_export_csv`
Exports a cell range to a CSV-formatted string.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Source worksheet name |
| `startCell` | ✅ | Top-left range cell |
| `endCell` | ✅ | Bottom-right range cell |
| `delimiter` | ❌ | Column delimiter (default: `,`) |

---

## 2. Duplicate Management

### `excel_find_duplicates`
Finds duplicate rows or values in a given range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |

### `excel_count_unique_values`
Counts unique and duplicate occurrences within a range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |

### `excel_highlight_duplicates`
Highlights duplicate cells in a range with a designated background color.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `color` | ❌ | ARGB color hex (default: `FFFFC7CE`) |

### `excel_remove_duplicates`
Removes duplicate rows across designated columns, keeping the first occurrence.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `columns` | ❌ | Array of 0-based column indices to evaluate |

---

## 3. Transformations & Flash Fill

### `excel_text_to_columns`
Splits delimited text from a source cell into adjacent columns.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `sourceCell` | ✅ | Source cell containing delimited string |
| `targetCell` | ✅ | Destination starting cell |
| `delimiter` | ✅ | Splitting character (e.g., `,`, `-`, ` `) |
| `numberOfColumns` | ✅ | Maximum columns to split into |

### `excel_flash_fill`
Infers transformation pattern from sample rows and applies it across target range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `sourceRange` | ✅ | Example/source range (e.g., `A1:B3`) |
| `targetRange` | ✅ | Target fill range (e.g., `C1:C10`) |

---

## 4. Lookups & Pivot Tables

### `excel_vlookup`
Writes a standard Excel `VLOOKUP` formula into the target cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `targetCell` | ✅ | Destination cell address |
| `lookupValue` | ✅ | Value or cell reference to search |
| `tableArray` | ✅ | Table range (e.g., `A1:D100`) |
| `colIndex` | ✅ | 1-based column return index |
| `rangeLookup` | ❌ | Approximate match boolean (default: `false`) |

### `excel_index_match`
Writes an idiomatic `INDEX(MATCH())` formula into the target cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `targetCell` | ✅ | Destination cell address |
| `returnRange` | ✅ | Return array range (e.g., `C1:C100`) |
| `lookupRange` | ✅ | Lookup array range (e.g., `A1:A100`) |
| `lookupValue` | ✅ | Value or cell reference to match |

### `excel_create_pivot_table`
Generates a summary cross-tabulation table from source data.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `sourceWorksheet` | ✅ | Source worksheet name |
| `sourceStartCell` | ✅ | Source data start cell |
| `sourceEndCell` | ✅ | Source data end cell |
| `targetWorksheet` | ✅ | Output worksheet name |
| `targetCell` | ✅ | Placement top-left cell |
| `rowFields` | ❌ | Array of row grouping field names |
| `valueFields` | ❌ | Array of numeric aggregation field names |

---

## 5. Analytics & Profiling

### `excel_get_column_stats`
Calculates comprehensive statistical metrics on a column, including distinct counts, frequencies, and missing values.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `column` | ✅ | Column letter (e.g., `A`, `b`, `AA`) |
| `hasHeader` | ❌ | Whether row 1 is a header (default: `true`) |

### `excel_filter_data`
Filters rows matching condition rules (`equals`, `contains`, `greaterThan`, `lessThan`, `between`).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `filters` | ✅ | Array of `{ column, operator, value, matchCase }` |
| `hasHeader` | ❌ | Include header row (default: `true`) |

### `excel_group_aggregate`
Groups data by a key column and aggregates values with `count`, `sum`, `avg`, `min`, or `max`.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `groupByColumn` | ✅ | Grouping column letter |
| `aggregateColumn` | ❌ | Column to aggregate (optional for `count`) |
| `operation` | ✅ | `count`, `sum`, `avg`, `min`, or `max` |
| `hasHeader` | ❌ | Include header row (default: `true`) |

### `excel_profile_data`
Profiles all columns in a range, analyzing data types, null counts, and pattern detection (emails, phones, dates).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell |
| `endCell` | ✅ | Range end cell |
| `hasHeader` | ❌ | Header flag (default: `true`) |

### `excel_search`
Searches cells matching query across sheet or specific column ranges.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `searchQuery` | ✅ | Search text |
| `searchType` | ❌ | `exact`, `contains`, `startsWith`, `endsWith`, `regex` |
| `columnRange` | ❌ | Restrict to column or range (e.g., `A:A`, `B:D`) |
| `matchCase` | ❌ | Case sensitivity (default: `false`) |
| `maxResults` | ❌ | Result cap (default: `100`) |

### `excel_compare_ranges`
Compares two cell ranges and detects value, formula, or format discrepancies.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet1` | ✅ | First worksheet name |
| `range1Start` | ✅ | Range 1 start cell |
| `range1End` | ✅ | Range 1 end cell |
| `worksheet2` | ❌ | Second worksheet name (default: same sheet) |
| `range2Start` | ✅ | Range 2 start cell |
| `range2End` | ✅ | Range 2 end cell |
| `compareType` | ❌ | `values`, `formulas`, `formats` |

---

← [Back to index](README.md)
