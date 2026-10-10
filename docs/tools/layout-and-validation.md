# Layout, Dimensions & Data Validation

Tools for workbook layout configuration, dimensions, printing, cell notes, hyperlinks, and data validation rules.

---

## 1. Sheet Dimensions & Layout

### `excel_auto_fit_columns`
Automatically sizes column widths based on maximum cell content length with dynamic padding.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startColumn` | ❌ | Range start column letter (e.g., `A`) |
| `endColumn` | ❌ | Range end column letter (e.g., `F`) |

### `excel_set_column_width`
Sets explicit width for a column.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `column` | ✅ | Column letter (e.g., `B`, `c`) |
| `width` | ✅ | Character width number (e.g., `25`) |

### `excel_set_row_height`
Sets explicit point height for a row.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `row` | ✅ | 1-based row index |
| `height` | ✅ | Height in points (e.g., `35`) |

### `excel_freeze_panes`
Freezes rows and columns above and to the left of the specified cell so headers stay fixed during scrolling.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `cellAddress` | ✅ | Freeze split cell (e.g., `B2` freezes row 1 and column A) |

### `excel_set_print_area`
Defines the printable page range for worksheet printing or PDF export.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Print area start cell (e.g., `A1`) |
| `endCell` | ✅ | Print area end cell (e.g., `G50`) |

### `excel_add_header_footer`
Configures page headers and footers for printed reports.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `header` | ❌ | Header text |
| `footer` | ❌ | Footer text |

---

## 2. Interactive Cell Features

### `excel_add_comment`
Attaches an auditor or reviewer note to a cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `cellAddress` | ✅ | Cell address (e.g., `C15`) |
| `comment` | ✅ | Note text |
| `author` | ❌ | Note author name |

### `excel_remove_comment`
Removes a comment from a cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `cellAddress` | ✅ | Cell address |

### `excel_add_hyperlink`
Inserts a clickable URL or sheet bookmark hyperlink into a cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `cellAddress` | ✅ | Target cell address |
| `url` | ✅ | Target URL (e.g., `https://example.com`) |
| `display` | ❌ | Friendly text label |

### `excel_add_data_validation`
Applies data validation rules (dropdown list, integer range, decimal range, text length) to prevent invalid input.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `cellAddress` | ✅ | Cell address |
| `type` | ✅ | `list`, `whole`, `decimal`, `date`, `textLength` |
| `formula1` | ✅ | Allowed list values (e.g., `"Active,Inactive"`) or minimum |
| `formula2` | ❌ | Maximum value (for `between` operator) |
| `operator` | ❌ | `between`, `equal`, `greaterThan`, etc. |
| `errorMessage` | ❌ | Prompt message when invalid value entered |

---

## 3. All-in-One Extended Formatting

### `excel_format_worksheet`
Formats an entire worksheet in a single atomic MCP call (title styling, header row styling, alternating data rows, borders, and column auto-fit).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `titleStyle` | ❌ | Object `{ cellAddress, fontSize, color, bold }` |
| `headerStyle` | ❌ | Object `{ startCell, endCell, backgroundColor, fontColor, bold }` |
| `dataStyle` | ❌ | Object `{ startCell, endCell, borders, borderStyle, alternateRows }` |
| `autoFitColumns` | ❌ | Boolean auto-fit all columns (default: `true`) |

### `excel_batch_format`
Applies a JSON-serialized array of formatting operations in one call.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `operations` | ✅ | JSON array of formatting operation objects |

---

← [Back to index](README.md)
