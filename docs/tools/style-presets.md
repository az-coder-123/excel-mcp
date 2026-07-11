# Style Presets

Pre-built style tools for headers, titles, and tables — apply professional formatting in a single call.

## `excel_apply_header_style`

Apply professional header style (bold, centered, gray background, bottom border).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `backgroundColor` | ❌ | Background color in hex (default: `FFD3D3D3`) |
| `fontColor` | ❌ | Font color in hex (default: `FF000000`) |

## `excel_apply_title_style`

Apply professional title style (large, bold, centered, dark blue text).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `fontSize` | ❌ | Font size in points (default: `16`) |
| `color` | ❌ | Font color in hex (default: `FF1F4E78`) |

## `excel_apply_table_style`

Apply professional table formatting with headers and alternating row colors.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell of table (e.g., `A1`) |
| `endCell` | ✅ | End cell of table (e.g., `D10`) |
| `headerBackgroundColor` | ❌ | Header background color in hex (default: `FFD3D3D3`) |
| `headerFontColor` | ❌ | Header font color in hex (default: `FF000000`) |
| `alternateRowColor` | ❌ | Alternate row background color in hex (default: `FFF9F9F9`) |
| `hasHeader` | ❌ | Whether first row is a header (default: `true`) |

---

← [Back to index](README.md)