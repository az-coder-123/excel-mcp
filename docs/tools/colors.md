# Colors & Backgrounds

Tools for setting background and font colors.

## `excel_set_background_color`

Set the background fill color for a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `color` | ✅ | Background color in hex (e.g., `FFFF0000` for red) |

## `excel_set_font_color`

Set the font color for a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `color` | ✅ | Font color in hex (e.g., `FF000000` for black) |

---

← [Back to index](README.md)