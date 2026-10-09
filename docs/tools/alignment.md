# Alignment & Text Control

Tools for cell alignment, wrap text, and text display control.

## `excel_set_alignment`

Set horizontal/vertical alignment and text display options for a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `horizontal` | ❌ | Horizontal alignment: `left`, `center`, `right`, `fill`, `justify`, `distributed` |
| `vertical` | ❌ | Vertical alignment: `top`, `middle`, `bottom`, `justify`, `distributed` |
| `wrapText` | ❌ | Wrap text within cells |
| `shrinkToFit` | ❌ | Shrink text to fit cell |
| `indent` | ❌ | Indent level (0–255) |
| `textRotation` | ❌ | Text rotation angle in degrees (0–180) |

## `excel_center_text`

Quickly center text both horizontally and vertically in a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |

---

← [Back to index](README.md)