# Borders & Styling

Tools for applying borders to cells and ranges.

## `excel_set_border`

Apply custom borders to specific sides of a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `borderStyle` | ❌ | Border style: `thin`, `medium`, `thick`, `double`, `dotted`, `dashed`, etc. |
| `borderColor` | ❌ | Border color in hex (e.g., `FF000000`) |
| `top` | ❌ | Apply top border |
| `bottom` | ❌ | Apply bottom border |
| `left` | ❌ | Apply left border |
| `right` | ❌ | Apply right border |
| `diagonal` | ❌ | Apply diagonal borders |

## `excel_apply_all_borders`

Apply borders to **all sides** of a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `borderStyle` | ❌ | Border style (default: `thin`) |
| `borderColor` | ❌ | Border color in hex (default: `FF000000`) |

## `excel_apply_outline_border`

Apply borders **only to the outer edges** (outline) of a range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `borderStyle` | ❌ | Border style (default: `thin`) |
| `borderColor` | ❌ | Border color in hex (default: `FF000000`) |

---

← [Back to index](README.md)