# Font & Text Styling

Tools for font styling, font size, and rich text formatting.

## `excel_set_font_style`

Apply font styling (bold, italic, underline, strikethrough) to a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `bold` | ❌ | Bold text |
| `italic` | ❌ | Italic text |
| `underline` | ❌ | Underline style: `true`, `"double"`, `"singleAccounting"`, `"doubleAccounting"` |
| `strikethrough` | ❌ | Strikethrough text |

## `excel_set_font_name_size`

Set font family and size for a cell or range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Start cell (e.g., `A1`) |
| `endCell` | ❌ | End cell for range |
| `fontName` | ❌ | Font name (e.g., `Arial`, `Calibri`, `Times New Roman`) |
| `fontSize` | ❌ | Font size in points |

## `excel_set_rich_text`

Apply multiple text styles within a single cell (rich text).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Cell address (e.g., `A1`) |
| `richText` | ✅ | Array of text segments with individual formatting |

### Structure of each `richText` element

| Field | Required | Description |
|-------|----------|-------------|
| `text` | ✅ | Text content |
| `bold` | ❌ | Bold |
| `italic` | ❌ | Italic |
| `underline` | ❌ | Underline |
| `fontColor` | ❌ | Font color (hex) |
| `fontSize` | ❌ | Font size |
| `fontName` | ❌ | Font name |

### Example

```javascript
await excel_set_rich_text({
  filename: "document",
  worksheet: "Sheet1",
  cellAddress: "A1",
  richText: [
    { text: "Important: ", bold: true, fontColor: "FFFF0000" },
    { text: "This is critical data", fontColor: "FF000000" }
  ]
});
```

---

← [Back to index](README.md)