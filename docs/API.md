# API Reference

## Overview

The Excel MCP Server provides 10 tools for Excel operations. All tools follow the MCP protocol specification and return structured responses.

## Response Format

### Success Response

```json
{
  "content": [
    {
      "type": "text",
      "text": "Operation result data"
    }
  ]
}
```

### Error Response

```json
{
  "content": [
    {
      "type": "text",
      "text": "Error: Description of what went wrong"
    }
  ],
  "isError": true
}
```

## Tools

### excel_open_workbook

Opens an Excel workbook file for subsequent operations.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filePath` | string | Yes | Absolute path to the Excel file |

**Returns**: Workbook information including filename and worksheet list.

**Example Request**:

```json
{
  "filePath": "/workspace/reports/sales-2024.xlsx"
}
```

**Example Response**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheetCount": 3,
  "worksheets": [
    {
      "name": "Summary",
      "index": 0,
      "rowCount": 50,
      "columnCount": 10,
      "hidden": false
    },
    {
      "name": "Details",
      "index": 1,
      "rowCount": 1000,
      "columnCount": 15,
      "hidden": false
    }
  ]
}
```

**Errors**:
- File not found
- Permission denied
- Invalid file format
- File exceeds size limit

---

### excel_create_workbook

Creates a new Excel workbook.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name for the workbook (without extension) |
| `outputPath` | string | Yes | Directory path where the file will be saved |

**Returns**: Workbook information for the new file.

**Example Request**:

```json
{
  "filename": "quarterly-report",
  "outputPath": "/workspace/reports"
}
```

**Example Response**:

```json
{
  "filename": "/workspace/reports/quarterly-report.xlsx",
  "worksheetCount": 1,
  "worksheets": [
    {
      "name": "Sheet1",
      "index": 0,
      "rowCount": 0,
      "columnCount": 0,
      "hidden": false
    }
  ]
}
```

**Errors**:
- Permission denied
- Directory not found
- Invalid filename

---

### excel_read_cell

Reads the value of a single cell.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address (e.g., A1, B2) |

**Returns**: Cell value and type information.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheet": "Summary",
  "cellAddress": "A1"
}
```

**Example Response**:

```json
{
  "address": "A1",
  "value": "Total Revenue",
  "type": "string"
}
```

**Cell Types**:
- `string` - Text value
- `number` - Numeric value
- `boolean` - True/false value
- `date` - Date/time value
- `formula` - Formula with computed result
- `error` - Error value

**Errors**:
- Workbook not opened
- Worksheet not found
- Invalid cell address

---

### excel_read_range

Reads values from a range of cells.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `startCell` | string | Yes | Start cell of range (e.g., A1) |
| `endCell` | string | Yes | End cell of range (e.g., C10) |

**Returns**: 2D array of cell values.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheet": "Summary",
  "startCell": "A1",
  "endCell": "C5"
}
```

**Example Response**:

```json
[
  [
    {"address": "A1", "value": "Product", "type": "string"},
    {"address": "B1", "value": "Revenue", "type": "string"},
    {"address": "C1", "value": "Units", "type": "string"}
  ],
  [
    {"address": "A2", "value": "Widget A", "type": "string"},
    {"address": "B2", "value": 50000, "type": "number"},
    {"address": "C2", "value": 1000, "type": "number"}
  ]
]
```

**Errors**:
- Workbook not opened
- Worksheet not found
- Invalid cell range

---

### excel_write_cell

Writes a value to a cell.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address (e.g., A1, B2) |
| `value` | any | Yes | Value to write |

**Returns**: Success confirmation.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheet": "Summary",
  "cellAddress": "A1",
  "value": "New Title"
}
```

**Value Types**:

- **String**: `"Hello World"`
- **Number**: `42` or `3.14`
- **Boolean**: `true` or `false`
- **Formula**: `"=SUM(A1:A10)"` (must start with `=`)

**Example Formula Request**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheet": "Summary",
  "cellAddress": "B10",
  "value": "=SUM(B1:B9)"
}
```

**Errors**:
- Workbook not opened
- Worksheet not found
- Permission denied
- Invalid cell address

---

### excel_list_worksheets

Lists all worksheets in an opened workbook.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |

**Returns**: Array of worksheet information.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx"
}
```

**Example Response**:

```json
[
  {
    "name": "Summary",
    "index": 0,
    "rowCount": 50,
    "columnCount": 10,
    "hidden": false
  },
  {
    "name": "Details",
    "index": 1,
    "rowCount": 1000,
    "columnCount": 15,
    "hidden": false
  },
  {
    "name": "Config",
    "index": 2,
    "rowCount": 5,
    "columnCount": 2,
    "hidden": true
  }
]
```

**Errors**:
- Workbook not opened

---

### excel_add_worksheet

Adds a new worksheet to an opened workbook.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheetName` | string | Yes | Name for the new worksheet |

**Returns**: Information about the new worksheet.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheetName": "Q4 Summary"
}
```

**Example Response**:

```json
{
  "name": "Q4 Summary",
  "index": 3,
  "rowCount": 0,
  "columnCount": 0,
  "hidden": false
}
```

**Errors**:
- Workbook not opened
- Permission denied
- Worksheet name already exists

---

### excel_save_workbook

Saves the workbook to disk.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `outputPath` | string | No | Alternative path to save the file |

**Returns**: Success confirmation.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx"
}
```

**Example Request (Save As)**:

```json
{
  "filename": "sales-2024.xlsx",
  "outputPath": "/workspace/archive/sales-2024-backup.xlsx"
}
```

**Errors**:
- Workbook not opened
- Permission denied
- Path not allowed
- Disk full

---

### excel_close_workbook

Closes an opened workbook and frees resources.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |

**Returns**: Success confirmation.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx"
}
```

**Note**: Always close workbooks when done to free memory.

**Errors**:
- Workbook not opened

---

### excel_get_cell_info

Gets detailed information about a cell, including formulas.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address (e.g., A1, B2) |

**Returns**: Detailed cell information.

**Example Request**:

```json
{
  "filename": "sales-2024.xlsx",
  "worksheet": "Summary",
  "cellAddress": "B10"
}
```

**Example Response (Formula Cell)**:

```json
{
  "address": "B10",
  "value": 150000,
  "type": "formula",
  "formula": "=SUM(B1:B9)"
}
```

**Example Response (Regular Cell)**:

```json
{
  "address": "A1",
  "value": "Total Revenue",
  "type": "string"
}
```

**Errors**:
- Workbook not opened
- Worksheet not found
- Invalid cell address

## Common Patterns

### Reading and Writing

```javascript
// 1. Open workbook
await excel_open_workbook({ filePath: "/data/report.xlsx" });

// 2. Read data
const cell = await excel_read_cell({
  filename: "report.xlsx",
  worksheet: "Sheet1",
  cellAddress: "A1"
});

// 3. Modify data
await excel_write_cell({
  filename: "report.xlsx",
  worksheet: "Sheet1",
  cellAddress: "A1",
  value: "Updated Title"
});

// 4. Save and close
await excel_save_workbook({ filename: "report.xlsx" });
await excel_close_workbook({ filename: "report.xlsx" });
```

### Creating New Files

```javascript
// 1. Create workbook
await excel_create_workbook({
  filename: "new-report",
  outputPath: "/workspace"
});

// 2. Write data
await excel_write_cell({
  filename: "new-report.xlsx",
  worksheet: "Sheet1",
  cellAddress: "A1",
  value: "Report Title"
});

// 3. Save
await excel_save_workbook({ filename: "new-report.xlsx" });
await excel_close_workbook({ filename: "new-report.xlsx" });
```

### Working with Formulas

```javascript
// Write a formula
await excel_write_cell({
  filename: "report.xlsx",
  worksheet: "Sheet1",
  cellAddress": "B10",
  value: "=SUM(B1:B9)"
});

// Read the formula and its result
const cell = await excel_get_cell_info({
  filename: "report.xlsx",
  worksheet: "Sheet1",
  cellAddress: "B10"
});
// Result: { value: 150000, type: "formula", formula: "=SUM(B1:B9)" }
```

## Rate Limits and Quotas

- Maximum file size: 50MB (configurable)
- Maximum concurrent workbooks: 10 (configurable)
- Operation timeout: 30 seconds (configurable)

## Versioning

Current API version: 1.0.0

All tools maintain backward compatibility. Breaking changes will result in a major version bump.