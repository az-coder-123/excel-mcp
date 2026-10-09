# API Reference

## Overview

The Excel MCP Server provides comprehensive tools for Excel operations. All tools follow the MCP protocol specification and return structured responses.

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

### excel_filter_data

Filter data based on multiple conditions. This is a powerful tool for extracting specific records from large datasets.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `startCell` | string | Yes | Start cell of data range (e.g., A1) |
| `endCell` | string | Yes | End cell of data range (e.g., D100) |
| `filters` | array | Yes | Array of filter conditions |
| `hasHeader` | boolean | No | Whether the first row contains header (default: true) |

**Filter Condition Object**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `column` | string | Yes | Column letter (e.g., A, B, C) |
| `operator` | string | Yes | Comparison operator (see below) |
| `value` | string | No | Value to compare with (not required for isBlank/isNotBlank) |
| `matchCase` | boolean | No | Case sensitive comparison (default: false) |

**Supported Operators**:

- `equals` - Exact match
- `notEquals` - Not equal to value
- `contains` - Contains the text
- `notContains` - Does not contain the text
- `greaterThan` - Greater than (numeric comparison)
- `lessThan` - Less than (numeric comparison)
- `greaterThanOrEqual` - Greater than or equal
- `lessThanOrEqual` - Less than or equal
- `isBlank` - Cell is empty
- `isNotBlank` - Cell is not empty

**Returns**: Filtered data array with matching rows.

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

---

### excel_list_formulas

Lists all formulas in a worksheet with their locations, values, and types.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |

**Returns**: Array of all formulas with their metadata.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary"
}
```

**Example Response**:

```json
{
  "total": 5,
  "formulas": [
    {
      "cell": "B10",
      "formula": "=SUM(B1:B9)",
      "value": 150000,
      "type": "number"
    },
    {
      "cell": "C10",
      "formula": "=AVERAGE(C1:C9)",
      "value": 16666.67,
      "type": "number"
    }
  ]
}
```

**Use Cases**:
- Audit all formulas in a worksheet
- Identify formula locations before editing
- Document formula structure
- Verify formula count and distribution

**Errors**:
- Workbook not opened
- Worksheet not found

---

### excel_analyze_formula

Analyzes a specific formula to extract its components and structure.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address containing formula (e.g., B10) |

**Returns**: Detailed analysis of formula components.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary",
  "cellAddress": "B10"
}
```

**Example Response**:

```json
{
  "cell": "B10",
  "formula": "=SUM(B1:B9)",
  "functions": [
    {
      "name": "SUM",
      "arguments": ["B1:B9"]
    }
  ],
  "cellReferences": ["B1", "B9"],
  "namedRanges": [],
  "complexity": 1,
  "error": null
}
```

**Analysis Components**:
- `functions`: Array of function names and their arguments
- `cellReferences`: All cells referenced in the formula
- `namedRanges`: Named ranges used (if any)
- `complexity`: Formula complexity score (higher = more complex)
- `error`: Any error in the formula

**Errors**:
- Workbook not opened
- Worksheet not found
- Cell does not contain formula

---

### excel_get_dependencies

Gets complete dependency information for a formula cell including precedents and dependents.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address to analyze |

**Returns**: Dependency graph information.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary",
  "cellAddress": "B10"
}
```

**Example Response**:

```json
{
  "cell": "B10",
  "formula": "=SUM(B1:B9)",
  "precedents": [
    {"cell": "B1", "value": 10000},
    {"cell": "B2", "value": 20000}
  ],
  "dependents": [
    {"cell": "C10", "formula": "=B10*0.1"}
  ],
  "isCircular": false,
  "circularChain": null
}
```

**Dependency Types**:
- `precedents`: Cells that provide input to this formula (data sources)
- `dependents`: Cells that use this formula's result (depend on this cell)
- `isCircular`: Whether there's a circular reference
- `circularChain`: Chain of cells if circular reference exists

**Errors**:
- Workbook not opened
- Worksheet not found
- Cell not found

---

### excel_trace_precedents

Traces data sources (precedents) for a cell - shows which cells contribute to the formula.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address to trace |

**Returns**: Precedent cells and their values.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary",
  "cellAddress": "D15"
}
```

**Example Response**:

```json
{
  "cell": "D15",
  "formula": "=SUM(D1:D14)",
  "precedents": [
    {"cell": "D1", "value": 5000},
    {"cell": "D2", "value": 7500},
    {"cell": "D3", "value": 10000}
  ],
  "count": 14
}
```

**Use Cases**:
- Debug formula results
- Identify data source cells
- Understand calculation flow
- Find incorrect input values

**Errors**:
- Workbook not opened
- Worksheet not found
- Cell not found

---

### excel_trace_dependents

Traces which cells use this cell's value - shows impact of changing this cell.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address to trace |

**Returns**: Dependent cells that reference this cell.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary",
  "cellAddress": "B5"
}
```

**Example Response**:

```json
{
  "cell": "B5",
  "dependents": [
    {
      "cell": "B10",
      "formula": "=SUM(B1:B9)"
    },
    {
      "cell": "B15",
      "formula": "=AVERAGE(B1:B14)"
    }
  ],
  "count": 2
}
```

**Use Cases**:
- Understand formula impact before editing
- Identify cascading effects
- Prevent breaking dependent calculations
- Document formula relationships

**Errors**:
- Workbook not opened
- Worksheet not found
- Cell not found

---

### excel_explain_formula

Explains a formula in Vietnamese with step-by-step breakdown and optimization suggestions.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |
| `cellAddress` | string | Yes | Cell address containing formula |

**Returns**: Detailed explanation of formula logic.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary",
  "cellAddress": "B10"
}
```

**Example Response**:

```json
{
  "cell": "B10",
  "formula": "=SUM(B1:B9)",
  "description": "Tính tổng các giá trị từ ô B1 đến B9",
  "steps": [
    "Bước 1: Xác định vùng dữ liệu B1:B9",
    "Bước 2: Cộng tất cả các giá trị trong vùng này",
    "Bước 3: Kết quả lưu ở ô B10"
  ],
  "functions": [
    {
      "name": "SUM",
      "description": "Hàm SUM tính tổng các số",
      "arguments": ["B1:B9"]
    }
  ],
  "inputParameters": ["B1:B9"],
  "outputType": "number",
  "suggestions": [
    "Có thể dùng named range 'Revenues' thay vì B1:B9",
    "Xem xét sử dụng SUMIF nếu có điều kiện lọc"
  ]
}
```

**Explanation Components**:
- `description`: Formula purpose in Vietnamese
- `steps`: Step-by-step execution
- `functions`: Functions used with descriptions
- `inputParameters`: Input ranges/cells
- `outputType`: Expected result type
- `suggestions`: Optimization tips

**Use Cases**:
- Understand complex formulas
- Learn formula functions
- Get optimization suggestions
- Document formula logic

**Errors**:
- Workbook not opened
- Worksheet not found
- Cell does not contain formula

---

### excel_audit_formulas

Audits all formulas in a worksheet and provides comprehensive analysis report.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |

**Returns**: Comprehensive audit report.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary"
}
```

**Example Response**:

```json
{
  "worksheet": "Summary",
  "totalFormulas": 15,
  "totalCells": 100,
  "formulaDensity": "15%",
  "errors": [],
  "circularReferences": [],
  "complexity": {
    "low": 10,
    "medium": 4,
    "high": 1
  },
  "recommendations": [
    "Formula density is good (15%)",
    "Consider breaking down 1 high complexity formula",
    "No circular references detected"
  ],
  "formulasByType": {
    "SUM": 5,
    "AVERAGE": 3,
    "VLOOKUP": 2,
    "IF": 5
  }
}
```

**Audit Metrics**:
- `totalFormulas`: Total formula count
- `formulaDensity`: Percentage of cells with formulas
- `errors`: Formula errors detected
- `circularReferences`: Circular reference issues
- `complexity`: Distribution of formula complexity
- `recommendations`: Improvement suggestions
- `formulasByType`: Usage frequency of each function

**Use Cases**:
- Comprehensive formula review
- Identify formula issues
- Quality assessment
- Optimization opportunities

**Errors**:
- Workbook not opened
- Worksheet not found

---

### excel_check_circular_references

Checks entire worksheet for circular references and reports all instances found.

**Parameters**:

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `filename` | string | Yes | Name of the opened workbook |
| `worksheet` | string | Yes | Name of the worksheet |

**Returns**: Circular reference report.

**Example Request**:

```json
{
  "filename": "financial-report.xlsx",
  "worksheet": "Summary"
}
```

**Example Response**:

```json
{
  "hasCircularReferences": true,
  "circularCells": [
    {
      "cell": "B5",
      "chain": ["B5", "C5", "D5", "B5"]
    }
  ],
  "totalChecked": 15,
  "circularCount": 1
}
```

**Circular Reference Components**:
- `hasCircularReferences`: Whether any circular references exist
- `circularCells`: Array of cells with circular references
- `chain`: Reference chain showing the circular dependency
- `totalChecked`: Number of formula cells checked
- `circularCount`: Total circular references found

**Circular References**:
A circular reference occurs when a formula references its own cell, either directly or through a chain of dependencies. This causes calculation errors.

**Example Circular Chain**:
```
B5 contains: =C5+D5
C5 contains: =D5*2
D5 contains: =B5+10
Result: B5 → C5 → D5 → B5 (circular!)
```

**Use Cases**:
- Debug calculation errors
- Fix circular reference issues
- Validate worksheet integrity
- Troubleshoot incorrect results

**Errors**:
- Workbook not opened
- Worksheet not found

## Data Filtering Guide

### Introduction

The `excel_filter_data` tool is one of the most powerful tools in the Excel MCP Server. It allows you to filter large datasets based on multiple conditions, making it easy to extract specific records without needing to read the entire file and filter manually.

### Basic Usage Examples

#### Example 1: Filter by Exact Match (Single Condition)

Find all "Quản lý khu vực" (Area Managers) in an employee list:

```javascript
// First, open the workbook
await excel_open_workbook({ 
  filePath: "/tmp/employees_full_02.xlsx" 
});

// Filter for area managers
const result = await excel_filter_data({
  filename: "employees_full_02.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "N215",
  filters: [
    { column: "I", operator: "equals", value: "Quản lý khu vực" }
  ],
  hasHeader: true
});

// Result: Array of rows where column I equals "Quản lý khu vực"
// Returns header row + all matching data rows

// Don't forget to close the workbook
await excel_close_workbook({ filename: "employees_full_02.xlsx" });
```

#### Example 2: Filter by Contains (Text Search)

Find all products containing "Widget":

```javascript
const result = await excel_filter_data({
  filename: "products.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "D1000",
  filters: [
    { column: "A", operator: "contains", value: "Widget" }
  ]
});

// Returns all rows where column A contains "Widget" anywhere in the text
```

#### Example 3: Multiple Conditions (AND Logic)

Find products with category "Electronics" AND price > 100:

```javascript
const result = await excel_filter_data({
  filename: "products.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "D1000",
  filters: [
    { column: "B", operator: "equals", value: "Electronics" },
    { column: "C", operator: "greaterThan", value: "100" }
  ]
});

// Note: Multiple filters are combined with AND logic
// Only rows matching ALL conditions will be returned
```

#### Example 4: Numeric Comparisons

Find orders with amount between 1000 and 5000:

```javascript
const result = await excel_filter_data({
  filename: "orders.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "H500",
  filters: [
    { column: "D", operator: "greaterThanOrEqual", value: "1000" },
    { column: "D", operator: "lessThanOrEqual", value: "5000" }
  ]
});

// Returns orders where amount >= 1000 AND amount <= 5000
```

#### Example 5: Case Sensitive Filter

Find exact "Apple" (not "apple" or "APPLE"):

```javascript
const result = await excel_filter_data({
  filename: "fruits.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "C100",
  filters: [
    { column: "A", operator: "equals", value: "Apple", matchCase: true }
  ]
});

// Only returns exact "Apple" (case-sensitive)
```

#### Example 6: Find Non-Empty Cells

Find all records where email is provided:

```javascript
const result = await excel_filter_data({
  filename: "contacts.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "G500",
  filters: [
    { column: "E", operator: "isNotBlank" }
  ]
});

// Returns all rows where column E has any value
```

#### Example 7: Find Empty Cells

Find records missing phone numbers:

```javascript
const result = await excel_filter_data({
  filename: "contacts.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "G500",
  filters: [
    { column: "F", operator: "isBlank" }
  ]
});

// Returns all rows where column F is empty
```

#### Example 8: Complex Multi-Column Filter

Find employees in "Sales" department with salary >= 50000 and status = "Active":

```javascript
const result = await excel_filter_data({
  filename: "employees.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "J200",
  filters: [
    { column: "D", operator: "equals", value: "Sales" },
    { column: "G", operator: "greaterThanOrEqual", value: "50000" },
    { column: "H", operator: "equals", value: "Active" }
  ]
});

// Only returns employees matching ALL three conditions
```

#### Example 9: Filter by Text Exclusion

Find all products except those containing "Deprecated":

```javascript
const result = await excel_filter_data({
  filename: "products.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "D1000",
  filters: [
    { column: "A", operator: "notContains", value: "Deprecated" }
  ]
});

// Returns products that do NOT contain "Deprecated" in column A
```

#### Example 10: Date Range Filtering (as text)

Find orders from 2024 (assuming dates in column A as text):

```javascript
const result = await excel_filter_data({
  filename: "orders.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "H500",
  filters: [
    { column: "A", operator: "greaterThanOrEqual", value: "2024-01-01" },
    { column: "A", operator: "lessThan", value: "2025-01-01" }
  ]
});

// Returns orders from 2024 (date strings are compared lexicographically)
```

### Common Use Cases

#### Use Case 1: Employee Management

Find all managers in a specific department:

```javascript
const managers = await excel_filter_data({
  filename: "employees.xlsx",
  worksheet: "Employees",
  startCell: "A1",
  endCell: "L500",
  filters: [
    { column: "C", operator: "equals", value: "Manager" },
    { column: "B", operator: "equals", value: "Marketing" }
  ]
});

// Returns all Marketing department managers
```

#### Use Case 2: Sales Analysis

Find high-value sales (amount > 10000):

```javascript
const highValueSales = await excel_filter_data({
  filename: "sales.xlsx",
  worksheet: "Q1",
  startCell: "A1",
  endCell: "K2000",
  filters: [
    { column: "J", operator: "greaterThan", value: "10000" }
  ]
});

// Returns all sales over $10,000
```

#### Use Case 3: Inventory Management

Find products with low stock (quantity < 10):

```javascript
const lowStock = await excel_filter_data({
  filename: "inventory.xlsx",
  worksheet: "Stock",
  startCell: "A1",
  endCell: "D500",
  filters: [
    { column: "C", operator: "lessThan", value: "10" }
  ]
});

// Returns products that need restocking
```

#### Use Case 4: Customer Service

Find unresolved tickets from VIP customers:

```javascript
const vipTickets = await excel_filter_data({
  filename: "tickets.xlsx",
  worksheet: "Support",
  startCell: "A1",
  endCell: "M1000",
  filters: [
    { column: "E", operator: "equals", value: "VIP" },
    { column: "L", operator: "equals", value: "Open" }
  ]
});

// Returns open tickets from VIP customers (priority handling)
```

#### Use Case 5: Data Quality Check

Find records with missing critical information:

```javascript
const incompleteRecords = await excel_filter_data({
  filename: "data.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "K1000",
  filters: [
    { column: "B", operator: "isBlank" }
  ]
});

// Returns rows where column B is empty (data quality issue)
```

#### Use Case 6: Regional Analysis

Find all customers in specific regions:

```javascript
const westCoastCustomers = await excel_filter_data({
  filename: "customers.xlsx",
  worksheet: "CustomerList",
  startCell: "A1",
  endCell: "J1500",
  filters: [
    { column: "G", operator: "equals", value: "California" },
    { column: "G", operator: "equals", value: "Oregon" },
    { column: "G", operator: "equals", value: "Washington" }
  ]
});

// Note: This won't work as expected (OR logic not supported).
// For OR logic, you would need to run separate filters and combine results.
```

### Comparison: Filter vs Read Range

**Approach 1: Read Entire Range + Manual Filter (NOT RECOMMENDED)**

```javascript
// ❌ Inefficient - reads all data then filters manually
const allData = await excel_read_range({
  filename: "employees.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "N215"
});

// Then manually filter through all 215 rows
const areaManagers = allData.filter(row => 
  row.some(cell => cell.value === "Quản lý khu vực")
);
```

**Approach 2: Use excel_filter_data (RECOMMENDED)**

```javascript
// ✅ Efficient - filters at the source, returns only matching rows
const areaManagers = await excel_filter_data({
  filename: "employees.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "N215",
  filters: [
    { column: "I", operator: "equals", value: "Quản lý khu vực" }
  ],
  hasHeader: true
});
```

**Benefits of using excel_filter_data:**
- ✅ Faster performance (filtering happens server-side)
- ✅ Less memory usage (only returns matching rows)
- ✅ Cleaner code (no manual filtering logic)
- ✅ Better for large datasets

### Best Practices

1. **Always specify the correct data range**: Use `startCell` and `endCell` to limit the filter to your actual data area for better performance.

2. **Use `hasHeader: true`**: When your data has headers, this ensures the first row is included in results and properly labeled.

3. **Start with single filters**: Test with one condition first, then add more filters incrementally to debug issues.

4. **Consider data types**: For numeric comparisons, ensure the column contains numbers, not text stored as numbers.

5. **Use `contains` for partial matches**: When you're unsure of exact values or need flexibility, `contains` is more flexible than `equals`.

6. **Case sensitivity**: Default is case-insensitive. Only use `matchCase: true` when you need exact case matching.

7. **Filter before reading**: Always try to filter data first, then read the filtered results rather than reading everything and filtering manually.

### Performance Tips

- **Filter on smaller ranges when possible**: Instead of filtering 10,000 rows, filter only the relevant 1,000 rows.
- **Use specific conditions**: `equals` is faster than `contains` when you know the exact value.
- **Multiple filters**: All filters are evaluated together, so combining conditions doesn't add significant overhead.
- **Avoid OR logic**: Multiple filters use AND logic. For OR logic, run separate filter operations.

### Troubleshooting

**Problem: No results returned**
- Check that the column letter is correct (A, B, C, etc.)
- Verify the value matches exactly (use `contains` for partial matches)
- Ensure the data range includes all your data

**Problem: Too many results returned**
- Add more filter conditions to narrow down results
- Check that you're using the correct operator (e.g., `equals` vs `contains`)
- Verify the data range doesn't include unnecessary rows

**Problem: Numeric comparison not working**
- Ensure the column contains actual numbers, not text
- Try wrapping the value in quotes: `"100"` instead of `100`
- Check for spaces or special characters in the data

### Filter Operators Reference

| Operator | Description | Example | When to Use |
|----------|-------------|---------|-------------|
| `equals` | Exact match | `{column: "A", operator: "equals", value: "John"}` | When you know the exact value |
| `notEquals` | Not equal | `{column: "B", operator: "notEquals", value: "Inactive"}` | To exclude specific values |
| `contains` | Contains text | `{column: "A", operator: "contains", value: "Manager"}` | For partial text matching |
| `notContains` | Doesn't contain | `{column: "C", operator: "notContains", value: "Test"}` | To exclude text patterns |
| `greaterThan` | > comparison | `{column: "D", operator: "greaterThan", value: "100"}` | Numeric thresholds |
| `lessThan` | < comparison | `{column: "D", operator: "lessThan", value: "500"}` | Numeric limits |
| `greaterThanOrEqual` | >= comparison | `{column: "D", operator: "greaterThanOrEqual", value: "0"}` | Numeric minimum |
| `lessThanOrEqual` | <= comparison | `{column: "D", operator: "lessThanOrEqual", value: "1000"}` | Numeric maximum |
| `isBlank` | Cell is empty | `{column: "E", operator: "isBlank"}` | Find missing data |
| `isNotBlank` | Cell not empty | `{column: "E", operator: "isNotBlank"}` | Find populated cells |

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
  cellAddress: "B10",
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

### Filtering Data

```javascript
// 1. Open workbook
await excel_open_workbook({ filePath: "/data/employees.xlsx" });

// 2. Filter data (efficient - server-side filtering)
const managers = await excel_filter_data({
  filename: "employees.xlsx",
  worksheet: "Sheet1",
  startCell: "A1",
  endCell: "N500",
  filters: [
    { column: "I", operator: "equals", value: "Manager" }
  ],
  hasHeader: true
});

// 3. Process filtered results
console.log(`Found ${managers.length} managers`);

// 4. Close workbook
await excel_close_workbook({ filename: "employees.xlsx" });
```

## Rate Limits and Quotas

- Maximum file size: 50MB (configurable)
- Maximum concurrent workbooks: 10 (configurable)
- Operation timeout: 30 seconds (configurable)

## Versioning

Current API version: 1.0.0

All tools maintain backward compatibility. Breaking changes will result in a major version bump.