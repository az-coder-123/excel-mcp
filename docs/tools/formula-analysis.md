# Formula Analysis & Audit Tools

Tools for inspecting, explaining, auditing, and debugging Excel formulas and worksheet dependencies.

---

## `excel_list_formulas`

Lists all formula-containing cells in a worksheet, including their formulas and current cached values.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |

**Returns:**
```json
{
  "totalFormulas": 12,
  "formulas": [
    { "cell": "C10", "formula": "SUM(C2:C9)", "value": 15000 },
    { "cell": "D10", "formula": "AVERAGE(D2:D9)", "value": 1875 }
  ]
}
```

---

## `excel_analyze_formula`

Analyzes the syntactic structure, functions used, cell references, and complexity score of a formula in a specific cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Target cell address (e.g., `C10`) |

**Returns:**
```json
{
  "cell": "C10",
  "formula": "IF(SUM(A1:A5)>100, NPV(0.1, B1:B5), 0)",
  "functions": ["IF", "SUM", "NPV"],
  "references": ["A1:A5", "B1:B5"],
  "complexity": 5,
  "depth": 2
}
```

---

## `excel_get_dependencies`

Retrieves all direct cell and range references that a formula relies on.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Target cell address (e.g., `D5`) |

---

## `excel_trace_precedents`

Traces all upstream precedent cells that feed into the specified formula cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Target cell address (e.g., `D5`) |

---

## `excel_trace_dependents`

Traces all downstream cells in the worksheet whose values depend on the specified cell.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Target cell address (e.g., `B2`) |

---

## `excel_check_circular`

Detects circular reference loops in worksheet formulas (where a cell depends on itself directly or through a dependency cycle) using cycle detection on the dependency graph.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |

**Returns:**
```json
{
  "hasCircularReferences": false,
  "circularCount": 0,
  "cycles": []
}
```

---

## `excel_explain_formula`

Provides a human-readable explanation of an Excel formula in Vietnamese and English, breaking down each function component and referenced range.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |
| `cellAddress` | ✅ | Target cell address (e.g., `C10`) |

---

## `excel_audit_formulas`

Audits all worksheet formulas for potential issues, including syntax errors, inconsistent formulas across rows/columns, missing references, and high complexity.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of the worksheet |

**Returns:**
```json
{
  "auditedFormulas": 25,
  "issuesFound": 0,
  "issues": []
}
```

---

← [Back to index](README.md)
