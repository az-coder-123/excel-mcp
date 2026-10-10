# Worksheet & Cell Protection

Tools for managing worksheet security, cell lock status, and password protection.

---

## `excel_protect_worksheet`

Protects a worksheet with an optional password, preventing unauthorized edits while allowing granular user interactions.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of worksheet to protect |
| `password` | ❌ | Protection password string |
| `allowSelectLockedCells` | ❌ | Allow users to select locked cells (default: `true`) |
| `allowSelectUnlockedCells` | ❌ | Allow users to select unlocked cells (default: `true`) |

**Example:**
```json
{
  "filename": "financial_report.xlsx",
  "worksheet": "Income Statement",
  "password": "AuditSecret2026",
  "allowSelectLockedCells": true,
  "allowSelectUnlockedCells": true
}
```

---

## `excel_unprotect_worksheet`

Removes protection from a worksheet.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Name of worksheet to unprotect |
| `password` | ❌ | Protection password |

---

## `excel_protect_cells`

Sets the `locked` protection flag on a specific cell range. When worksheet protection is enabled, locked cells cannot be modified by users while unlocked cells remain editable (ideal for input forms and assumptions).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Opened workbook filename |
| `worksheet` | ✅ | Worksheet name |
| `startCell` | ✅ | Range start cell (e.g., `B5`) |
| `endCell` | ✅ | Range end cell (e.g., `B10`) |
| `locked` | ✅ | `true` to lock cells, `false` to unlock |

**Example (Allow User Input in B5:B10):**
```json
{
  "filename": "model.xlsx",
  "worksheet": "Assumptions",
  "startCell": "B5",
  "endCell": "B10",
  "locked": false
}
```

---

## `excel_protect_workbook` & `excel_unprotect_workbook`

Workbook-level structure protection operations.

> [!NOTE]
> Workbook-level structural protection is not supported by the underlying ExcelJS serialization engine. The server returns a transparent engine-limitation error advising the caller to protect individual worksheets instead.

---

← [Back to index](README.md)
