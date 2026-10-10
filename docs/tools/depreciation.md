# Asset Depreciation

Tools for calculating fixed asset depreciation schedules using standard accounting methods.

## `excel_calculate_depreciation`

Calculate asset depreciation schedules using Straight-Line (SL), Double-Declining Balance (DDB), or Sum-of-Years'-Digits (SYD) methods. Writes a complete depreciation schedule table to the specified worksheet.

| Parameter | Required | Description |
|-----------|----------|-------------|
| `filename` | ✅ | Name of the opened workbook |
| `worksheet` | ✅ | Name of the worksheet |
| `startCell` | ✅ | Top-left cell where schedule table begins (e.g., `A1`) |
| `cost` | ✅ | Initial asset acquisition cost (must be > 0) |
| `salvageValue` | ✅ | Estimated salvage / residual value at end of life (>= 0, < cost) |
| `usefulLife` | ✅ | Useful life of the asset in years (positive integer) |
| `method` | ✅ | Depreciation method: `straight-line`, `double-declining`, or `sum-of-years-digits` |

### Depreciation Methods

| Method | Formula | Description |
|--------|---------|-------------|
| `straight-line` | `(Cost - Salvage) / Life` | Equal annual depreciation expense across all years. Standard for financial reporting. |
| `double-declining` | `BookValue * (2 / Life)` | Accelerated depreciation method. Annual expense decreases each year; guarded so book value never falls below salvage value. |
| `sum-of-years-digits` | `(Cost - Salvage) * (RemainingLife / Denom)` | Accelerated method where denominator is `Life * (Life + 1) / 2`. Higher depreciation in earlier years. |

### Output Schedule Table Structure

Starting at `startCell`, generates:
- **Header row**: `Year` | `Depreciation` | `Accumulated` | `Book Value`
- **Data rows**: One row per year with values formatted as currency (`#,##0.00`).

---

← [Back to index](README.md)
