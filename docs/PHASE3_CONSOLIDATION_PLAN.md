# Phase 3 (C1) — Tool Consolidation Blueprint: 127 → 26 Composite Tools

> **Status**: Blueprint approved for execution — requires a dedicated session.
> **Update (post-wiring)**: the 24 previously-undispatched tools (charts, conditional
> formatting, protection, CSV, lookups, pivot, flash-fill, text-to-columns,
> remove-duplicates, gradient) are now **fully implemented and dispatched**
> (127/127 defined tools live, covered by `tests/dead-tools-wiring.test.ts`).
> The composite mapping below absorbs them as noted — nothing is dropped.
> **Why dedicated**: C1 is a **breaking change** for every MCP client (all 127 tool
> names disappear) and mandates a full `docs/` + `README.md` overhaul in the same
> release. Executing it piecemeal would leave a hybrid API that is worse than either
> endpoint. The runtime groundwork is already merged: Zod validation (C2), LRU
> workbook manager (C3), and format-preserving export (Finding 6.2).

## Design rules

1. Each composite tool takes an **`action` enum** selecting the sub-operation.
2. Zod schemas are the **single source of truth** for both runtime validation and
   MCP JSON schemas (via `zod-to-json-schema`) — handcrafted `ToolParameter[]` are removed.
3. Handlers infer arg types via `z.infer<typeof Schema>` — zero `any`.
4. Business logic stays in the existing domain services (Facade untouched);
   only the `src/tools/` layer is rewritten.
5. `requiredPermissions` move to the composite level using the strictest
   permission of the merged actions.

## Mapping table

| # | Composite tool | `action` values | Absorbs (existing tools) |
|---|---|---|---|
| 1 | `excel_workbook_open` | — | open_workbook |
| 2 | `excel_workbook_create` | — | create_workbook |
| 3 | `excel_workbook_save` | — | save_workbook |
| 4 | `excel_workbook_close` | — | close_workbook |
| 5 | `excel_workbook_info` | — | get_workbook_context, list_worksheets |
| 6 | `excel_worksheet_manage` | add, delete, rename, copy | add/delete/rename/copy/list worksheet |
| 7 | `excel_worksheet_structure` | insert_rows, delete_rows, insert_columns, delete_columns, merge, unmerge, freeze, print_area, header_footer | structure tools (≈12) |
| 8 | `excel_cell_read` | cell, range | read_cell, read_range |
| 9 | `excel_cell_write` | cell, batch, copy | write_cell, write_batch, copy_range |
| 10 | `excel_find_replace` | — | find_replace |
| 11 | `excel_format_cells` | font, font_name_size, font_color, alignment, borders, outline_border, fill, gradient, number_format, cell_format, range_format, rich_text | formatting-tools + formatting-tools-extended (≈24) |
| 12 | `excel_format_presets` | header_style, title_style, table_style, format_worksheet, batch_format | preset tools (≈5) |
| 13 | `excel_conditional_format` | rule, databar, colorscale, iconset, remove | conditional-formatting-tools (5) |
| 14 | `excel_chart` | add, update, delete, list | chart-tools (4) |
| 15 | `excel_comment` | add, remove | comment tools (2) |
| 16 | `excel_hyperlink` | — | add_hyperlink |
| 17 | `excel_data_validation` | — | add_data_validation |
| 18 | `excel_data_transform` | sort, filter, remove_duplicates, text_to_columns, flash_fill, split_to_worksheets, group_aggregate | advanced-data-tools (≈7) |
| 19 | `excel_data_transfer` | import_csv, export_csv | CSV tools (2) |
| 20 | `excel_data_analyze` | profile, stats, search, compare, duplicates, unique_values, anomalies | analysis + data quality tools (≈7) |
| 21 | `excel_financial_basic` | sum, average, running_total, percentage, ytd, period_comparison, variance, balance, zeros | accounting-tools (≈9) |
| 22 | `excel_financial_format` | accounting, vnd, negative_red | format-variant tools (3) |
| 23 | `excel_financial_advanced` | npv, irr, ratio, amortization, aging, tax, currency | advanced-accounting-tools (7) |
| 24 | `excel_formula_analyze` | list, analyze, dependencies, precedents, dependents, explain, audit, circular | formula-analysis-tools (8) |
| 25 | `excel_protection` | protect_sheet, unprotect_sheet, protect_workbook, unprotect_workbook, lock_cells | protection-tools (≈5) |
| 26 | `excel_health_check` | — | health_check |

### Estimated context saving
127 tool definitions ≈ ~38k tokens of MCP schema payload → 26 composites with
action enums ≈ ~7–9k tokens (**~75% reduction**).

## Execution checklist (next session)

1. Create `src/tools/definitions/v2/` with Zod-first definitions (one file per family).
2. Rewrite `src/tools/handlers/` to `z.infer`-typed dispatchers over existing services.
3. Delete legacy `ToolParameter[]` definitions; flip `tool-definitions.ts` registry.
4. Generate MCP JSON schemas with `zod-to-json-schema`.
5. Port the dispatch switch in `tool-handler.ts` to a composite handler map.
6. Add per-composite Jest suites (action matrix + validation errors).
7. Update `README.md` tool tables and `docs/tools/*.md`; add a migration guide
   (`docs/MIGRATION_V2_TOOLS.md`) mapping every retired tool name → composite + action.
8. Baseline key collision fix (Finding 1.3 remainder): switch workbook keys to
   canonical paths with basename resolution — safe once the tool layer is rewritten.

## Non-goals (explicitly deferred)

- TTL-based eviction beyond the LRU cap (C3 already caps memory).
- Fail-closed default for empty `allowedPaths` (Finding 5.2) — pinned by test;
  changing it is an operator-facing config decision, not a code refactor.
