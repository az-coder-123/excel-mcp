# Excel MCP — Tools Reference

Comprehensive documentation for all 80+ tools provided by the Excel Model Context Protocol (MCP) server.

---

## 1. Core Workbook & Cell Operations

| Guide | Tools | Description |
|-------|-------|-------------|
| [Cell Operations](cell-operations.md) | `excel_read_cell`, `excel_read_range`, `excel_write_cell`, `excel_write_batch`, `excel_copy_range`, `excel_find_replace`, `excel_sort_range`, `excel_get_named_ranges`, `excel_add_named_range` | Cell reading, batch writes, search & replace, range sorting, named ranges |
| [Worksheet Operations](worksheet-operations.md) | `excel_list_worksheets`, `excel_add_worksheet`, `excel_delete_worksheet`, `excel_rename_worksheet`, `excel_copy_worksheet`, `excel_insert_rows`, `excel_insert_columns`, `excel_delete_rows`, `excel_delete_columns`, `excel_merge_cells`, `excel_unmerge_cells`, `excel_add_table`, `excel_add_filter`, `excel_remove_filter` | Worksheet lifecycle, rows/cols insertion & deletion, merging, tables, auto-filters |

---

## 2. Formatting & Layout

| Guide | Tools | Description |
|-------|-------|-------------|
| [Font & Text](font-and-text.md) | `excel_set_font_style`, `excel_set_font_name_size`, `excel_set_rich_text` | Bold, italic, underline, strike, font families, font sizes, rich text runs |
| [Alignment](alignment.md) | `excel_set_alignment`, `excel_center_text` | Horizontal, vertical, wrap text, shrink to fit, text rotation |
| [Borders](borders.md) | `excel_set_border`, `excel_apply_all_borders`, `excel_apply_outline_border` | Cell borders, grid borders, bounding outline borders |
| [Colors](colors.md) | `excel_set_background_color`, `excel_set_font_color`, `excel_set_gradient_fill` | Solid fills, font colors, 2-stop gradient fills |
| [Number Formats](number-formats.md) | `excel_set_number_format`, `excel_apply_currency_format`, `excel_apply_percentage_format`, `excel_apply_date_format` | Custom number formats, currency symbols, percentages, ISO/short dates |
| [Style Presets](style-presets.md) | `excel_apply_header_style`, `excel_apply_title_style`, `excel_apply_table_style` | Professional enterprise presets for table headers, sheet titles, zebra striping |
| [Layout & Validation](layout-and-validation.md) | `excel_auto_fit_columns`, `excel_set_column_width`, `excel_set_row_height`, `excel_freeze_panes`, `excel_set_print_area`, `excel_add_header_footer`, `excel_add_comment`, `excel_remove_comment`, `excel_add_hyperlink`, `excel_add_data_validation`, `excel_format_worksheet`, `excel_batch_format` | Dynamic column auto-fitting, freeze panes, print area, notes, hyperlinks, dropdown validation, batch formatting |

---

## 3. Data Processing & Analytics

| Guide | Tools | Description |
|-------|-------|-------------|
| [Data Utilities & Analytics](data-utilities.md) | `excel_import_csv`, `excel_export_csv`, `excel_find_duplicates`, `excel_count_unique_values`, `excel_highlight_duplicates`, `excel_get_duplicate_info`, `excel_remove_duplicates`, `excel_text_to_columns`, `excel_flash_fill`, `excel_split_data_to_worksheets`, `excel_get_unique_values`, `excel_vlookup`, `excel_index_match`, `excel_create_pivot_table`, `excel_get_column_stats`, `excel_filter_data`, `excel_group_aggregate`, `excel_profile_data`, `excel_search`, `excel_compare_ranges` | CSV transfer, duplicate removal, text splitting, flash fill, VLOOKUP/INDEX-MATCH, pivot summaries, column statistics, filtering, grouping, data profiling, search, range comparison |

---

## 4. Visualizations & Security

| Guide | Tools | Description |
|-------|-------|-------------|
| [Charts & Visuals](charts-and-visuals.md) | `excel_add_chart`, `excel_update_chart`, `excel_delete_chart`, `excel_list_charts`, `excel_add_conditional_format`, `excel_remove_conditional_format`, `excel_add_data_bar`, `excel_add_color_scale`, `excel_add_icon_set` | Column/bar/line/pie/donut charts, conditional format rules, data bars, color scale heatmaps, icon rating sets |
| [Worksheet Protection](protection.md) | `excel_protect_worksheet`, `excel_unprotect_worksheet`, `excel_protect_cells`, `excel_protect_workbook`, `excel_unprotect_workbook` | Password protection, input cell unlocking, worksheet lock policies |

---

## 5. Formula Engineering & Auditing

| Guide | Tools | Description |
|-------|-------|-------------|
| [Formula Analysis](formula-analysis.md) | `excel_list_formulas`, `excel_analyze_formula`, `excel_get_dependencies`, `excel_trace_precedents`, `excel_trace_dependents`, `excel_check_circular`, `excel_explain_formula`, `excel_audit_formulas` | Formula inventory, AST structure analysis, precedent/dependent graph tracing, circular reference detection, multi-language explanation, formula auditing |

---

## 6. Accounting & Corporate Finance

| Guide | Tools | Description |
|-------|-------|-------------|
| [Financial Calculations](financial-calculations.md) | `excel_financial_sum`, `excel_financial_average`, `excel_running_total`, `excel_percentage_of_total`, `excel_year_to_date` | Mathematical sum, average, cumulative running totals, percent of total, YTD |
| [Accounting Formats](accounting-formats.md) | `excel_accounting_format`, `excel_vnd_currency_format`, `excel_negative_red_format`, `excel_show_zeros_instead_of_empty` | GAAP/IFRS accounting formats, Vietnamese Dong (₫) format, red negative highlight |
| [Financial Analysis](financial-analysis.md) | `excel_period_comparison`, `excel_variance_analysis`, `excel_check_balance`, `excel_find_anomalies` | Period-over-period comparison, variance analysis, trial balance check, Z-score anomaly detection |
| [Investment Analysis](investment-analysis.md) | `excel_calculate_npv`, `excel_calculate_irr`, `excel_calculate_xirr`, `excel_calculate_financial_ratio` | Net Present Value (NPV), Internal Rate of Return (IRR), Exact Dates IRR (XIRR), liquidity/profitability ratios |
| [Loan & Debt](loan-and-debt.md) | `excel_create_amortization_schedule`, `excel_create_aging_report` | Loan amortization schedules (fixed payment & fixed principal), accounts receivable aging buckets |
| [Tax & Currency](tax-and-currency.md) | `excel_calculate_tax`, `excel_calculate_progressive_tax`, `excel_convert_currency` | Flat sales/VAT tax, progressive personal income tax tiers, multi-currency conversion |
| [Asset Depreciation](depreciation.md) | `excel_calculate_depreciation` | Straight-Line (SL), Double-Declining Balance (DDB), Sum-of-Years'-Digits (SYD) depreciation schedules |

---

## 7. General Reference

| Guide | Content |
|-------|---------|
| [Examples](examples.md) | End-to-end multi-step workflow examples for financial modeling and reporting |
| [Reference](reference.md) | ARGB hex color codes, border styles, font options, and coordinate syntax |