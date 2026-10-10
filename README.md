# Excel MCP Server

A Model Context Protocol (MCP) server for Excel operations with enterprise-grade security and clean architecture. Seamlessly integrates with GitHub Copilot, Cline, and other MCP-compatible tools in VS Code.

## Features

- **Complete Excel Operations**: Read, write, and manipulate Excel files (.xlsx, .xls)
- **Clean Architecture**: Follows Single Responsibility Principle for maintainability
- **Enterprise Security**: Robust permission system with path restrictions and access control
- **VS Code Integration**: Works seamlessly with GitHub Copilot and Cline
- **TypeScript**: Full type safety with modern ES2022+ features
- **Zero Deprecated Dependencies**: Uses only current, actively maintained libraries

## Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/az-coder-123/excel-mcp.git
cd excel-mcp

# Install dependencies
npm install

# Build the project
npm run build

# Copy example environment file
cp .env.example .env
```

### 📌 Đã cấu hình sẵn cho Cline

Server đã được cấu hình và kết nối sẵn với Cline! Xem [Hướng dẫn cài đặt chi tiết](docs/SETUP_GUIDE.md) để biết:

✅ **Trạng thái hiện tại**: Đã kết nối và hoạt động bình thường
- Server đang chạy: `node dist/index.js`
- Cấu hình Cline: `.cline/mcp_servers.json`
- File cấu hình đã sẵn sàng sử dụng

**Xem ngay**: [docs/SETUP_GUIDE.md](docs/SETUP_GUIDE.md) để biết cách sử dụng

### Configuration

Edit `.env` file to customize permissions:

```env
# Restrict access to specific directories
MCP_ALLOWED_PATHS=/home/user/documents,/workspace

# Block sensitive system directories
MCP_DENIED_PATHS=/etc/*,/sys/*,C:\Windows\*

# Set maximum file size (50MB default)
MCP_MAX_FILE_SIZE=52428800
```

### VS Code Integration

This server works seamlessly with MCP-compatible tools in VS Code. See the dedicated setup guides for detailed configuration instructions:

- **[Google Antigravity Setup](docs/ANTIGRAVITY_SETUP.md)** — Configuration for Google Antigravity IDE and `agy` CLI.
- **[GitHub Copilot Setup](docs/GITHUB_COPILOT_SETUP.md)** — Configuration for GitHub Copilot Chat in VS Code.
- **[Cline Setup](docs/CLINE_SETUP.md)** — Configuration for the Cline extension in VS Code.

## Available Tools

For detailed parameter references, see the [Tools Reference](docs/tools/README.md).

### System

| Tool | Description |
|------|-------------|
| `excel_health_check` | Check server health, dependencies, and configuration |

### Workbook Operations

| Tool | Description |
|------|-------------|
| `excel_open_workbook` | Open an Excel file for operations |
| `excel_create_workbook` | Create a new Excel workbook |
| `excel_save_workbook` | Save the workbook to disk |
| `excel_close_workbook` | Close an opened workbook |
| `excel_get_workbook_context` | Get workbook context and state |
| `excel_export_worksheet_to_new_file` | Export worksheet to a new Excel file |

### Worksheet Operations → [Detailed docs](docs/tools/worksheet-operations.md)

| Tool | Description |
|------|-------------|
| `excel_list_worksheets` | List all worksheets in a workbook |
| `excel_add_worksheet` | Add a new worksheet |
| `excel_delete_worksheet` | Delete a worksheet |
| `excel_rename_worksheet` | Rename a worksheet |
| `excel_copy_worksheet` | Copy a worksheet |
| `excel_insert_rows` | Insert rows at position |
| `excel_delete_rows` | Delete rows from position |
| `excel_insert_columns` | Insert columns at letter |
| `excel_delete_columns` | Delete columns from letter |
| `excel_merge_cells` | Merge a cell range |
| `excel_unmerge_cells` | Unmerge cells |
| `excel_add_table` | Add styled Excel table |
| `excel_add_filter` | Add auto-filter to range |
| `excel_remove_filter` | Remove auto-filter |

### Cell Operations → [Detailed docs](docs/tools/cell-operations.md)

| Tool | Description |
|------|-------------|
| `excel_read_cell` | Read a single cell value |
| `excel_read_range` | Read multiple cells |
| `excel_write_cell` | Write a value to a cell |
| `excel_write_batch` | Write multiple cells at once |
| `excel_get_cell_info` | Get detailed cell information (type, formula) |
| `excel_copy_range` | Copy a cell range to another location |
| `excel_find_replace` | Find and replace text |
| `excel_sort_range` | Sort a range ascending/descending |
| `excel_get_named_ranges` | List defined named ranges |
| `excel_add_named_range` | Create a new named range |

### Formatting & Styles → [Detailed docs](docs/tools/README.md#2-formatting--layout)

| Tool | Category | Docs |
|------|----------|------|
| `excel_set_font_style` | Font & Text | [→](docs/tools/font-and-text.md) |
| `excel_set_font_name_size` | Font & Text | [→](docs/tools/font-and-text.md) |
| `excel_set_rich_text` | Font & Text | [→](docs/tools/font-and-text.md) |
| `excel_set_alignment` | Alignment | [→](docs/tools/alignment.md) |
| `excel_center_text` | Alignment | [→](docs/tools/alignment.md) |
| `excel_set_border` | Borders | [→](docs/tools/borders.md) |
| `excel_apply_all_borders` | Borders | [→](docs/tools/borders.md) |
| `excel_apply_outline_border` | Borders | [→](docs/tools/borders.md) |
| `excel_set_background_color` | Colors | [→](docs/tools/colors.md) |
| `excel_set_font_color` | Colors | [→](docs/tools/colors.md) |
| `excel_set_gradient_fill` | Colors | [→](docs/tools/colors.md) |
| `excel_set_number_format` | Number Formats | [→](docs/tools/number-formats.md) |
| `excel_apply_currency_format` | Number Formats | [→](docs/tools/number-formats.md) |
| `excel_apply_percentage_format` | Number Formats | [→](docs/tools/number-formats.md) |
| `excel_apply_date_format` | Number Formats | [→](docs/tools/number-formats.md) |
| `excel_apply_header_style` | Style Presets | [→](docs/tools/style-presets.md) |
| `excel_apply_title_style` | Style Presets | [→](docs/tools/style-presets.md) |
| `excel_apply_table_style` | Style Presets | [→](docs/tools/style-presets.md) |
| `excel_auto_fit_columns` | Layout & Dimensions | [→](docs/tools/layout-and-validation.md) |
| `excel_set_column_width` | Layout & Dimensions | [→](docs/tools/layout-and-validation.md) |
| `excel_set_row_height` | Layout & Dimensions | [→](docs/tools/layout-and-validation.md) |
| `excel_freeze_panes` | Layout & Dimensions | [→](docs/tools/layout-and-validation.md) |
| `excel_set_print_area` | Layout & Dimensions | [→](docs/tools/layout-and-validation.md) |
| `excel_add_header_footer` | Layout & Dimensions | [→](docs/tools/layout-and-validation.md) |
| `excel_add_comment` | Notes & Links | [→](docs/tools/layout-and-validation.md) |
| `excel_remove_comment` | Notes & Links | [→](docs/tools/layout-and-validation.md) |
| `excel_add_hyperlink` | Notes & Links | [→](docs/tools/layout-and-validation.md) |
| `excel_add_data_validation` | Validation | [→](docs/tools/layout-and-validation.md) |
| `excel_format_worksheet` | Batch Formatting | [→](docs/tools/layout-and-validation.md) |
| `excel_batch_format` | Batch Formatting | [→](docs/tools/layout-and-validation.md) |

### Data Utilities & Analytics → [Detailed docs](docs/tools/data-utilities.md)

| Tool | Category | Docs |
|------|----------|------|
| `excel_import_csv` | CSV Transfer | [→](docs/tools/data-utilities.md) |
| `excel_export_csv` | CSV Transfer | [→](docs/tools/data-utilities.md) |
| `excel_find_duplicates` | Duplicates | [→](docs/tools/data-utilities.md) |
| `excel_count_unique_values` | Duplicates | [→](docs/tools/data-utilities.md) |
| `excel_highlight_duplicates` | Duplicates | [→](docs/tools/data-utilities.md) |
| `excel_remove_duplicates` | Duplicates | [→](docs/tools/data-utilities.md) |
| `excel_text_to_columns` | Transformations | [→](docs/tools/data-utilities.md) |
| `excel_flash_fill` | Transformations | [→](docs/tools/data-utilities.md) |
| `excel_vlookup` | Lookups | [→](docs/tools/data-utilities.md) |
| `excel_index_match` | Lookups | [→](docs/tools/data-utilities.md) |
| `excel_create_pivot_table` | Pivot Tables | [→](docs/tools/data-utilities.md) |
| `excel_get_column_stats` | Analytics | [→](docs/tools/data-utilities.md) |
| `excel_filter_data` | Analytics | [→](docs/tools/data-utilities.md) |
| `excel_group_aggregate` | Analytics | [→](docs/tools/data-utilities.md) |
| `excel_profile_data` | Analytics | [→](docs/tools/data-utilities.md) |
| `excel_search` | Search | [→](docs/tools/data-utilities.md) |
| `excel_compare_ranges` | Range Comparison | [→](docs/tools/data-utilities.md) |

### Visualizations & Protection → [Charts](docs/tools/charts-and-visuals.md) | [Protection](docs/tools/protection.md)

| Tool | Category | Docs |
|------|----------|------|
| `excel_add_chart` | Charts | [→](docs/tools/charts-and-visuals.md) |
| `excel_update_chart` | Charts | [→](docs/tools/charts-and-visuals.md) |
| `excel_delete_chart` | Charts | [→](docs/tools/charts-and-visuals.md) |
| `excel_list_charts` | Charts | [→](docs/tools/charts-and-visuals.md) |
| `excel_add_conditional_format` | Conditional Formats | [→](docs/tools/charts-and-visuals.md) |
| `excel_remove_conditional_format` | Conditional Formats | [→](docs/tools/charts-and-visuals.md) |
| `excel_add_data_bar` | Visual Rules | [→](docs/tools/charts-and-visuals.md) |
| `excel_add_color_scale` | Visual Rules | [→](docs/tools/charts-and-visuals.md) |
| `excel_add_icon_set` | Visual Rules | [→](docs/tools/charts-and-visuals.md) |
| `excel_protect_worksheet` | Protection | [→](docs/tools/protection.md) |
| `excel_unprotect_worksheet` | Protection | [→](docs/tools/protection.md) |
| `excel_protect_cells` | Protection | [→](docs/tools/protection.md) |

### Formula Engineering & Auditing → [Detailed docs](docs/tools/formula-analysis.md)

| Tool | Category | Docs |
|------|----------|------|
| `excel_list_formulas` | Inventory | [→](docs/tools/formula-analysis.md) |
| `excel_analyze_formula` | Structure Analysis | [→](docs/tools/formula-analysis.md) |
| `excel_get_dependencies` | Dependencies | [→](docs/tools/formula-analysis.md) |
| `excel_trace_precedents` | Precedent Tracing | [→](docs/tools/formula-analysis.md) |
| `excel_trace_dependents` | Dependent Tracing | [→](docs/tools/formula-analysis.md) |
| `excel_check_circular` | Cycle Detection | [→](docs/tools/formula-analysis.md) |
| `excel_explain_formula` | Multi-language Explanation | [→](docs/tools/formula-analysis.md) |
| `excel_audit_formulas` | Automated Audit | [→](docs/tools/formula-analysis.md) |

### Accounting & Finance → [Detailed docs](docs/tools/README.md#6-accounting--corporate-finance)

| Tool | Category | Docs |
|------|----------|------|
| `excel_financial_sum` | Financial Calculations | [→](docs/tools/financial-calculations.md) |
| `excel_financial_average` | Financial Calculations | [→](docs/tools/financial-calculations.md) |
| `excel_running_total` | Financial Calculations | [→](docs/tools/financial-calculations.md) |
| `excel_percentage_of_total` | Financial Calculations | [→](docs/tools/financial-calculations.md) |
| `excel_year_to_date` | Financial Calculations | [→](docs/tools/financial-calculations.md) |
| `excel_accounting_format` | Accounting Formats | [→](docs/tools/accounting-formats.md) |
| `excel_vnd_currency_format` | Accounting Formats | [→](docs/tools/accounting-formats.md) |
| `excel_negative_red_format` | Accounting Formats | [→](docs/tools/accounting-formats.md) |
| `excel_show_zeros_instead_of_empty` | Accounting Formats | [→](docs/tools/accounting-formats.md) |
| `excel_period_comparison` | Financial Analysis | [→](docs/tools/financial-analysis.md) |
| `excel_variance_analysis` | Financial Analysis | [→](docs/tools/financial-analysis.md) |
| `excel_check_balance` | Financial Analysis | [→](docs/tools/financial-analysis.md) |
| `excel_find_anomalies` | Financial Analysis | [→](docs/tools/financial-analysis.md) |
| `excel_calculate_npv` | Investment Analysis | [→](docs/tools/investment-analysis.md) |
| `excel_calculate_irr` | Investment Analysis | [→](docs/tools/investment-analysis.md) |
| `excel_calculate_xirr` | Investment Analysis | [→](docs/tools/investment-analysis.md) |
| `excel_calculate_financial_ratio` | Investment Analysis | [→](docs/tools/investment-analysis.md) |
| `excel_create_amortization_schedule` | Loan & Debt | [→](docs/tools/loan-and-debt.md) |
| `excel_create_aging_report` | Loan & Debt | [→](docs/tools/loan-and-debt.md) |
| `excel_calculate_tax` | Tax & Currency | [→](docs/tools/tax-and-currency.md) |
| `excel_calculate_progressive_tax` | Tax & Currency | [→](docs/tools/tax-and-currency.md) |
| `excel_convert_currency` | Tax & Currency | [→](docs/tools/tax-and-currency.md) |
| `excel_calculate_depreciation` | Asset Depreciation | [→](docs/tools/depreciation.md) |

### Usage Examples & Reference

- [Examples](docs/tools/examples.md) — Practical usage examples for all tool groups
- [Reference](docs/tools/reference.md) — Color codes, border styles, font names, AARRGGBB format

## Security Features

### Permission Levels

- **read**: Read-only access to Excel files
- **write**: Create and modify Excel files
- **delete**: Delete worksheets and clear data
- **admin**: Full administrative access

### Path Restrictions

Configure allowed and denied paths using wildcards:

```env
# Allow only specific directories
MCP_ALLOWED_PATHS=/workspace/*,/home/user/documents/*

# Block sensitive system paths
MCP_DENIED_PATHS=/etc/*,/sys/*,/proc/*,C:\Windows\*
```

### File Size Limits

Prevent processing of excessively large files:

```env
MCP_MAX_FILE_SIZE=52428800  # 50MB
```

### Extension Whitelist

Only process specific file types (configured by default):

- `.xlsx` - Excel Workbook
- `.xls` - Excel 97-2003 Workbook
- `.xlsm` - Excel Macro-Enabled Workbook
- `.xlsb` - Excel Binary Workbook

## Architecture

```
excel-mcp/
├── src/
│   ├── index.ts                 # Application entry point
│   ├── types/
│   │   └── index.ts            # Type definitions
│   ├── security/
│   │   └── permission-checker.ts # Access control
│   ├── services/
│   │   └── excel-service.ts    # Excel operations
│   ├── tools/
│   │   ├── tool-definitions.ts # Tool schemas
│   │   └── tool-handler.ts     # Tool execution
│   ├── server/
│   │   └── excel-mcp-server.ts # MCP server core
│   └── utils/
│       └── logger.ts           # Logging utility
├── docs/
│   ├── AI_WORKFLOW.md          # AI agent SOP & collaboration guidelines
│   ├── ARCHITECTURE.md         # Architecture details
│   ├── API.md                  # API reference
│   ├── SECURITY.md             # Security guide
│   └── TROUBLESHOOTING.md      # Common issues
└── dist/                       # Compiled JavaScript
```

### Design Principles

1. **Single Responsibility**: Each class/module has one clear purpose
2. **Dependency Injection**: Components are loosely coupled
3. **Type Safety**: Full TypeScript coverage with strict mode
4. **Error Handling**: Comprehensive error catching and reporting
5. **Security First**: Permission checks before all operations

## Development

### Scripts

```bash
# Development with hot reload
npm run dev

# Build for production
npm run build

# Run tests
npm test

# Lint code
npm run lint

# Format code
npm run format

# Clean build artifacts
npm run clean
```

### Project Structure

- `src/` - TypeScript source files
- `dist/` - Compiled JavaScript output
- `docs/` - Documentation
- `package.json` - Dependencies and scripts
- `tsconfig.json` - TypeScript configuration

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `MCP_SERVER_NAME` | Server identifier | `excel-mcp` |
| `MCP_SERVER_VERSION` | Server version | `1.0.0` |
| `MCP_LOG_LEVEL` | Logging verbosity | `info` |
| `MCP_ALLOWED_PATHS` | Comma-separated allowed paths | All paths |
| `MCP_DENIED_PATHS` | Comma-separated denied paths | System paths |
| `MCP_MAX_FILE_SIZE` | Maximum file size in bytes | `52428800` (50MB) |

## Troubleshooting

### Common Issues

**"Permission denied" errors**
- Check your `MCP_ALLOWED_PATHS` configuration
- Ensure the file isn't in a denied path
- Verify you have the required permission level

**"Workbook not opened" errors**
- Call `excel_open_workbook` before other operations
- Check that the file path is correct
- Ensure the file isn't corrupted

**TypeScript compilation errors**
- Run `npm install` to ensure all dependencies are installed
- Check Node.js version (requires >= 18.0.0)

For more troubleshooting help, see [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md).

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please read our contributing guidelines and code of conduct.

## Support

- GitHub Issues: [Report bugs](https://github.com/az-coder-123/excel-mcp/issues)
- Documentation: [Full docs](docs/)
- Tools Reference: [Tools index](docs/tools/README.md)
