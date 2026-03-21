# Excel MCP Server - Setup Guide

## Overview

Excel MCP Server is a Model Context Protocol server that enables interaction with Excel files through AI tools like Cline and GitHub Copilot.

## Prerequisites

- Node.js (v16 or higher)
- npm or yarn
- Access to VS Code with Cline extension or GitHub Copilot

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/az-coder-123/excel-mcp.git
cd excel-mcp
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Build the Server

```bash
npm run build
```

This will compile the TypeScript code and create the `dist/` folder with the built server.

## Quick Reference

### Build server:
```bash
cd "/path/to/excel-mcp"
npm run build
```

### Check server health:
```json
{
  "name": "excel_health_check",
  "arguments": {}
}
```

## Available Tools

### 1. System Operations
- `excel_health_check` - Check server status and configuration

### 2. Workbook Operations
- `excel_open_workbook` - Open an Excel file
- `excel_create_workbook` - Create a new workbook
- `excel_save_workbook` - Save workbook to disk
- `excel_close_workbook` - Close an opened workbook
- `excel_list_worksheets` - List all worksheets

### 3. Cell Operations
- `excel_read_cell` - Read value from a specific cell
- `excel_write_cell` - Write value to a specific cell
- `excel_read_range` - Read values from a range of cells
- `excel_write_batch` - Write multiple values at once
- `excel_get_cell_info` - Get detailed cell information

### 4. Worksheet Operations
- `excel_add_worksheet` - Add a new worksheet
- `excel_delete_worksheet` - Delete a worksheet
- `excel_rename_worksheet` - Rename a worksheet
- `excel_copy_worksheet` - Copy a worksheet

### 5. Formatting Operations
- `excel_set_cell_format` - Apply formatting to cells
- `excel_set_range_format` - Apply formatting to a range
- `excel_auto_fit_columns` - Auto-fit column widths
- `excel_set_column_width` - Set specific column width
- `excel_set_row_height` - Set specific row height

### 6. Advanced Operations
- `excel_add_chart` - Create charts
- `excel_add_table` - Create Excel tables
- `excel_add_filter` - Add auto-filter
- `excel_merge_cells` - Merge cells
- `excel_add_data_validation` - Add data validation rules
- `excel_add_conditional_format` - Add conditional formatting

### 7. Data Analysis
- `excel_filter_data` - Filter data based on conditions
- `excel_group_aggregate` - Group and aggregate data
- `excel_find_duplicates` - Find duplicate values
- `excel_profile_data` - Analyze data profile
- `excel_search` - Advanced search in worksheet

## Usage Examples

### Reading an Excel File

```json
{
  "name": "excel_open_workbook",
  "arguments": {
    "filePath": "/path/to/file.xlsx"
  }
}
```

```json
{
  "name": "excel_read_range",
  "arguments": {
    "filename": "file.xlsx",
    "worksheet": "Sheet1",
    "startCell": "A1",
    "endCell": "C10"
  }
}
```

### Creating a New Workbook

```json
{
  "name": "excel_create_workbook",
  "arguments": {
    "filename": "report",
    "outputPath": "/path/to/output"
  }
}
```

```json
{
  "name": "excel_write_cell",
  "arguments": {
    "filename": "report.xlsx",
    "worksheet": "Sheet1",
    "cellAddress": "A1",
    "value": "Hello World"
  }
}
```

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_LOG_LEVEL` | info | Log level: debug, info, warn, error |
| `MCP_ALLOWED_PATHS` | All paths | Comma-separated list of allowed paths |
| `MCP_DENIED_PATHS` | System paths | Comma-separated list of denied paths |
| `MCP_MAX_FILE_SIZE` | 52428800 | Maximum file size in bytes (default 50MB) |

## Security

The server implements several security measures:

- **Path-based access control**: Only allows access to specified paths
- **Denied paths**: Blocks access to system directories
- **File size limits**: Prevents processing of oversized files
- **Permission checks**: Validates operations before execution

For detailed security information, see [SECURITY.md](./SECURITY.md).

## Troubleshooting

### Server Not Responding

1. Check if the server is running:
```bash
ps aux | grep "node dist/index.js"
```

2. Rebuild the project:
```bash
npm run build
```

### Access Denied Errors

1. Verify file paths in configuration
2. Ensure files are in allowed paths
3. Check denied paths configuration

### Workbook Not Opened

1. Ensure `excel_open_workbook` is called first
2. Verify file path is correct
3. Check file is not corrupted

For more troubleshooting information, see [TROUBLESHOOTING.md](./TROUBLESHOOTING.md).

## Documentation

- [Cline Setup Guide](./CLINE_SETUP.md) - Setup instructions for Cline
- [GitHub Copilot Setup Guide](./GITHUB_COPILOT_SETUP.md) - Setup instructions for GitHub Copilot
- [API Documentation](./API.md) - Complete API reference
- [Architecture](./ARCHITECTURE.md) - System architecture and design
- [Security Guide](./SECURITY.md) - Security best practices
- [Troubleshooting](./TROUBLESHOOTING.md) - Common issues and solutions

## Support

If you encounter issues:
1. Check the terminal logs when server is running
2. Check VS Code console for errors
3. Review [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)

## License

See LICENSE file for details.