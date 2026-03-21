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

#### For GitHub Copilot

Add to your VS Code `settings.json`:

```json
{
  "github.copilot.chat.experimental.mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["/path/to/excel-mcp/dist/index.js"],
      "env": {
        "MCP_LOG_LEVEL": "info"
      }
    }
  }
}
```

#### For Cline

Add to your Cline MCP configuration file (`~/.config/cline/mcp_servers.json`):

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["/path/to/excel-mcp/dist/index.js"],
      "env": {
        "MCP_LOG_LEVEL": "info"
      }
    }
  }
}
```

## Available Tools

### excel_open_workbook

Open an Excel file for operations.

```json
{
  "filePath": "/path/to/file.xlsx"
}
```

### excel_create_workbook

Create a new Excel workbook.

```json
{
  "filename": "report",
  "outputPath": "/workspace"
}
```

### excel_read_cell

Read a single cell value.

```json
{
  "filename": "report.xlsx",
  "worksheet": "Sheet1",
  "cellAddress": "A1"
}
```

### excel_read_range

Read multiple cells.

```json
{
  "filename": "report.xlsx",
  "worksheet": "Sheet1",
  "startCell": "A1",
  "endCell": "C10"
}
```

### excel_write_cell

Write a value to a cell.

```json
{
  "filename": "report.xlsx",
  "worksheet": "Sheet1",
  "cellAddress": "A1",
  "value": "Hello World"
}
```

### excel_list_worksheets

List all worksheets in a workbook.

```json
{
  "filename": "report.xlsx"
}
```

### excel_add_worksheet

Add a new worksheet.

```json
{
  "filename": "report.xlsx",
  "worksheetName": "Summary"
}
```

### excel_save_workbook

Save the workbook to disk.

```json
{
  "filename": "report.xlsx",
  "outputPath": "/workspace/report-final.xlsx"
}
```

### excel_close_workbook

Close an opened workbook.

```json
{
  "filename": "report.xlsx"
}
```

### excel_get_cell_info

Get detailed cell information including formulas.

```json
{
  "filename": "report.xlsx",
  "worksheet": "Sheet1",
  "cellAddress": "B5"
}
```

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