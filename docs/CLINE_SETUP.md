# Excel MCP Server - Cline Setup Guide

## Overview

This guide provides detailed instructions for setting up and configuring Excel MCP Server to work with Cline (Claude-powered AI assistant in VS Code).

## Prerequisites

- VS Code installed
- Cline extension installed
- Node.js (v16 or higher)
- Excel MCP Server project built and ready

## Quick Reference

### Open Cline MCP configuration file:
```bash
open "~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json"
```

### Build server:
```bash
cd "/Volumes/Macintosh HD - Data/projects/excel-mcp"
npm run build
```

### Check server health:
```json
{
  "name": "excel_health_check",
  "arguments": {}
}
```

### Reload Cline configuration:
- VS Code: `Cmd+Shift+P` → "Developer: Reload Window"

## Installation Steps

### 1. Build the Excel MCP Server

```bash
cd "/path/to/excel-mcp"
npm install
npm run build
```

This creates the `dist/index.js` file that will be used by Cline.

### 2. Locate Cline Configuration File

Cline stores MCP server configurations in:
```
~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json
```

Open this file in your editor:
```bash
open "~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json"
```

### 3. Configure Excel MCP Server

Add the following configuration to `cline_mcp_settings.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "autoApprove": [
        "excel_open_workbook"
      ],
      "disabled": false,
      "timeout": 60,
      "type": "stdio",
      "command": "node",
      "args": [
        "/path/to/excel-mcp/dist/index.js"
      ],
      "env": {
        "MCP_LOG_LEVEL": "info",
        "MCP_ALLOWED_PATHS": "/path/to/excel-mcp/*",
        "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*,/path/to/excel-mcp/node_modules/*"
      }
    }
  }
}
```

**Important:** Replace `/path/to/excel-mcp` with your actual project path.

## Configuration Parameters

| Parameter | Value | Description |
|-----------|-------|-------------|
| `autoApprove` | `["tool_name"]` | List of tools that don't require manual approval |
| `disabled` | `false` | Enable/disable the server (true = disabled) |
| `timeout` | `60` | Maximum wait time for server response (seconds) |
| `type` | `"stdio"` | Connection type (stdio = standard input/output) |
| `command` | `"node"` | Command to run the server |
| `args` | `[path]` | Path to the MCP server executable |
| `MCP_LOG_LEVEL` | `"info"` | Log level: debug, info, warn, error |
| `MCP_ALLOWED_PATHS` | `path/*` | Comma-separated allowed paths (use `/*` for all files) |
| `MCP_DENIED_PATHS` | `paths` | Comma-separated denied paths |

## Current Configuration Status

- **Server Path**: `/Volumes/Macintosh HD - Data/projects/excel-mcp/dist/index.js`
- **Server Status**: ✅ Active (disabled: false)
- **Log Level**: info
- **Timeout**: 60 seconds
- **Auto-approve Tool**: excel_open_workbook
- **Allowed Paths**: Project directory only
- **Denied Paths**: System directories and node_modules

## Advanced Configuration

### Configuring Auto-approve Tools

To avoid manual approval for frequently used tools, add them to the `autoApprove` list:

```json
{
  "autoApprove": [
    "excel_health_check",
    "excel_open_workbook",
    "excel_read_cell",
    "excel_read_range",
    "excel_list_worksheets",
    "excel_write_cell",
    "excel_save_workbook",
    "excel_close_workbook"
  ]
}
```

⚠️ **Security Warning**: Only auto-approve safe read/write operations. Keep dangerous tools like `excel_delete_worksheet` requiring manual approval.

### Changing Access Permissions

To allow access to additional directories:

```json
{
  "env": {
    "MCP_ALLOWED_PATHS": "/path/to/excel-mcp/*,/Users/yourname/Documents/*,/Users/yourname/Desktop/*"
  }
}
```

### Increasing Timeout for Large Files

For processing large Excel files, increase the timeout:

```json
{
  "timeout": 120
}
```

### Complete Configuration Example

```json
{
  "mcpServers": {
    "excel-mcp": {
      "autoApprove": [
        "excel_health_check",
        "excel_open_workbook",
        "excel_read_cell",
        "excel_read_range",
        "excel_list_worksheets",
        "excel_write_cell",
        "excel_save_workbook",
        "excel_close_workbook"
      ],
      "disabled": false,
      "timeout": 90,
      "type": "stdio",
      "command": "node",
      "args": [
        "/Volumes/Macintosh HD - Data/projects/excel-mcp/dist/index.js"
      ],
      "env": {
        "MCP_LOG_LEVEL": "info",
        "MCP_ALLOWED_PATHS": "/Volumes/Macintosh HD - Data/projects/excel-mcp/*,/Users/yourname/Documents/*",
        "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*,/Volumes/Macintosh HD - Data/projects/excel-mcp/node_modules/*",
        "MCP_MAX_FILE_SIZE": "209715200"
      }
    }
  }
}
```

### Temporarily Disabling the Server

To disable the server without removing configuration:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "disabled": true,
      ...
    }
  }
}
```

## Using Excel MCP with Cline

### Basic Workflow

1. **Open a Workbook**:
```json
{
  "name": "excel_open_workbook",
  "arguments": {
    "filePath": "/path/to/your/file.xlsx"
  }
}
```

2. **List Worksheets**:
```json
{
  "name": "excel_list_worksheets",
  "arguments": {
    "filename": "file.xlsx"
  }
}
```

3. **Read Data**:
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

4. **Write Data**:
```json
{
  "name": "excel_write_cell",
  "arguments": {
    "filename": "file.xlsx",
    "worksheet": "Sheet1",
    "cellAddress": "A1",
    "value": "New Value"
  }
}
```

5. **Save and Close**:
```json
{
  "name": "excel_save_workbook",
  "arguments": {
    "filename": "file.xlsx"
  }
}
```

## Troubleshooting

### Server Not Responding

**Check if server is running:**
```bash
ps aux | grep "node dist/index.js"
```

**Restart server:**
```bash
cd "/path/to/excel-mcp"
npm run build
```

**Verify configuration file path:**
```bash
test -f "/path/to/excel-mcp/dist/index.js" && echo "✅ File exists" || echo "❌ File not found"
```

### "Server not found" Error

**Possible causes:**
- Incorrect path to `dist/index.js`
- Server not built
- Server disabled in configuration

**Solutions:**

1. Rebuild the project:
```bash
cd "/path/to/excel-mcp"
npm run build
```

2. Check if server is disabled:
```json
{
  "mcpServers": {
    "excel-mcp": {
      "disabled": false
    }
  }
}
```

3. Check VS Code Output:
   - Open Output panel (`Cmd+Shift+U`)
   - Select "Cline MCP" from dropdown
   - Review error logs

### Access Denied Errors

**Check configuration:**
1. Open configuration file:
```bash
open "~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json"
```

2. Verify `MCP_ALLOWED_PATHS`:
   - Ensure Excel file is in an allowed path
   - Paths should end with `/*` to include all files

3. Check `MCP_DENIED_PATHS`:
   - Ensure file doesn't match denied patterns

**Correct configuration example:**
```json
"env": {
  "MCP_ALLOWED_PATHS": "/path/to/excel-mcp/*,/Users/yourname/Documents/*",
  "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*"
}
```

### "Workbook not opened" Error

**Possible causes:**
- `excel_open_workbook` not called first
- Incorrect file path
- File corrupted or doesn't exist

**Solutions:**

1. Open workbook before operations:
```json
{
  "name": "excel_open_workbook",
  "arguments": {
    "filePath": "/full/path/to/file.xlsx"
  }
}
```

2. Verify file exists:
```bash
test -f "/path/to/file.xlsx" && echo "✅ File exists" || echo "❌ File not found"
```

3. Use health check to verify server:
```json
{
  "name": "excel_health_check",
  "arguments": {}
}
```

### Timeout with Large Files

**Possible causes:**
- File too large
- Timeout too short

**Solutions:**

1. Increase timeout:
```json
{
  "timeout": 120
}
```

2. Increase file size limit:
```json
"env": {
  "MCP_MAX_FILE_SIZE": "209715200"
}
```

### Continuous Approval Requests

**Possible cause:**
- Tool not in `autoApprove` list

**Solution:**

1. Add tools to autoApprove:
```json
{
  "autoApprove": [
    "excel_health_check",
    "excel_open_workbook",
    "excel_read_cell",
    "excel_read_range",
    "excel_list_worksheets",
    "excel_write_cell",
    "excel_save_workbook",
    "excel_close_workbook"
  ]
}
```

2. Save file and reload VS Code

### No Debug Logs Visible

**Solution:**

1. Change log level:
```json
"env": {
  "MCP_LOG_LEVEL": "debug"
}
```

2. View logs in VS Code:
   - Open Output panel (`Cmd+Shift+U`)
   - Select "Cline MCP"

### Configuration Not Applied

**Possible causes:**
- JSON syntax error
- Cline not reloaded

**Solutions:**

1. Validate JSON:
```bash
cat "~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json" | jq .
```

2. Reload Cline:
   - `Cmd+Shift+P` → "Developer: Reload Window"
   - Or restart VS Code completely

## Environment Variables Reference

| Variable | Default | Description | Example |
|----------|---------|-------------|---------|
| `MCP_LOG_LEVEL` | info | Log level: debug, info, warn, error | `"debug"` |
| `MCP_ALLOWED_PATHS` | All paths | Comma-separated allowed paths | `"/path/*,/other/*"` |
| `MCP_DENIED_PATHS` | System paths | Comma-separated denied paths | `"/etc/*,/tmp/*"` |
| `MCP_MAX_FILE_SIZE` | 52428800 | Max file size in bytes (default 50MB) | `"104857600"` (100MB) |

## Verification

To verify the connection is working:

1. Ask Cline to check server health:
```
Please check if the Excel MCP server is running properly.
```

2. Cline should call:
```json
{
  "name": "excel_health_check",
  "arguments": {}
}
```

3. If successful, you'll see server status information.

## Best Practices

1. **Auto-approve carefully**: Only auto-approve safe operations
2. **Use specific paths**: Configure allowed paths precisely
3. **Monitor logs**: Enable debug logs when troubleshooting
4. **Close workbooks**: Always close workbooks when done
5. **Save regularly**: Use `excel_save_workbook` frequently
6. **Validate paths**: Always verify file paths before operations
7. **Handle errors**: Check for errors after each operation

## Next Steps

- Read [API Documentation](./API.md) for complete tool reference
- Review [Security Guide](./SECURITY.md) for security best practices
- Check [Troubleshooting](./TROUBLESHOOTING.md) for common issues
- Explore [Architecture](./ARCHITECTURE.md) to understand system design

## Support

If you encounter issues:
1. Check terminal logs when server runs
2. Check VS Code console for errors
3. Review this troubleshooting section
4. Check [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) for detailed solutions