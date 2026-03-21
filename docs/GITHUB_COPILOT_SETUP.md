# Excel MCP Server - GitHub Copilot Setup Guide

## Overview

This guide provides detailed instructions for setting up and configuring Excel MCP Server to work with GitHub Copilot in VS Code.

## Prerequisites

- VS Code installed
- GitHub Copilot extension installed
- GitHub Copilot Chat extension installed (required for MCP)
- Node.js (v16 or higher)
- Excel MCP Server project built and ready

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

### Open VS Code settings:
- `Cmd+,` (Mac) or `Ctrl+,` (Windows/Linux)
- Search for "MCP" or "Copilot"

## Installation Steps

### 1. Build Excel MCP Server

```bash
cd "/path/to/excel-mcp"
npm install
npm run build
```

This creates the `dist/index.js` file that will be used by GitHub Copilot.

### 2. Locate GitHub Copilot Configuration File

GitHub Copilot stores MCP server configurations in different locations depending on your OS:

**macOS:**
```
~/Library/Application Support/Copilot/user/settings.json
```

**Windows:**
```
%APPDATA%/Copilot/user/settings.json
```

**Linux:**
```
~/.config/Copilot/user/settings.json
```

Open this file in your editor:

**macOS:**
```bash
open ~/Library/Application\ Support/Copilot/user/settings.json
```

**Windows (PowerShell):**
```powershell
notepad $env:APPDATA\Copilot\user\settings.json
```

**Linux:**
```bash
nano ~/.config/Copilot/user/settings.json
```

### 3. Configure Excel MCP Server

Add the following configuration to the `mcpServers` section in `settings.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "transport": {
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
}
```

**Important:** Replace `/path/to/excel-mcp` with your actual project path.

If the file doesn't exist or is empty, create it with the above content.

### 4. Restart VS Code

After making changes to the configuration file, restart VS Code for the changes to take effect:

- **macOS/Linux:** Close and reopen VS Code
- **Windows:** Close and reopen VS Code, or use `Ctrl+Shift+P` → "Developer: Reload Window"

## Configuration Parameters

| Parameter | Value | Description |
|-----------|-------|-------------|
| `transport.type` | `"stdio"` | Transport type (stdio for local servers) |
| `transport.command` | `"node"` | Command to run the server |
| `transport.args` | `[path]` | Path to the MCP server executable |
| `transport.env.MCP_LOG_LEVEL` | `"info"` | Log level: debug, info, warn, error |
| `transport.env.MCP_ALLOWED_PATHS` | `path/*` | Comma-separated allowed paths (use `/*` for all files) |
| `transport.env.MCP_DENIED_PATHS` | `paths` | Comma-separated denied paths |

## Advanced Configuration

### Changing Access Permissions

To allow access to additional directories:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "transport": {
        "type": "stdio",
        "command": "node",
        "args": [
          "/path/to/excel-mcp/dist/index.js"
        ],
        "env": {
          "MCP_LOG_LEVEL": "info",
          "MCP_ALLOWED_PATHS": "/path/to/excel-mcp/*,/Users/yourname/Documents/*,/Users/yourname/Desktop/*",
          "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*"
        }
      }
    }
  }
}
```

### Enabling Debug Logging

To enable detailed logging for troubleshooting:

```json
{
  "env": {
    "MCP_LOG_LEVEL": "debug"
  }
}
```

### Increasing File Size Limit

To allow processing larger Excel files:

```json
{
  "env": {
    "MCP_MAX_FILE_SIZE": "209715200"
  }
}
```

### Complete Configuration Example

```json
{
  "mcpServers": {
    "excel-mcp": {
      "transport": {
        "type": "stdio",
        "command": "node",
        "args": [
          "/path/to/excel-mcp/dist/index.js"
        ],
        "env": {
          "MCP_LOG_LEVEL": "info",
          "MCP_ALLOWED_PATHS": "/path/to/excel-mcp/*,/Users/yourname/Documents/*",
          "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*,/path/to/excel-mcp/node_modules/*",
          "MCP_MAX_FILE_SIZE": "209715200"
        }
      }
    }
  }
}
```

## Using Excel MCP with GitHub Copilot

### Accessing Excel Tools

1. Open GitHub Copilot Chat in VS Code:
   - Click the Copilot Chat icon in the sidebar
   - Or use shortcut: `Ctrl+Shift+I` (Windows/Linux) or `Cmd+Shift+I` (Mac)

2. Ask Copilot to use Excel tools:
```
Please open the Excel file at /path/to/file.xlsx
```

3. Copilot will automatically use the Excel MCP tools when appropriate.

### Basic Workflow

1. **Open a Workbook**:
```
Open the Excel file located at /path/to/your/file.xlsx
```

2. **List Worksheets**:
```
List all worksheets in the file.xlsx workbook
```

3. **Read Data**:
```
Read the data range from A1 to C10 in Sheet1
```

4. **Write Data**:
```
Write "New Value" to cell A1 in Sheet1
```

5. **Save and Close**:
```
Save the workbook and close it
```

### Example Conversations

**Reading Excel Data:**
```
You: Can you read the sales data from my Excel file?
Copilot: I'll help you read the sales data. Let me open the Excel file first.
[Calls excel_open_workbook]
Copilot: I've opened the file. Now let me read the sales data from Sheet1.
[Calls excel_read_range]
Copilot: Here's the sales data from cells A1:C10:
...
```

**Creating a Report:**
```
You: Create a new Excel workbook with sales data
Copilot: I'll create a new workbook and add sales data to it.
[Calls excel_create_workbook]
[Calls excel_write_cell]
[Calls excel_save_workbook]
Copilot: I've created report.xlsx with the sales data. The file is saved at /path/to/report.xlsx.
```

**Data Analysis:**
```
You: Analyze the data in my Excel file and find duplicates
Copilot: Let me read the data and check for duplicates.
[Calls excel_open_workbook]
[Calls excel_read_range]
[Calls excel_find_duplicates]
Copilot: I found 5 duplicate entries in the data. Here are the details:
...
```

## Troubleshooting

### Server Not Found

**Possible causes:**
- Incorrect path to `dist/index.js`
- Server not built
- Configuration file not in correct location
- VS Code not restarted after configuration

**Solutions:**

1. Rebuild the project:
```bash
cd "/path/to/excel-mcp"
npm run build
```

2. Verify the file exists:
```bash
test -f "/path/to/excel-mcp/dist/index.js" && echo "✅ File exists" || echo "❌ File not found"
```

3. Check the configuration file location:
   - macOS: `~/Library/Application Support/Copilot/user/settings.json`
   - Windows: `%APPDATA%/Copilot/user/settings.json`
   - Linux: `~/.config/Copilot/user/settings.json`

4. Restart VS Code completely (not just reload)

### MCP Server Not Appearing in Copilot Chat

**Possible causes:**
- Configuration syntax error
- Server disabled
- Copilot Chat extension not installed

**Solutions:**

1. Verify JSON syntax:
```bash
# macOS/Linux
cat ~/Library/Application\ Support/Copilot/user/settings.json | jq .

# Windows (PowerShell)
cat $env:APPDATA\Copilot\user\settings.json | ConvertFrom-Json
```

2. Ensure GitHub Copilot Chat extension is installed:
   - Open Extensions panel (`Ctrl+Shift+X`)
   - Search for "GitHub Copilot Chat"
   - Install if not present

3. Check Copilot Chat output:
   - Open Output panel (`Ctrl+Shift+U` or `Cmd+Shift+U`)
   - Select "GitHub Copilot Chat" from dropdown
   - Look for MCP server initialization messages

### Access Denied Errors

**Check configuration:**

1. Open the settings file:
```bash
# macOS
open ~/Library/Application\ Support/Copilot/user/settings.json

# Windows
notepad $env:APPDATA\Copilot\user\settings.json

# Linux
nano ~/.config/Copilot/user/settings.json
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
- Workbook not opened before operations
- Incorrect file path
- File corrupted or doesn't exist

**Solutions:**

1. Always ask Copilot to open the file first:
```
Please open the Excel file at /full/path/to/file.xlsx first, then read the data
```

2. Verify file exists:
```bash
test -f "/path/to/file.xlsx" && echo "✅ File exists" || echo "❌ File not found"
```

3. Use health check:
```
Please check if the Excel MCP server is working properly
```

### Copilot Not Using Excel Tools

**Possible causes:**
- MCP server not initialized
- Copilot doesn't recognize the request needs Excel tools
- Server not properly configured

**Solutions:**

1. Be explicit in your request:
```
Please use the Excel MCP tools to open and read the file at /path/to/file.xlsx
```

2. Verify server is listed in Copilot Chat:
   - Open Copilot Chat
   - Look for MCP servers in the available tools
   - Excel tools should be visible

3. Check output logs for errors:
   - Open Output panel
   - Select "GitHub Copilot Chat"
   - Look for initialization or error messages

### Performance Issues with Large Files

**Possible causes:**
- File too large
- System resources limited

**Solutions:**

1. Increase file size limit:
```json
{
  "env": {
    "MCP_MAX_FILE_SIZE": "209715200"
  }
}
```

2. Work with smaller ranges:
```
Read data in smaller chunks, like A1:C100 first, then C101:C200
```

3. Enable debug logs to identify bottlenecks:
```json
{
  "env": {
    "MCP_LOG_LEVEL": "debug"
  }
}
```

## Environment Variables Reference

| Variable | Default | Description | Example |
|----------|---------|-------------|---------|
| `MCP_LOG_LEVEL` | info | Log level: debug, info, warn, error | `"debug"` |
| `MCP_ALLOWED_PATHS` | All paths | Comma-separated allowed paths | `"/path/*,/other/*"` |
| `MCP_DENIED_PATHS` | System paths | Comma-separated denied paths | `"/etc/*,/tmp/*"` |
| `MCP_MAX_FILE_SIZE` | 52428800 | Max file size in bytes (default 50MB) | `"104857600"` (100MB) |

## Verification

To verify the connection is working:

1. Open GitHub Copilot Chat in VS Code
2. Ask:
```
Please check if the Excel MCP server is running properly
```

3. Copilot should call the health check tool and return server status

4. If successful, you should see server information in the response

## Best Practices

1. **Be explicit**: Clearly state you want to use Excel MCP tools
2. **Use full paths**: Always provide complete file paths
3. **Open first**: Always ask to open workbooks before other operations
4. **Save regularly**: Explicitly ask to save workbooks after changes
5. **Close workbooks**: Ask to close workbooks when done
6. **Work in chunks**: For large files, process data in smaller ranges
7. **Handle errors**: Ask Copilot to check for errors after operations
8. **Provide context**: Give Copilot enough context about what you want to achieve

## Comparison with Cline

| Feature | GitHub Copilot | Cline |
|---------|----------------|--------|
| Configuration Location | `~/Library/Application Support/Copilot/user/settings.json` | `~/Library/Application Support/Code/User/globalStorage/saoudrizwan.claude-dev/settings/cline_mcp_settings.json` |
| Auto-approval | Not available | Available via `autoApprove` parameter |
| Timeout Control | Managed by Copilot | Configurable via `timeout` parameter |
| Tool Invocation | Automatic based on context | Manual tool selection or AI-driven |
| Chat Interface | Copilot Chat panel | Cline sidebar |

## Next Steps

- Read [API Documentation](./API.md) for complete tool reference
- Review [Security Guide](./SECURITY.md) for security best practices
- Check [Troubleshooting](./TROUBLESHOOTING.md) for common issues
- Explore [Architecture](./ARCHITECTURE.md) to understand system design

## Support

If you encounter issues:
1. Check VS Code Output panel for error messages
2. Review GitHub Copilot Chat logs
3. Verify the configuration file is correct
4. Check this troubleshooting section
5. Review [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) for detailed solutions

## Additional Resources

- [GitHub Copilot Documentation](https://docs.github.com/copilot)
- [Model Context Protocol Specification](https://modelcontextprotocol.io/)
- [Excel MCP Repository](https://github.com/az-coder-123/excel-mcp)