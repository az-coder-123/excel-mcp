# Troubleshooting Guide

## Common Issues

### Installation Issues

#### "npm install" fails

**Symptom**: Errors during dependency installation

**Solutions**:

1. Check Node.js version:
```bash
node --version
# Should be >= 18.0.0
```

2. Clear npm cache:
```bash
npm cache clean --force
rm -rf node_modules package-lock.json
npm install
```

3. Check network connectivity:
```bash
npm ping
```

#### TypeScript compilation errors

**Symptom**: `npm run build` fails with type errors

**Solutions**:

1. Ensure dependencies are installed:
```bash
npm install
```

2. Check tsconfig.json exists and is valid

3. Try cleaning and rebuilding:
```bash
npm run clean
npm run build
```

### Permission Errors

#### "Permission denied" when opening files

**Symptom**: Error message like `Access denied: Path "/path/to/file.xlsx" matches denied pattern`

**Causes & Solutions**:

1. **File in denied path**:
   - Check `MCP_DENIED_PATHS` configuration
   - Move file to allowed directory
   - Adjust denied paths if needed

2. **File not in allowed path** (when `MCP_ALLOWED_PATHS` is set):
   - Add parent directory to `MCP_ALLOWED_PATHS`
   - Or move file to allowed directory

3. **Insufficient permission level**:
   - Check tool's `requiredPermissions`
   - Ensure `MCP_PERMISSIONS_CONFIG` includes needed permissions

**Example Fix**:

```env
# Before (restrictive)
MCP_ALLOWED_PATHS=/workspace/safe
MCP_DENIED_PATHS=/etc/*,/sys/*

# After (allows specific directory)
MCP_ALLOWED_PATHS=/workspace/safe,/data/reports
MCP_DENIED_PATHS=/etc/*,/sys/*
```

#### "Permission denied: Required 'write' permission"

**Symptom**: Write operations fail

**Solution**: Add `write` to permissions:

```env
MCP_PERMISSIONS_CONFIG={"permissions":["read","write"]}
```

### Workbook Errors

#### "Workbook not opened"

**Symptom**: `excel_read_cell` or similar fails with "Workbook not opened"

**Causes & Solutions**:

1. **Forgot to call `excel_open_workbook` first**:
   ```javascript
   // Wrong
   await excel_read_cell({ filename: "report.xlsx", ... });
   
   // Correct
   await excel_open_workbook({ filePath: "/path/to/report.xlsx" });
   await excel_read_cell({ filename: "report.xlsx", ... });
   ```

2. **Workbook was closed**:
   - Workbooks are tracked by filename
   - Must reopen if previously closed

3. **Filename mismatch**:
   - Use exact filename from `excel_open_workbook` response
   - Case-sensitive on some systems

#### "Worksheet not found"

**Symptom**: Operations fail with worksheet name errors

**Solutions**:

1. **List worksheets first**:
   ```javascript
   const result = await excel_list_worksheets({ filename: "report.xlsx" });
   console.log(result); // See available worksheet names
   ```

2. **Check worksheet name spelling**:
   - Worksheet names are case-sensitive
   - Check for extra spaces

3. **Worksheet might be hidden**:
   - Hidden worksheets are still accessible
   - Check `hidden` property in worksheet info

### File System Errors

#### "File not found"

**Symptom**: `excel_open_workbook` fails with file not found

**Solutions**:

1. **Verify file exists**:
   ```bash
   ls -la /path/to/file.xlsx
   ```

2. **Use absolute path**:
   ```javascript
   // Bad (relative path)
   await excel_open_workbook({ filePath: "report.xlsx" });
   
   // Good (absolute path)
   await excel_open_workbook({ filePath: "/workspace/report.xlsx" });
   ```

3. **Check file permissions**:
   ```bash
   ls -la /path/to/file.xlsx
   # Ensure read permission
   ```

#### "File exceeds maximum allowed size"

**Symptom**: Large files cannot be opened

**Solutions**:

1. **Increase file size limit**:
   ```env
   MCP_MAX_FILE_SIZE=104857600  # 100MB
   ```

2. **Split large files**:
   - Use multiple smaller workbooks
   - Or extract data into database

### Integration Issues

#### Server not appearing in GitHub Copilot

**Symptom**: Excel tools don't show in Copilot chat

**Solutions**:

1. **Check VS Code settings.json**:
   ```json
   {
     "github.copilot.chat.experimental.mcpServers": {
       "excel-mcp": {
         "command": "node",
         "args": ["/absolute/path/to/excel-mcp/dist/index.js"]
       }
     }
   }
   ```

2. **Ensure server is built**:
   ```bash
   npm run build
   ls dist/index.js  # Should exist
   ```

3. **Restart VS Code** after configuration changes

4. **Check path is absolute** (not relative)

#### Server not appearing in Cline

**Symptom**: Excel tools don't show in Cline

**Solutions**:

1. **Check MCP configuration file location**:
   - macOS/Linux: `~/.config/cline/mcp_servers.json`
   - Windows: `%APPDATA%\cline\mcp_servers.json`

2. **Verify JSON syntax**:
   ```json
   {
     "mcpServers": {
       "excel-mcp": {
         "command": "node",
         "args": ["/absolute/path/to/excel-mcp/dist/index.js"]
       }
     }
   }
   ```

3. **Restart Cline/VS Code**

#### Tools execute but return errors

**Symptom**: Tools run but fail with errors

**Solutions**:

1. **Check server logs**:
   ```bash
   # If running manually
   node dist/index.js 2>&1 | tee server.log
   ```

2. **Enable debug logging**:
   ```env
   MCP_LOG_LEVEL=debug
   ```

3. **Test with simple operation first**:
   ```javascript
   // Try creating a new workbook
   await excel_create_workbook({
     filename: "test",
     outputPath: "/tmp"
   });
   ```

### Memory Issues

#### Server becomes slow or crashes

**Symptom**: Performance degrades over time

**Solutions**:

1. **Close workbooks when done**:
   ```javascript
   await excel_open_workbook({ filePath: "/data/report.xlsx" });
   // ... operations ...
   await excel_close_workbook({ filename: "report.xlsx" });
   ```

2. **Limit concurrent workbooks**:
   ```env
   # Default is 10, reduce if needed
   MCP_MAX_CONCURRENT_OPERATIONS=5
   ```

3. **Restart server periodically** in long-running scenarios

### Formula Issues

#### Formulas not calculating

**Symptom**: Formula cells show `0` or error

**Solutions**:

1. **Formulas are not evaluated by ExcelJS**:
   - Formula result is stored in file
   - If file wasn't saved with Excel, result may be missing

2. **Read formula and value separately**:
   ```javascript
   const cell = await excel_get_cell_info({
     filename: "report.xlsx",
     worksheet: "Sheet1",
     cellAddress: "B10"
   });
   // cell.formula contains "=SUM(B1:B9)"
   // cell.value contains computed result
   ```

3. **Use Excel to recalculate**:
   - Open file in Excel
   - Save it
   - Formulas will be recalculated

### Encoding Issues

#### Special characters display incorrectly

**Symptom**: Unicode or special characters are garbled

**Solutions**:

1. **Ensure file is UTF-8 encoded**:
   - Excel usually handles this automatically
   - Re-save file with Excel if needed

2. **Check cell type**:
   ```javascript
   const cell = await excel_read_cell({ ... });
   if (cell.type === 'string') {
     console.log(cell.value);
   }
   ```

## Error Messages Reference

| Error Message | Cause | Solution |
|---------------|-------|----------|
| `Workbook "X" not opened` | Operation before open | Call `excel_open_workbook` first |
| `Worksheet "X" not found` | Wrong worksheet name | Use `excel_list_worksheets` to check names |
| `Permission denied: Path "X"` | Path blocked | Adjust `MCP_ALLOWED_PATHS` or `MCP_DENIED_PATHS` |
| `Permission denied: Required "X"` | Insufficient permission | Add permission to config |
| `File not found` | Invalid path | Check file exists and path is absolute |
| `File exceeds maximum size` | File too large | Increase `MCP_MAX_FILE_SIZE` |
| `Unknown tool: X` | Tool doesn't exist | Check tool name spelling |
| `Missing required parameter` | Missing argument | Check tool's required parameters |

## Debugging Steps

### 1. Enable Debug Logging

```env
MCP_LOG_LEVEL=debug
```

### 2. Test Server Manually

```bash
# Build first
npm run build

# Run server
node dist/index.js

# In another terminal, test with MCP client
```

### 3. Check Configuration

```bash
# View current environment
env | grep MCP_

# Or check .env file
cat .env
```

### 4. Verify File Permissions

```bash
# Check file is readable
cat /path/to/file.xlsx > /dev/null && echo "Readable"

# Check directory is accessible
ls -la /path/to/
```

### 5. Test with Simple Operation

```javascript
// Create a test file
const result = await excel_create_workbook({
  filename: "test",
  outputPath: "/tmp"
});
console.log(result);

// Read it back
await excel_open_workbook({ filePath: "/tmp/test.xlsx" });
const sheets = await excel_list_worksheets({ filename: "test.xlsx" });
console.log(sheets);
await excel_close_workbook({ filename: "test.xlsx" });
```

## Getting Help

### Before Asking for Help

1. Check this troubleshooting guide
2. Enable debug logging and capture logs
3. Try to reproduce with minimal example
4. Note your:
   - Node.js version
   - Operating system
   - Excel MCP version
   - Configuration (without sensitive data)

### Providing Information

When reporting issues, include:

```
**Environment**:
- Node.js version: 
- OS: 
- Excel MCP version: 

**Configuration**:
```
MCP_LOG_LEVEL=debug
# (other non-sensitive config)
```

**Steps to Reproduce**:
1. 
2. 
3. 

**Expected Behavior**:


**Actual Behavior**:


**Logs**:
```
(paste relevant logs here)
```
```

## Frequently Asked Questions

### Q: Can I use relative paths?

A: Yes, but absolute paths are recommended. Relative paths are resolved from the current working directory, which may vary.

### Q: How many workbooks can I open simultaneously?

A: Default is 10 concurrent workbooks. Configure with `MCP_MAX_CONCURRENT_OPERATIONS`.

### Q: Does it support .csv files?

A: No, only Excel formats (.xlsx, .xls, .xlsm, .xlsb). Convert CSV to Excel first.

### Q: Can it open password-protected files?

A: No, password-protected files must be decrypted first.

### Q: Are Excel macros executed?

A: No, macros in .xlsm files are not executed for security reasons.

### Q: How do I update to a new version?

A: 
```bash
git pull
npm install
npm run build
# Restart your MCP client