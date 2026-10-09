# Security Guide

## Overview

The Excel MCP Server implements a comprehensive security model to protect your system and data. This guide explains the security features and how to configure them.

## Security Architecture

### Defense in Depth

The server uses multiple layers of security:

1. **Permission Level**: Role-based access control
2. **Path Restrictions**: Filesystem access control
3. **File Type Whitelisting**: Extension filtering
4. **File Size Limits**: Resource protection
5. **Input Validation**: Parameter validation

### Security Check Flow

```
Request
  ↓
Permission Check (read/write/delete/admin)
  ↓
Path Validation (allowed/denied patterns)
  ↓
Extension Check (.xlsx, .xls, etc.)
  ↓
Size Validation (< maxFileSize)
  ↓
Operation Execution
```

## Permission System

### Permission Levels

| Permission | Description | Operations |
|------------|-------------|------------|
| `read` | Read-only access | Open, read cells, list worksheets |
| `write` | Write access | Create, write cells, add worksheets |
| `delete` | Delete access | Remove worksheets, clear cells |
| `admin` | Full access | All operations including config changes |

### Default Configuration

By default, the server grants:
- `read`
- `write`
- `delete`

To restrict permissions, set `MCP_PERMISSIONS_CONFIG`:

```json
{
  "permissions": ["read"]
}
```

### Permission Hierarchy

```
admin > delete > write > read
```

- `admin` includes all other permissions
- `delete` includes `write` and `read`
- `write` includes `read`

## Path Restrictions

### Allowed Paths

Specify directories where files can be accessed:

```env
MCP_ALLOWED_PATHS=/workspace,/home/user/documents
```

**Behavior**:
- Empty list = all paths allowed (except denied)
- Non-empty = only listed paths allowed

### Denied Paths

Specify paths that should never be accessed:

```env
MCP_DENIED_PATHS=/etc/*,/sys/*,/proc/*,C:\Windows\*
```

**Behavior**:
- Always checked first (higher priority than allowed)
- Supports wildcards: `*` and `?`

### Pattern Matching

| Pattern | Matches |
|---------|---------|
| `/etc/*` | `/etc/hosts`, `/etc/passwd`, etc. |
| `/workspace/*.xlsx` | All Excel files in `/workspace` |
| `/home/*/documents` | `/home/user/documents`, `/home/admin/documents` |

### Recommended Configuration

```env
# Allow only workspace directories
MCP_ALLOWED_PATHS=/workspace,/home/*/workspace

# Block system directories
MCP_DENIED_PATHS=/etc/*,/sys/*,/proc/*,/dev/*,C:\Windows\*,C:\Program Files\*
```

## File Type Restrictions

### Allowed Extensions

Only process specific file types:

```env
# Default extensions
.xlsx    # Excel Workbook
.xls     # Excel 97-2003
.xlsm    # Macro-Enabled Workbook
.xlsb    # Binary Workbook
```

Configuration:

```json
{
  "allowedExtensions": [".xlsx", ".xls"]
}
```

## File Size Limits

### Maximum File Size

Prevent resource exhaustion:

```env
MCP_MAX_FILE_SIZE=52428800  # 50MB (default)
```

**Recommended Sizes**:
- Small workload: 10MB (10485760)
- Medium workload: 50MB (52428800)
- Large workload: 100MB (104857600)

## Security Best Practices

### 1. Principle of Least Privilege

Grant only required permissions:

```json
{
  "permissions": ["read"]  // Read-only for untrusted clients
}
```

### 2. Restrict Filesystem Access

Limit to specific directories:

```env
MCP_ALLOWED_PATHS=/workspace/data
MCP_DENIED_PATHS=/
```

### 3. Use Absolute Paths

Always use absolute paths to avoid ambiguity:

```json
{
  "filePath": "/workspace/report.xlsx"  // Good
}
```

```json
{
  "filePath": "../report.xlsx"  // Bad - path traversal risk
}
```

### 4. Validate File Sources

Only open files from trusted sources:

```javascript
// Check file source before opening
if (isTrustedSource(filePath)) {
  await excel_open_workbook({ filePath });
}
```

### 5. Close Workbooks When Done

Free resources and prevent memory leaks:

```javascript
await excel_open_workbook({ filePath: "/data/report.xlsx" });
// ... operations ...
await excel_close_workbook({ filename: "report.xlsx" });
```

### 6. Monitor Logs

Review logs for suspicious activity:

```bash
# View error logs
grep "ERROR" /var/log/excel-mcp.log

# View permission denials
grep "Permission denied" /var/log/excel-mcp.log
```

## Configuration Examples

### Read-Only Mode

```env
MCP_PERMISSIONS_CONFIG={"permissions":["read"],"allowedPaths":["/workspace"],"deniedPaths":["/*"]}
```

### Restricted Workspace

```env
MCP_ALLOWED_PATHS=/workspace/projects
MCP_DENIED_PATHS=/etc/*,/sys/*,/proc/*,/home/*/.ssh/*
MCP_MAX_FILE_SIZE=10485760
```

### Production Configuration

```env
MCP_SERVER_NAME=excel-mcp-prod
MCP_LOG_LEVEL=warn
MCP_ALLOWED_PATHS=/data/excel
MCP_DENIED_PATHS=/etc/*,/sys/*,/proc/*,/dev/*,/boot/*,C:\Windows\*,C:\Program Files\*
MCP_MAX_FILE_SIZE=52428800
MCP_PERMISSIONS_CONFIG={"allowedPaths":["/data/excel"],"deniedPaths":[],"maxFileSize":52428800,"allowedExtensions":[".xlsx"],"permissions":["read","write"]}
```

### Development Configuration

```env
MCP_SERVER_NAME=excel-mcp-dev
MCP_LOG_LEVEL=debug
MCP_ALLOWED_PATHS=
MCP_DENIED_PATHS=/etc/*,/sys/*
MCP_MAX_FILE_SIZE=104857600
```

## Threat Model

### Mitigated Threats

| Threat | Mitigation |
|--------|------------|
| Path Traversal | Path normalization and validation |
| Unauthorized Access | Permission checking |
| Resource Exhaustion | File size limits |
| Malicious Files | Extension whitelisting |
| Data Exfiltration | Path restrictions |
| System File Access | Denied path patterns |

### Known Limitations

1. **Macro Execution**: Excel macros in `.xlsm` files are not executed
2. **External Links**: External data links are not followed
3. **Encryption**: Password-protected files require manual decryption
4. **Network Access**: No network operations are performed

## Incident Response

### Permission Denied Events

When a permission is denied:

1. Check server logs for details
2. Verify configuration is correct
3. Adjust permissions if legitimate
4. Block source if suspicious

### Suspicious Activity

If suspicious activity is detected:

1. Review recent logs
2. Check file access patterns
3. Verify configuration hasn't been tampered with
4. Consider restricting permissions further

## Audit Logging

### Log Format

```
[timestamp] LEVEL message {context}
```

Example:

```
[2024-01-15T10:30:00.000Z] INFO  Executing tool: excel_read_cell {filename: "report.xlsx"}
[2024-01-15T10:30:01.000Z] ERROR Permission denied: Path "/etc/passwd" matches denied pattern "/etc/*"
```

### Log Levels

| Level | Use Case |
|-------|----------|
| `debug` | Development, troubleshooting |
| `info` | Normal operations |
| `warn` | Potential issues |
| `error` | Failed operations, security violations |

### Log Storage

Logs are output to `stderr` for MCP compatibility. To persist logs:

```bash
node dist/index.js 2> /var/log/excel-mcp.log
```

Or use process manager logging:

```bash
pm2 start dist/index.js --name excel-mcp
pm2 logs excel-mcp
```

## Security Checklist

- [ ] Set appropriate permission levels
- [ ] Configure allowed paths
- [ ] Block sensitive system directories
- [ ] Set reasonable file size limits
- [ ] Enable appropriate logging level
- [ ] Review logs regularly
- [ ] Keep dependencies updated
- [ ] Use absolute paths in operations
- [ ] Close workbooks when done
- [ ] Validate file sources

## Reporting Security Issues

If you discover a security vulnerability:

1. Do not open a public issue
2. Email security concerns privately
3. Include steps to reproduce
4. Allow time for patch before disclosure

## Compliance

### Data Protection

- No data is sent to external services
- All operations are local
- Logs can be configured to avoid sensitive data

### Access Control

- Role-based permission system
- Path-based restrictions
- Audit logging for compliance

### Resource Protection

- File size limits prevent DoS
- Concurrent operation limits
- Memory management via workbook cleanup