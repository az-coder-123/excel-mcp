# Excel MCP Server — Antigravity Setup Guide

## Overview

This guide provides step-by-step instructions for configuring and using the **Excel MCP Server** within **Google Antigravity** (including Antigravity IDE, `agy` CLI, and Antigravity 2.0).

With this integration, Antigravity AI agents can directly create, read, edit, format, and mathematically analyze Excel workbooks (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`) through native MCP tool calls.

---

## Prerequisites

- **Google Antigravity IDE** or **Antigravity CLI (`agy`)** installed.
- **Node.js**: v18.0.0 or higher (`node -v`).
- **Excel MCP Server**: Repository cloned and compiled with `npm run build`.

---

## Quick Reference

### 1. Build the Server
```bash
cd "/path/to/excel-mcp"
npm install
npm run build
```

### 2. Global Configuration File Location
```bash
# macOS / Linux
~/.gemini/config/mcp_config.json

# Windows
%USERPROFILE%\.gemini\config\mcp_config.json
```

### 3. Workspace Configuration File Location (Optional, Per-Project)
```bash
<project-root>/.agents/mcp_config.json
```

### 4. Health Check
In Antigravity chat, ask the assistant:
> "Run `excel_health_check` to verify the Excel MCP server connection."

---

## Configuration Methods

Antigravity supports two scopes for MCP server configuration:

### Method 1: Global Configuration (Recommended)

Applies across all projects and workspaces opened in Antigravity.

1. Open or create `~/.gemini/config/mcp_config.json`:
   ```bash
   nano ~/.gemini/config/mcp_config.json
   ```
2. Add the `excel` server definition under `mcpServers`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": [
        "/Users/trannamlong/PROJECT/excel-mcp/dist/index.js"
      ],
      "env": {
        "MCP_ALLOWED_PATHS": "/Users/trannamlong/PROJECT",
        "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*,C:\\Windows\\*",
        "MCP_MAX_FILE_SIZE": "52428800",
        "MCP_LOG_LEVEL": "info"
      }
    }
  }
}
```

> **Note**: Update the path in `args` and `MCP_ALLOWED_PATHS` to match your actual environment.

---

### Method 2: Project-Level Workspace Configuration

If you want to configure the server only for a specific repository or share settings with your team:

1. Create `.agents/mcp_config.json` inside your project root:
   ```bash
   mkdir -p .agents
   touch .agents/mcp_config.json
   ```
2. Add the configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": [
        "./node_modules/excel-mcp/dist/index.js"
      ],
      "env": {
        "MCP_ALLOWED_PATHS": ".",
        "MCP_LOG_LEVEL": "info"
      }
    }
  }
}
```

---

## Environment Variables & Security Options

| Environment Variable | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `MCP_ALLOWED_PATHS` | `string` (comma-separated) | `process.cwd()` | Sandbox whitelist. The server will reject access to any file outside these directory trees. |
| `MCP_DENIED_PATHS` | `string` (comma-separated) | OS system dirs | Blacklist paths (takes precedence over allowed paths). |
| `MCP_MAX_FILE_SIZE` | `number` (bytes) | `52428800` (50MB) | Maximum file size allowed for processing. |
| `MCP_LOG_LEVEL` | `string` | `"info"` | Logging verbosity: `"debug"`, `"info"`, `"warn"`, `"error"`, `"silent"`. |
| `MCP_SERVER_NAME` | `string` | `"excel-mcp"` | Identifier name of the MCP server. |

---

## Activating the Server in Antigravity

1. Save `~/.gemini/config/mcp_config.json`.
2. **Reload Antigravity IDE**:
   - Press `Cmd + Shift + P` (macOS) or `Ctrl + Shift + P` (Windows/Linux).
   - Select **Developer: Reload Window**, or simply start a new conversation session.
3. The Antigravity runtime starts the Node.js server via `stdio` transport and registers all **127 Excel tools**.

---

## Verifying in Antigravity Chat

Prompt Antigravity with any of the following natural commands:

### 1. Health & Status Check
> "Kiểm tra kết nối và trạng thái của Excel MCP server."
- The agent calls `excel_health_check` and reports memory, active workbooks, allowed paths, and tool readiness.

### 2. Creating & Formatting Workbooks
> "Tạo file Excel mới tại `/Users/trannamlong/PROJECT/test.xlsx`, thêm bảng dữ liệu gồm cột 'Tháng', 'Doanh thu', 'Chi phí' và tô màu tiêu đề đậm."
- The agent automatically calls:
  1. `excel_create_workbook`
  2. `excel_write_batch`
  3. `excel_apply_header_style`
  4. `excel_save_workbook`

### 3. Financial & Accounting Calculations
> "Tính giá trị hiện tại thuần (NPV) với lãi suất chiết khấu 10% cho dòng tiền tại Sheet1!B2:B10 trong file test.xlsx."
- The agent calls `excel_calculate_npv` using exact Excel-compatible semantics.

### 4. Currency & Number Formatting
> "Định dạng cột B thành tiền tệ Việt Nam Đồng (VND) và tô màu đỏ các số âm."
- The agent calls `excel_vnd_currency_format` and `excel_negative_red_format`.

---

## Troubleshooting

### 1. Tools Not Showing Up in Antigravity
* **Check Node.js Path**: If you use `nvm` or custom version managers, GUI apps on macOS may not have `node` in default PATH.
  Find your absolute node path:
  ```bash
  which node
  ```
  And specify the absolute path in `mcp_config.json`:
  ```json
  "command": "/Users/trannamlong/.nvm/versions/node/v22.18.0/bin/node"
  ```
* **Verify Build Output**: Ensure `dist/index.js` exists:
  ```bash
  ls -la /Users/trannamlong/PROJECT/excel-mcp/dist/index.js
  ```
  If missing, run `npm run build`.

### 2. "Access denied: path is outside allowed paths"
* The file you are trying to open is not in `MCP_ALLOWED_PATHS`.
* Update `MCP_ALLOWED_PATHS` in `mcp_config.json` to include the parent folder of your Excel files, e.g.:
  ```json
  "MCP_ALLOWED_PATHS": "/Users/trannamlong/PROJECT,/Users/trannamlong/Documents"
  ```

### 3. Test Server Locally via Command Line
You can test the server stdio output directly in terminal:
```bash
node -e "
import('./dist/tools/tool-definitions.js').then(m => {
  console.log('✅ Excel MCP loaded successfully! Total tools:', m.TOOL_DEFINITIONS.length);
});
"
```

---

## Related Documentation

- [`SETUP_GUIDE.md`](./SETUP_GUIDE.md) — General installation and tool catalog
- [`CLINE_SETUP.md`](./CLINE_SETUP.md) — Configuration guide for Cline extension
- [`GITHUB_COPILOT_SETUP.md`](./GITHUB_COPILOT_SETUP.md) — Configuration guide for GitHub Copilot
- [`SECURITY.md`](./SECURITY.md) — Security model, path sandboxing, and permissions
