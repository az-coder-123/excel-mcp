# Excel MCP Server - Hướng dẫn cài đặt và kết nối

## Tổng quan

Excel MCP Server là một Model Context Protocol server cho phép tương tác với file Excel thông qua các công cụ AI như Cline và GitHub Copilot.

## Đã cài đặt sẵn

✅ Server đã được build trong folder `dist/`
✅ Dependencies đã được cài đặt
✅ Server đang chạy trong background

## Cấu hình cho Cline

### 1. Tạo file cấu hình

File cấu hình đã được tạo tại `.cline/mcp_servers.json` với nội dung:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["/Volumes/Macintosh HD - Data/projects/excel-mcp/dist/index.js"],
      "env": {
        "MCP_LOG_LEVEL": "info",
        "MCP_ALLOWED_PATHS": "/Volumes/Macintosh HD - Data/projects/excel-mcp",
        "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*,/Volumes/Macintosh HD - Data/projects/excel-mcp/node_modules/*"
      }
    }
  }
}
```

### 2. Cấu hình đã được thiết lập

- **Server Path**: `/Volumes/Macintosh HD - Data/projects/excel-mcp/dist/index.js`
- **Log Level**: info
- **Allowed Paths**: Chỉ cho phép truy cập vào project directory
- **Denied Paths**: Chặn truy cập vào các thư mục hệ thống và node_modules

## Kiểm tra kết nối

Server đã được test và đang hoạt động:

```
✅ MCP Server is responding!
✅ Server initialized and listening on stdio
✅ Tool calls are being received
```

## Các công cụ (Tools) có sẵn

### 1. Health Check
Kiểm tra trạng thái server:
```json
{
  "name": "excel_health_check",
  "arguments": {}
}
```

### 2. Workbook Operations
- `excel_open_workbook` - Mở file Excel
- `excel_create_workbook` - Tạo workbook mới
- `excel_save_workbook` - Lưu workbook
- `excel_close_workbook` - Đóng workbook
- `excel_list_worksheets` - Liệt kê tất cả worksheets

### 3. Cell Operations
- `excel_read_cell` - Đọc giá trị một ô
- `excel_write_cell` - Ghi giá trị vào một ô
- `excel_read_range` - Đọc nhiều ô
- `excel_write_range` - Ghi nhiều ô
- `excel_get_cell_info` - Lấy thông tin chi tiết ô

### 4. Worksheet Operations
- `excel_add_worksheet` - Thêm worksheet mới
- `excel_delete_worksheet` - Xóa worksheet
- `excel_rename_worksheet` - Đổi tên worksheet

### 5. Formatting Operations
- `excel_set_cell_format` - Định dạng ô
- `excel_set_cell_border` - Thêm border
- `excel_set_cell_fill` - Đổ màu ô
- `excel_set_cell_font` - Định dạng font

## Ví dụ sử dụng

### Đọc file Excel có sẵn

File mẫu: `employees_full_02.xlsx`

```json
{
  "name": "excel_open_workbook",
  "arguments": {
    "filePath": "/Volumes/Macintosh HD - Data/projects/excel-mcp/employees_full_02.xlsx"
  }
}
```

```json
{
  "name": "excel_list_worksheets",
  "arguments": {
    "filename": "employees_full_02.xlsx"
  }
}
```

```json
{
  "name": "excel_read_range",
  "arguments": {
    "filename": "employees_full_02.xlsx",
    "worksheet": "Sheet1",
    "startCell": "A1",
    "endCell": "C10"
  }
}
```

### Tạo và chỉnh sửa file Excel mới

```json
{
  "name": "excel_create_workbook",
  "arguments": {
    "filename": "report",
    "outputPath": "/Volumes/Macintosh HD - Data/projects/excel-mcp"
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

```json
{
  "name": "excel_save_workbook",
  "arguments": {
    "filename": "report.xlsx",
    "outputPath": "/Volumes/Macintosh HD - Data/projects/excel-mcp/report-final.xlsx"
  }
}
```

## Cấu hình nâng cao

### Thay đổi quyền truy cập

Sửa file `.cline/mcp_servers.json`:

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "node",
      "args": ["/Volumes/Macintosh HD - Data/projects/excel-mcp/dist/index.js"],
      "env": {
        "MCP_LOG_LEVEL": "debug",
        "MCP_ALLOWED_PATHS": "/Users/yourname/Documents,/workspace",
        "MCP_DENIED_PATHS": "/etc/*,/sys/*,/proc/*",
        "MCP_MAX_FILE_SIZE": "104857600"
      }
    }
  }
}
```

### Environment Variables

| Variable | Mặc định | Mô tả |
|----------|----------|-------|
| `MCP_LOG_LEVEL` | info | debug, info, warn, error |
| `MCP_ALLOWED_PATHS` | Tất cả | Danh sách đường dẫn được phép |
| `MCP_DENIED_PATHS` | System paths | Danh sách đường dẫn bị chặn |
| `MCP_MAX_FILE_SIZE` | 52428800 | Kích thước file tối đa (bytes) |

## Khắc phục sự cố

### Server không phản hồi

1. Kiểm tra server đang chạy:
```bash
ps aux | grep "node dist/index.js"
```

2. Khởi động lại server:
```bash
npm run start
```

### Lỗi quyền truy cập

1. Kiểm tra cấu hình trong `.cline/mcp_servers.json`
2. Đảm bảo file Excel nằm trong `MCP_ALLOWED_PATHS`
3. Đảm bảo file không nằm trong `MCP_DENIED_PATHS`

### Lỗi "Workbook not opened"

1. Đảm bảo đã gọi `excel_open_workbook` trước
2. Kiểm tra đường dẫn file chính xác
3. Xác định file không bị hỏng

## Tài liệu tham khảo

- [API Documentation](./API.md)
- [Architecture](./ARCHITECTURE.md)
- [Security Guide](./SECURITY.md)
- [Troubleshooting](./TROUBLESHOOTING.md)

## Hỗ trợ

Nếu gặp vấn đề, hãy kiểm tra:
1. Logs trong terminal khi server chạy
2. Console trong VS Code
3. File [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)

---

**Trạng thái hiện tại**: ✅ Đã kết nối và hoạt động bình thường