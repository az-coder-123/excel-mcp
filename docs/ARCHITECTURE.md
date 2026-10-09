# Architecture Documentation

## Overview

The Excel MCP Server follows clean architecture principles with a focus on the Single Responsibility Principle (SRP). Each component has a clearly defined purpose, making the codebase maintainable, testable, and extensible.

## Design Principles

### 1. Single Responsibility Principle (SRP)

Every class and module has exactly one reason to change:

- **`PermissionChecker`**: Only handles access control validation
- **`ExcelService`**: Only manages Excel file operations
- **`Logger`**: Only handles logging output
- **`ToolHandler`**: Only executes tool operations
- **`ToolDefinitions`**: Only defines tool schemas

### 2. Dependency Inversion

High-level modules don't depend on low-level modules. Both depend on abstractions:

```
ExcelMCPServer (High-level)
    ├── depends on → Logger (abstraction)
    ├── depends on → PermissionChecker (abstraction)
    ├── depends on → ExcelService (abstraction)
    └── depends on → ToolHandler (abstraction)
```

### 3. Interface Segregation

Clients aren't forced to depend on interfaces they don't use. Each interface is focused:

- **`OperationResult<T>`**: Generic result wrapper
- **`PermissionConfig`**: Permission settings only
- **`ServerConfig`**: Server configuration only

### 4. Open/Closed Principle

The system is open for extension but closed for modification:

- New tools can be added to `TOOL_DEFINITIONS` without modifying existing code
- New permissions can be added without changing the permission checker logic

## Module Breakdown

### Entry Point (`src/index.ts`)

**Responsibility**: Application bootstrap and configuration loading

- Loads configuration from environment variables
- Creates server instance
- Handles graceful shutdown

### Type System (`src/types/index.ts`)

**Responsibility**: Type definitions for the entire application

- Permission types
- Excel operation types
- Configuration types
- Result types

### Security Module (`src/security/`)

#### `permission-checker.ts`

**Responsibility**: Access control validation

**Key Methods**:
- `isPathAllowed()`: Validates file path against whitelist/blacklist
- `isExtensionAllowed()`: Checks file extension
- `isFileSizeAllowed()`: Validates file size limits
- `hasPermission()`: Checks user permission level
- `validateFileAccess()`: Comprehensive validation

**Dependencies**: None (pure logic)

### Services (`src/services/`)

#### `excel-service.ts`

**Responsibility**: Excel file operations

**Key Methods**:
- `openWorkbook()`: Load Excel file
- `readCell()`: Read single cell value
- `readRange()`: Read multiple cells
- `writeCell()`: Write to a cell
- `saveWorkbook()`: Save to disk
- `createWorkbook()`: Create new file

**Dependencies**:
- `PermissionChecker` (injected)
- `Logger` (injected)

### Tools (`src/tools/`)

#### `tool-definitions.ts`

**Responsibility**: Tool schema definitions for MCP protocol

- Defines all available tools
- Specifies parameters and permissions
- Provides lookup functions

#### `tool-handler.ts`

**Responsibility**: Tool execution orchestration

- Validates tool existence
- Checks permissions
- Routes to appropriate service methods
- Handles errors

**Dependencies**:
- `ExcelService` (injected)
- `PermissionChecker` (injected)
- `Logger` (injected)

### Server (`src/server/`)

#### `excel-mcp-server.ts`

**Responsibility**: MCP protocol handling and server lifecycle

- Initializes all components
- Sets up MCP request handlers
- Manages server connection
- Handles graceful shutdown

**Dependencies**:
- All service modules (injected)

### Utilities (`src/utils/`)

#### `logger.ts`

**Responsibility**: Centralized logging

- Log level filtering
- Formatted output
- Log storage and retrieval

**Dependencies**: None

## Data Flow

### Tool Execution Flow

```
Client Request
    ↓
MCP Server (receives request)
    ↓
Tool Handler (validates and routes)
    ↓
Permission Checker (validates access)
    ↓
Excel Service (performs operation)
    ↓
Logger (logs operation)
    ↓
MCP Server (sends response)
    ↓
Client
```

### Detailed Flow Example: `excel_read_cell`

1. **Client** sends tool call request
2. **MCP Server** receives via `CallToolRequestSchema` handler
3. **Tool Handler** validates tool exists in `TOOL_DEFINITIONS`
4. **Permission Checker** validates `read` permission
5. **Excel Service** reads cell value from workbook
6. **Logger** records successful operation
7. **MCP Server** formats and sends response

## Security Architecture

### Layered Security Model

```
┌─────────────────────────────────────┐
│          Permission Layer           │
│  (permission-checker.ts)            │
├─────────────────────────────────────┤
│          Validation Layer           │
│  (tool-handler.ts)                  │
├─────────────────────────────────────┤
│          Service Layer              │
│  (excel-service.ts)                 │
├─────────────────────────────────────┤
│          File System                │
└─────────────────────────────────────┘
```

### Security Checks Order

1. **Permission Level**: Does user have required permission?
2. **Path Validation**: Is file path allowed?
3. **Extension Check**: Is file type permitted?
4. **Size Validation**: Is file size within limits?
5. **Operation**: Execute only if all checks pass

### Defense in Depth

- **Input Validation**: All inputs are validated before processing
- **Path Traversal Prevention**: Paths are normalized and checked
- **Permission Enforcement**: Every operation checks permissions
- **Error Handling**: Errors don't leak sensitive information
- **Logging**: All operations are logged for audit

## Configuration Architecture

### Configuration Sources (Priority Order)

1. Environment variables
2. Constructor parameters
3. Default values

### Configuration Categories

**Server Configuration**:
- Server name and version
- Log level
- Operation limits

**Permission Configuration**:
- Allowed paths
- Denied paths
- File size limits
- Allowed extensions
- Permission levels

## Error Handling Strategy

### Error Categories

1. **Validation Errors**: Invalid input or parameters
2. **Permission Errors**: Access denied
3. **File System Errors**: File not found, I/O errors
4. **Service Errors**: Excel operation failures
5. **Server Errors**: Unexpected errors

### Error Response Format

```typescript
{
  success: false,
  error: "Human-readable error message",
  metadata?: {
    // Additional context
  }
}
```

### Error Propagation

- Errors are caught at service boundaries
- Errors are logged with context
- Errors are converted to user-friendly messages
- Sensitive information is never exposed

## Extensibility

### Adding New Tools

1. Add tool definition to `TOOL_DEFINITIONS` in `tool-definitions.ts`
2. Add handler method in `tool-handler.ts`
3. Add service method if needed in `excel-service.ts`
4. No other files need modification

### Adding New Permissions

1. Add permission type to `Permission` union type in `types/index.ts`
2. Update `permission-checker.ts` if special logic needed
3. Add to tool definitions' `requiredPermissions`

### Adding New Configuration

1. Add to `ServerConfig` or `PermissionConfig` in `types/index.ts`
2. Update `loadConfig()` in `index.ts`
3. Update `mergeConfig()` in `excel-mcp-server.ts`
4. Update `.env.example`

## Testing Strategy

### Unit Testing

Each module can be tested in isolation:

```typescript
// Example: Testing PermissionChecker
const checker = new PermissionChecker(config);
const result = checker.isPathAllowed('/workspace/file.xlsx');
expect(result.success).toBe(true);
```

### Integration Testing

Test tool execution flow:

```typescript
// Example: Testing excel_read_cell tool
const result = await toolHandler.executeTool('excel_read_cell', {
  filename: 'test.xlsx',
  worksheet: 'Sheet1',
  cellAddress: 'A1'
});
expect(result.success).toBe(true);
```

### Security Testing

Verify permission enforcement:

```typescript
// Test: Permission denied
const checker = new PermissionChecker({
  permissions: ['read'], // No write permission
  // ...
});
const result = checker.hasPermission('write');
expect(result.success).toBe(false);
```

## Performance Considerations

### Memory Management

- Workbooks are stored in memory while open
- `closeWorkbook()` releases memory
- Maximum concurrent operations configurable

### I/O Optimization

- File reads are buffered
- Writes are batched where possible
- Large files can be streamed

### Logging Performance

- Log level filtering prevents unnecessary formatting
- stderr is used for logging (doesn't block stdio)
- Log rotation prevents unbounded growth

## Future Enhancements

### Planned Features

- **Streaming**: Support for streaming large files
- **Caching**: Workbook caching for repeated operations
- **Batch Operations**: Process multiple cells at once
- **Charts**: Read and create Excel charts
- **Formatting**: Cell formatting operations
- **Formulas**: Advanced formula support

### Extension Points

- Custom permission providers
- Plugin system for additional operations
- Webhook notifications for operations
- Audit logging to external systems