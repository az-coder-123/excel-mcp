/**
 * Core type definitions for Excel MCP Server
 * Following Single Responsibility Principle - each type has a clear purpose
 */

// Permission types for security
export type Permission = 'read' | 'write' | 'delete' | 'admin';

export interface PermissionConfig {
  allowedPaths: string[];
  deniedPaths: string[];
  maxFileSize: number;
  allowedExtensions: string[];
  permissions: Permission[];
}

// Excel operation types
export interface CellAddress {
  row: number;
  column: number;
}

export interface CellRange {
  start: CellAddress;
  end: CellAddress;
}

export interface CellValue {
  address: string;
  value: unknown;
  type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'error';
  formula?: string;
}

export interface WorksheetInfo {
  name: string;
  index: number;
  rowCount: number;
  columnCount: number;
  hidden: boolean;
}

export interface WorkbookInfo {
  filename: string;
  worksheetCount: number;
  worksheets: WorksheetInfo[];
  created?: Date;
  modified?: Date;
}

// Operation result types
export interface OperationResult<T = unknown> {
  success: boolean;
  data?: T;
  error?: string;
  metadata?: Record<string, unknown>;
}

// Tool definition types
export interface ToolParameter {
  name: string;
  type: 'string' | 'number' | 'boolean' | 'array' | 'object';
  description: string;
  required: boolean;
  enum?: string[];
  items?: ToolParameter;
  properties?: Record<string, ToolParameter>;
}

export interface ToolDefinition {
  name: string;
  description: string;
  parameters: ToolParameter[];
  requiredPermissions: Permission[];
}

// Configuration types
export interface ServerConfig {
  name: string;
  version: string;
  permissions: PermissionConfig;
  logLevel: 'debug' | 'info' | 'warn' | 'error';
  maxConcurrentOperations: number;
  operationTimeout: number;
}

// Log types
export interface LogEntry {
  timestamp: Date;
  level: 'debug' | 'info' | 'warn' | 'error';
  message: string;
  context?: Record<string, unknown>;
}

// Validation types
export interface ValidationRule {
  field: string;
  type: 'required' | 'type' | 'range' | 'pattern' | 'custom';
  value?: unknown;
  message: string;
  validator?: (value: unknown) => boolean;
}

export interface ValidationResult {
  valid: boolean;
  errors: string[];
}