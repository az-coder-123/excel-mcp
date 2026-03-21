/**
 * Permission Checker - Security module for access control
 * Single Responsibility: Handles all permission validation logic
 */

import { Permission, PermissionConfig, OperationResult } from '../types/index.js';

export class PermissionChecker {
  private config: PermissionConfig;

  constructor(config: PermissionConfig) {
    this.config = config;
  }

  /**
   * Check if a path is allowed based on configuration
   */
  public isPathAllowed(filePath: string): OperationResult<boolean> {
    const normalizedPath = this.normalizePath(filePath);

    // Check denied paths first (higher priority)
    for (const deniedPath of this.config.deniedPaths) {
      if (this.matchesPath(normalizedPath, deniedPath)) {
        return {
          success: false,
          error: `Access denied: Path "${filePath}" matches denied pattern "${deniedPath}"`,
        };
      }
    }

    // Check allowed paths
    if (this.config.allowedPaths.length === 0) {
      // If no allowed paths specified, allow all (except denied)
      return { success: true, data: true };
    }

    for (const allowedPath of this.config.allowedPaths) {
      if (this.matchesPath(normalizedPath, allowedPath)) {
        return { success: true, data: true };
      }
    }

    return {
      success: false,
      error: `Access denied: Path "${filePath}" does not match any allowed patterns`,
    };
  }

  /**
   * Check if file extension is allowed
   */
  public isExtensionAllowed(filePath: string): OperationResult<boolean> {
    const extension = this.getExtension(filePath);

    if (this.config.allowedExtensions.length === 0) {
      return { success: true, data: true };
    }

    const isAllowed = this.config.allowedExtensions.some(
      (ext) => ext.toLowerCase() === extension.toLowerCase()
    );

    if (!isAllowed) {
      return {
        success: false,
        error: `File extension "${extension}" is not allowed. Allowed: ${this.config.allowedExtensions.join(', ')}`,
      };
    }

    return { success: true, data: true };
  }

  /**
   * Check if file size is within limits
   */
  public isFileSizeAllowed(sizeInBytes: number): OperationResult<boolean> {
    if (sizeInBytes > this.config.maxFileSize) {
      return {
        success: false,
        error: `File size ${this.formatBytes(sizeInBytes)} exceeds maximum allowed size ${this.formatBytes(this.config.maxFileSize)}`,
      };
    }

    return { success: true, data: true };
  }

  /**
   * Check if user has required permission
   */
  public hasPermission(required: Permission): OperationResult<boolean> {
    if (this.config.permissions.includes('admin')) {
      return { success: true, data: true };
    }

    if (!this.config.permissions.includes(required)) {
      return {
        success: false,
        error: `Permission denied: Required "${required}" permission`,
      };
    }

    return { success: true, data: true };
  }

  /**
   * Validate all security checks for a file operation
   */
  public validateFileAccess(
    filePath: string,
    fileSize: number,
    requiredPermission: Permission
  ): OperationResult<boolean> {
    // Check permission
    const permissionResult = this.hasPermission(requiredPermission);
    if (!permissionResult.success) {
      return permissionResult;
    }

    // Check path
    const pathResult = this.isPathAllowed(filePath);
    if (!pathResult.success) {
      return pathResult;
    }

    // Check extension
    const extensionResult = this.isExtensionAllowed(filePath);
    if (!extensionResult.success) {
      return extensionResult;
    }

    // Check file size (if provided)
    if (fileSize > 0) {
      const sizeResult = this.isFileSizeAllowed(fileSize);
      if (!sizeResult.success) {
        return sizeResult;
      }
    }

    return { success: true, data: true };
  }

  /**
   * Update permission configuration
   */
  public updateConfig(config: Partial<PermissionConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * Get current configuration
   */
  public getConfig(): PermissionConfig {
    return { ...this.config };
  }

  // Private helper methods
  private normalizePath(filePath: string): string {
    return filePath.replace(/\\/g, '/').toLowerCase();
  }

  private matchesPath(filePath: string, pattern: string): boolean {
    const normalizedPattern = pattern.replace(/\\/g, '/').toLowerCase();

    // Simple wildcard matching
    if (normalizedPattern.includes('*')) {
      const regexPattern = normalizedPattern
        .replace(/\*/g, '.*')
        .replace(/\?/g, '.');
      const regex = new RegExp(`^${regexPattern}$`);
      return regex.test(filePath);
    }

    // Exact match or prefix match
    return filePath === normalizedPattern || 
           filePath.startsWith(normalizedPattern + '/');
  }

  private getExtension(filePath: string): string {
    const lastDot = filePath.lastIndexOf('.');
    return lastDot === -1 ? '' : filePath.slice(lastDot);
  }

  private formatBytes(bytes: number): string {
    const units = ['B', 'KB', 'MB', 'GB'];
    let unitIndex = 0;
    let size = bytes;

    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex++;
    }

    return `${size.toFixed(2)} ${units[unitIndex]}`;
  }
}