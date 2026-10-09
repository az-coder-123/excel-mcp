/**
 * Logger - Centralized logging utility
 * Single Responsibility: Handles all logging operations
 */

import { LogEntry } from '../types/index.js';

export class Logger {
  private logLevel: 'debug' | 'info' | 'warn' | 'error';
  private logs: LogEntry[] = [];
  private maxLogs: number = 1000;

  constructor(logLevel: 'debug' | 'info' | 'warn' | 'error' = 'info') {
    this.logLevel = logLevel;
  }

  /**
   * Log debug message
   */
  public debug(message: string, context?: Record<string, unknown>): void {
    this.log('debug', message, context);
  }

  /**
   * Log info message
   */
  public info(message: string, context?: Record<string, unknown>): void {
    this.log('info', message, context);
  }

  /**
   * Log warning message
   */
  public warn(message: string, context?: Record<string, unknown>): void {
    this.log('warn', message, context);
  }

  /**
   * Log error message
   */
  public error(message: string, context?: Record<string, unknown>): void {
    this.log('error', message, context);
  }

  /**
   * Core logging method
   */
  private log(
    level: 'debug' | 'info' | 'warn' | 'error',
    message: string,
    context?: Record<string, unknown>
  ): void {
    if (!this.shouldLog(level)) {
      return;
    }

    const entry: LogEntry = {
      timestamp: new Date(),
      level,
      message,
      context,
    };

    // Store log entry
    this.logs.push(entry);

    // Trim logs if exceeding max
    if (this.logs.length > this.maxLogs) {
      this.logs = this.logs.slice(-this.maxLogs);
    }

    // Output to stderr (MCP servers communicate via stdio)
    const logMessage = this.formatLog(entry);
    process.stderr.write(logMessage + '\n');
  }

  /**
   * Check if message should be logged based on level
   */
  private shouldLog(level: 'debug' | 'info' | 'warn' | 'error'): boolean {
    const levels = ['debug', 'info', 'warn', 'error'];
    const currentLevelIndex = levels.indexOf(this.logLevel);
    const messageLevelIndex = levels.indexOf(level);
    return messageLevelIndex >= currentLevelIndex;
  }

  /**
   * Format log entry for output
   */
  private formatLog(entry: LogEntry): string {
    const timestamp = entry.timestamp.toISOString();
    const level = entry.level.toUpperCase().padEnd(5);
    let message = `[${timestamp}] ${level} ${entry.message}`;

    if (entry.context) {
      message += ` ${JSON.stringify(entry.context)}`;
    }

    return message;
  }

  /**
   * Get recent logs
   */
  public getRecentLogs(count: number = 100): LogEntry[] {
    return this.logs.slice(-count);
  }

  /**
   * Get logs by level
   */
  public getLogsByLevel(level: 'debug' | 'info' | 'warn' | 'error'): LogEntry[] {
    return this.logs.filter((log) => log.level === level);
  }

  /**
   * Clear all logs
   */
  public clearLogs(): void {
    this.logs = [];
  }

  /**
   * Update log level
   */
  public setLogLevel(level: 'debug' | 'info' | 'warn' | 'error'): void {
    this.logLevel = level;
  }

  /**
   * Get current log level
   */
  public getLogLevel(): string {
    return this.logLevel;
  }
}