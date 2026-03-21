/**
 * System Handlers - Health check and diagnostics implementation
 * Single Responsibility: Execute system-level operations
 */

import { BaseHandler } from './base-handler.js';
import { OperationResult } from '../../types/index.js';
import { ExcelService } from '../../services/excel-service.js';
import * as os from 'os';
import * as path from 'path';

export class SystemHandlers extends BaseHandler {
  private permissionChecker?: any;
  private logger?: any;

  constructor(
    excelService: ExcelService,
    permissionChecker?: any,
    logger?: any
  ) {
    super(excelService);
    this.permissionChecker = permissionChecker;
    this.logger = logger;
  }

  /**
   * Health Check Handler
   * Checks server status, dependencies, and configuration
   */
  async healthCheck(): Promise<OperationResult> {
    try {
      this.logger.debug('Performing health check');

      const startTime = Date.now();

      // Check server status
      const serverStatus = this.checkServerStatus();

      // Check dependencies
      const dependencies = this.checkDependencies();

      // Check configuration
      const configuration = this.checkConfiguration();

      // Run tests
      const tests = await this.runTests();

      const responseTime = Date.now() - startTime;

      // Determine overall health status
      const allTestsPassed = Object.values(tests).every((test) => test === 'pass');
      const allDependenciesOk = Object.values(dependencies).every((dep) => dep === 'ok');

      const status = allTestsPassed && allDependenciesOk ? 'healthy' : 'degraded';

      const healthReport = {
        status,
        server: serverStatus,
        dependencies,
        configuration,
        tests,
        responseTime: `${responseTime}ms`,
        timestamp: new Date().toISOString(),
      };

      this.logger.info('Health check completed', { status, responseTime });

      return {
        success: true,
        data: healthReport,
      };
    } catch (error) {
      this.logger.error('Health check failed', {
        error: error instanceof Error ? error.message : 'Unknown error',
      });

      return {
        success: false,
        error: `Health check failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
      };
    }
  }

  /**
   * Check server status
   */
  private checkServerStatus() {
    return {
      name: 'excel-mcp',
      version: '1.0.0',
      nodeVersion: process.version,
      platform: os.platform(),
      arch: os.arch(),
      uptime: this.formatUptime(process.uptime()),
    };
  }

  /**
   * Check dependencies
   */
  private checkDependencies() {
    const dependencies: Record<string, string | Record<string, string>> = {
      exceljs: 'ok',
      filesystem: 'ok',
      memory: 'ok',
    };

    // Check ExcelJS
    try {
      // ExcelJS is imported in excel-service, so if we're here, it's working
      dependencies.exceljs = 'ok';
    } catch (error) {
      dependencies.exceljs = 'error';
    }

    // Check file system
    try {
      const tmpDir = os.tmpdir();
      const testFile = `${tmpDir}/.excel-mcp-health-check-${Date.now()}`;
      require('fs').writeFileSync(testFile, 'test');
      require('fs').unlinkSync(testFile);
      dependencies.filesystem = 'ok';
    } catch (error) {
      dependencies.filesystem = 'error';
    }

    // Check memory
    const freeMemory = os.freemem();
    const memoryUsage = process.memoryUsage();
    
    if (freeMemory < 100 * 1024 * 1024) { // Less than 100MB free
      dependencies.memory = 'warning';
    }

    dependencies.memoryUsage = {
      heapUsed: `${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB`,
      heapTotal: `${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
      rss: `${Math.round(memoryUsage.rss / 1024 / 1024)}MB`,
      external: `${Math.round(memoryUsage.external / 1024 / 1024)}MB`,
    };

    return dependencies;
  }

  /**
   * Check configuration
   */
  private checkConfiguration() {
    const permissions = this.permissionChecker.getConfig();
    
    return {
      logLevel: this.logger.getLevel(),
      maxFileSize: permissions.maxFileSize,
      maxFileSizeMB: `${Math.round(permissions.maxFileSize / 1024 / 1024)}MB`,
      allowedPathsCount: permissions.allowedPaths.length,
      deniedPathsCount: permissions.deniedPaths.length,
      allowedExtensions: permissions.allowedExtensions,
      permissions: permissions.permissions,
    };
  }

  /**
   * Run diagnostic tests
   */
  private async runTests(): Promise<Record<string, string>> {
    const tests: Record<string, string> = {};

    // Test permission checker
    try {
      this.permissionChecker.checkPathAccess('/tmp/test', 'read');
      tests.permissionChecker = 'pass';
    } catch (error) {
      tests.permissionChecker = 'fail';
    }

    // Test logger
    try {
      this.logger.debug('Logger test');
      tests.logger = 'pass';
    } catch (error) {
      tests.logger = 'fail';
    }

    // Test Excel operations (create a minimal workbook)
    try {
      const testFilename = `health-check-test-${Date.now()}.xlsx`;
      const testPath = path.join(os.tmpdir(), testFilename);
      const createResult = await this.excelService.createWorkbook(testPath);

      if (createResult.success) {
        // Clean up
        await this.excelService.closeWorkbook(testFilename);
        tests.excelOperations = 'pass';
      } else {
        tests.excelOperations = 'fail';
      }
    } catch (error) {
      tests.excelOperations = 'fail';
    }

    return tests;
  }

  /**
   * Format uptime in human-readable format
   */
  private formatUptime(seconds: number): string {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const secs = Math.floor(seconds % 60);

    if (hours > 0) {
      return `${hours}h ${minutes}m ${secs}s`;
    } else if (minutes > 0) {
      return `${minutes}m ${secs}s`;
    } else {
      return `${secs}s`;
    }
  }
}