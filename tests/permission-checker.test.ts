/**
 * Security tests for PermissionChecker
 *
 * Regression coverage for audit findings:
 * - 5.1: Path traversal (`../`) prevention via canonical path resolution
 * - 5.2: Fail-open default when allowedPaths is empty (pinned as documented behavior)
 * Plus: denied-path priority, wildcards, extension whitelist, size limits,
 * operation permissions, and composite validateFileAccess ordering.
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { PermissionChecker } from '../src/security/permission-checker.js';
import { PermissionConfig } from '../src/types/index.js';

const makeConfig = (overrides: Partial<PermissionConfig> = {}): PermissionConfig => ({
  allowedPaths: [],
  deniedPaths: [],
  maxFileSize: 50 * 1024 * 1024,
  allowedExtensions: ['.xlsx', '.xls', '.xlsm', '.xlsb'],
  permissions: ['read', 'write', 'delete'],
  ...overrides,
});

describe('PermissionChecker — path containment (audit Finding 5.1: path traversal)', () => {
  const checker = new PermissionChecker(makeConfig({ allowedPaths: ['/workspace/project'] }));

  it('allows files directly inside an allowed directory', () => {
    expect(checker.isPathAllowed('/workspace/project/report.xlsx').success).toBe(true);
  });

  it('allows nested subdirectories inside an allowed directory', () => {
    expect(checker.isPathAllowed('/workspace/project/sub/dir/report.xlsx').success).toBe(true);
  });

  it('BLOCKS ../ traversal escaping the sandbox', () => {
    const result = checker.isPathAllowed('/workspace/project/../../etc/passwd.xlsx');
    expect(result.success).toBe(false);
  });

  it('BLOCKS ../ traversal targeting a sibling directory', () => {
    const result = checker.isPathAllowed('/workspace/project/../project2/secret.xlsx');
    expect(result.success).toBe(false);
  });

  it('BLOCKS deep traversal chains resolving outside the allowed root', () => {
    const result = checker.isPathAllowed('/workspace/project/a/b/../../../../root/secret.xlsx');
    expect(result.success).toBe(false);
  });

  it('allows ../ segments that still resolve INSIDE the allowed directory', () => {
    expect(checker.isPathAllowed('/workspace/project/sub/../data.xlsx').success).toBe(true);
  });

  it('BLOCKS a sibling directory sharing the same string prefix (project2 vs project)', () => {
    expect(checker.isPathAllowed('/workspace/project2/secret.xlsx').success).toBe(false);
  });

  it('BLOCKS absolute paths unrelated to any allowed root', () => {
    expect(checker.isPathAllowed('/etc/passwd.xlsx').success).toBe(false);
  });

  it('matches case-insensitively (macOS/Windows default filesystems)', () => {
    const caseChecker = new PermissionChecker(makeConfig({ allowedPaths: ['/Workspace/Project'] }));
    expect(caseChecker.isPathAllowed('/workspace/project/file.xlsx').success).toBe(true);
  });
});

describe('PermissionChecker — denied paths take priority over allowed paths', () => {
  const checker = new PermissionChecker(
    makeConfig({ allowedPaths: ['/data'], deniedPaths: ['/data/secret/*'] })
  );

  it('blocks paths matching a denied wildcard even though the parent dir is allowed', () => {
    expect(checker.isPathAllowed('/data/secret/priv.xlsx').success).toBe(false);
  });

  it('allows non-denied paths under the allowed root', () => {
    expect(checker.isPathAllowed('/data/normal.xlsx').success).toBe(true);
  });

  it('blocks denied paths even when allowedPaths is empty (allow-all mode)', () => {
    const denyChecker = new PermissionChecker(makeConfig({ deniedPaths: ['/etc/*'] }));
    expect(denyChecker.isPathAllowed('/etc/passwd').success).toBe(false);
  });
});

describe('PermissionChecker — fail-open default when allowedPaths is empty (audit Finding 5.2)', () => {
  it('currently allows any path when no allowedPaths are configured (pinned documented risk)', () => {
    const checker = new PermissionChecker(makeConfig());
    expect(checker.isPathAllowed('/anywhere/on/disk/file.xlsx').success).toBe(true);
  });
});

describe('PermissionChecker — wildcard patterns', () => {
  const checker = new PermissionChecker(makeConfig({ allowedPaths: ['/data/excel/*.xlsx'] }));

  it('matches wildcard extensions', () => {
    expect(checker.isPathAllowed('/data/excel/report.xlsx').success).toBe(true);
  });

  it('does not match other extensions', () => {
    expect(checker.isPathAllowed('/data/excel/report.txt').success).toBe(false);
  });

  it('treats * as matching across path separators (stronger deny for /etc/* style rules)', () => {
    // Documented behavior: '*' compiles to '.*' and therefore crosses '/'.
    // This is intentionally kept: it makes denied wildcard rules like
    // '/etc/*' also cover nested paths such as '/etc/sub/x'.
    expect(checker.isPathAllowed('/data/excel/sub/report.xlsx').success).toBe(true);
  });
});

describe('PermissionChecker — extension whitelist', () => {
  const checker = new PermissionChecker(makeConfig());

  it('allows whitelisted extensions', () => {
    expect(checker.isExtensionAllowed('/data/report.xlsx').success).toBe(true);
    expect(checker.isExtensionAllowed('/data/report.xlsm').success).toBe(true);
  });

  it('blocks non-whitelisted extensions', () => {
    const result = checker.isExtensionAllowed('/data/report.txt');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/extension/i);
  });

  it('compares extensions case-insensitively', () => {
    expect(checker.isExtensionAllowed('/data/REPORT.XLSX').success).toBe(true);
  });

  it('blocks extension-less paths when a whitelist is configured', () => {
    expect(checker.isExtensionAllowed('/data/noextension').success).toBe(false);
  });

  it('allows any extension when the whitelist is empty (pinned behavior)', () => {
    const lax = new PermissionChecker(makeConfig({ allowedExtensions: [] }));
    expect(lax.isExtensionAllowed('/data/anything.xyz').success).toBe(true);
  });
});

describe('PermissionChecker — file size limits', () => {
  const checker = new PermissionChecker(makeConfig({ maxFileSize: 1000 }));

  it('allows sizes at or under the limit', () => {
    expect(checker.isFileSizeAllowed(1000).success).toBe(true);
  });

  it('blocks sizes over the limit', () => {
    const result = checker.isFileSizeAllowed(1001);
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/exceeds/i);
  });
});

describe('PermissionChecker — operation permissions', () => {
  const readOnly = new PermissionChecker(makeConfig({ permissions: ['read'] }));

  it('grants a configured permission', () => {
    expect(readOnly.hasPermission('read').success).toBe(true);
  });

  it('denies an unconfigured permission', () => {
    const result = readOnly.hasPermission('write');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/permission/i);
  });

  it('admin bypasses all permission checks', () => {
    const admin = new PermissionChecker(makeConfig({ permissions: ['admin'] }));
    expect(admin.hasPermission('write').success).toBe(true);
    expect(admin.hasPermission('delete').success).toBe(true);
  });
});

describe('PermissionChecker — validateFileAccess composite checks', () => {
  it('passes when every check succeeds', () => {
    const checker = new PermissionChecker(
      makeConfig({ allowedPaths: ['/data'], permissions: ['read', 'write'] })
    );
    expect(checker.validateFileAccess('/data/report.xlsx', 10, 'write').success).toBe(true);
  });

  it('fails on missing permission before path is even considered', () => {
    const checker = new PermissionChecker(
      makeConfig({ allowedPaths: ['/data'], permissions: ['read'] })
    );
    const result = checker.validateFileAccess('/data/report.xlsx', 0, 'write');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/permission/i);
  });

  it('fails on a disallowed path', () => {
    const checker = new PermissionChecker(makeConfig({ allowedPaths: ['/data'] }));
    const result = checker.validateFileAccess('/elsewhere/report.xlsx', 0, 'read');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/does not match/i);
  });

  it('fails on a disallowed extension', () => {
    const checker = new PermissionChecker(makeConfig({ allowedPaths: ['/data'] }));
    const result = checker.validateFileAccess('/data/report.txt', 0, 'read');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/extension/i);
  });

  it('fails on an oversized file when size is provided', () => {
    const checker = new PermissionChecker(makeConfig({ allowedPaths: ['/data'], maxFileSize: 100 }));
    const result = checker.validateFileAccess('/data/report.xlsx', 500, 'read');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/exceeds/i);
  });
});

describe('PermissionChecker — relative path resolution is cwd-aware', () => {
  const originalCwd = process.cwd();
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'perm-cwd-'));
    // macOS: os.tmpdir() returns a symlinked path (/var/...) while process.cwd()
    // reports the canonical path (/private/var/...). Canonicalize the pattern
    // so prefix comparison aligns with what path.resolve produces.
    tmpDir = fs.realpathSync(tmpDir);
    process.chdir(tmpDir);
  });

  afterEach(() => {
    process.chdir(originalCwd);
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('resolves a bare relative filename against cwd inside the allowed root', () => {
    const checker = new PermissionChecker(makeConfig({ allowedPaths: [tmpDir] }));
    expect(checker.isPathAllowed('file.xlsx').success).toBe(true);
  });

  it('blocks a relative ../ escape from cwd', () => {
    const checker = new PermissionChecker(makeConfig({ allowedPaths: [tmpDir] }));
    expect(checker.isPathAllowed('../outside.xlsx').success).toBe(false);
  });
});
