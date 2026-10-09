import { ExcelMCPServer } from '../src/server/excel-mcp-server.js';

describe('ExcelMCPServer — configuration & default containment (audit Finding 5.1)', () => {
  it('defaults unconfigured allowedPaths to [process.cwd()] for default-deny containment', () => {
    const server = new ExcelMCPServer({ logLevel: 'silent' });
    const config = server.getConfig();
    expect(config.permissions.allowedPaths).toEqual([process.cwd()]);
  });

  it('allows explicit override of allowedPaths', () => {
    const server = new ExcelMCPServer({
      logLevel: 'silent',
      permissions: {
        allowedPaths: ['/custom/allowed/dir'],
        deniedPaths: [],
        maxFileSize: 1024,
        allowedExtensions: ['.xlsx'],
        permissions: ['read'],
      },
    });
    const config = server.getConfig();
    expect(config.permissions.allowedPaths).toEqual(['/custom/allowed/dir']);
  });
});
