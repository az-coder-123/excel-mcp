/**
 * Zod runtime input validation tests (Phase 3 C2 / audit Finding 2.1)
 *
 * validateToolArgs builds Zod schemas from ToolDefinition metadata and gates
 * every executeTool dispatch. These tests pin the validation contract,
 * including the deliberately lenient primitive union for 'string' params
 * (legacy definitions declare cell "value" as string though numbers and
 * booleans are legal cell values).
 */
import { validateToolArgs } from '../src/tools/input-validator.js';
import { ToolDefinition } from '../src/types/index.js';

const TOOL: ToolDefinition = {
  name: 'excel_test_tool',
  description: 'fixture',
  parameters: [
    { name: 'filename', type: 'string', description: 'file', required: true },
    { name: 'count', type: 'number', description: 'n', required: true },
    { name: 'flag', type: 'boolean', description: 'b', required: false },
    { name: 'mode', type: 'string', description: 'm', required: true, enum: ['fast', 'slow'] },
    { name: 'value', type: 'string', description: 'cell value', required: true },
    { name: 'tags', type: 'array', description: 't', required: false },
    { name: 'options', type: 'object', description: 'o', required: false },
  ],
  requiredPermissions: ['read'],
};

const VALID = {
  filename: 'report.xlsx',
  count: 3,
  mode: 'fast',
  value: 'hello',
};

describe('validateToolArgs — happy paths', () => {
  it('accepts fully valid arguments and returns the parsed data', () => {
    const result = validateToolArgs(TOOL, { ...VALID, flag: true });
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.filename).toBe('report.xlsx');
    }
  });

  it('accepts omitted optional parameters', () => {
    expect(validateToolArgs(TOOL, VALID).success).toBe(true);
  });

  it("accepts numbers for 'string' params (lenient union for cell values)", () => {
    expect(validateToolArgs(TOOL, { ...VALID, value: 42 }).success).toBe(true);
  });

  it("accepts booleans for 'string' params (lenient union for cell values)", () => {
    expect(validateToolArgs(TOOL, { ...VALID, value: false }).success).toBe(true);
  });

  it('passes unknown keys through (handlers own their extras)', () => {
    const result = validateToolArgs(TOOL, { ...VALID, surprise: 'ok' });
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.surprise).toBe('ok');
    }
  });

  it('accepts arrays for array params and objects for object params', () => {
    expect(validateToolArgs(TOOL, { ...VALID, tags: ['a', 'b'], options: { x: 1 } }).success).toBe(true);
  });

  it('accepts anything for a tool with zero declared parameters', () => {
    const bare: ToolDefinition = {
      name: 'excel_bare',
      description: 'no params',
      parameters: [],
      requiredPermissions: ['read'],
    };
    expect(validateToolArgs(bare, { anything: 1 }).success).toBe(true);
  });
});

describe('validateToolArgs — rejections', () => {
  it('rejects a missing required parameter and names it', () => {
    const { filename, ...rest } = VALID;
    void filename;
    const result = validateToolArgs(TOOL, rest);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error).toContain('excel_test_tool');
      expect(result.error).toContain('filename');
    }
  });

  it("rejects a non-number for 'number' params", () => {
    const result = validateToolArgs(TOOL, { ...VALID, count: 'three' });
    expect(result.success).toBe(false);
  });

  it('rejects a value outside the declared enum', () => {
    const result = validateToolArgs(TOOL, { ...VALID, mode: 'warp' });
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error).toContain('mode');
    }
  });

  it("rejects objects for 'string' params", () => {
    const result = validateToolArgs(TOOL, { ...VALID, filename: { nested: true } });
    expect(result.success).toBe(false);
  });

  it("rejects null for required 'string' params", () => {
    const result = validateToolArgs(TOOL, { ...VALID, filename: null });
    expect(result.success).toBe(false);
  });

  it('rejects a non-boolean for boolean params when provided', () => {
    const result = validateToolArgs(TOOL, { ...VALID, flag: 'yes' });
    expect(result.success).toBe(false);
  });

  it('rejects a non-array for array params when provided', () => {
    const result = validateToolArgs(TOOL, { ...VALID, tags: 'not-an-array' });
    expect(result.success).toBe(false);
  });
});
