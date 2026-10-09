/**
 * Input Validator - Zod runtime validation for MCP tool arguments
 * Single Responsibility: build Zod schemas from ToolDefinition metadata and
 * validate incoming args before handler dispatch (audit Finding 2.1, Phase 3 C2).
 */

import { z } from 'zod';
import { ToolDefinition, ToolParameter } from '../types/index.js';

/**
 * Build a Zod type for a single declared tool parameter.
 *
 * Calibration note: legacy definitions declare flexible parameters such as the
 * cell "value" as type 'string' even though numbers and booleans are perfectly
 * legal cell values. To avoid regressing existing clients, 'string' parameters
 * accept the primitive union while still rejecting objects, arrays, and null.
 */
function buildParamSchema(param: ToolParameter): z.ZodTypeAny {
  let base: z.ZodTypeAny;
  switch (param.type) {
    case 'string':
      base = z.union([z.string(), z.number(), z.boolean()]);
      break;
    case 'number':
      base = z.number();
      break;
    case 'boolean':
      base = z.boolean();
      break;
    case 'array':
      base = z.array(z.unknown());
      break;
    case 'object':
      base = z.record(z.unknown());
      break;
    default:
      base = z.unknown();
  }

  if (param.enum && param.enum.length > 0) {
    base = z.enum(param.enum as [string, ...string[]]);
  }

  return param.required ? base : base.optional();
}

/**
 * Validate tool arguments against the tool's declared parameters.
 * Unknown keys are passed through (handlers own their optional extras).
 */
export function validateToolArgs(
  tool: ToolDefinition,
  args: Record<string, unknown>
): { success: true; data: Record<string, unknown> } | { success: false; error: string } {
  const shape: Record<string, z.ZodTypeAny> = {};
  for (const param of tool.parameters) {
    shape[param.name] = buildParamSchema(param);
  }
  const schema = z.object(shape).passthrough();

  const result = schema.safeParse(args ?? {});
  if (result.success) {
    return { success: true, data: result.data };
  }

  const issues = result.error.issues
    .map((issue) => ` "${issue.path.join('.') || '(root)'}": ${issue.message}`)
    .join(';');
  return {
    success: false,
    error: `Invalid parameters for ${tool.name}:${issues}`,
  };
}