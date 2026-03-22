/**
 * Protection Tool Definitions
 * Single Responsibility: Worksheet and workbook protection
 */

import { ToolDefinition } from '../../types/index.js';

export const PROTECTION_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_protect_worksheet',
    description: 'Protect worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Password',
        required: false,
      },
      {
        name: 'allowSelectLockedCells',
        type: 'boolean',
        description: 'Allow select locked',
        required: false,
      },
      {
        name: 'allowSelectUnlockedCells',
        type: 'boolean',
        description: 'Allow select unlocked',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_unprotect_worksheet',
    description: 'Unprotect worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Password',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_protect_cells',
    description: 'Lock/unlock cells',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Worksheet name',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Range start',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'Range end',
        required: true,
      },
      {
        name: 'locked',
        type: 'boolean',
        description: 'Lock (true/false)',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_protect_workbook',
    description: 'Protect workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Password',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_unprotect_workbook',
    description: 'Unprotect workbook',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Workbook name',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Password',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
];