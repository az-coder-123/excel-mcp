/**
 * Protection Tool Definitions
 * Single Responsibility: Worksheet and workbook protection
 */

import { ToolDefinition } from '../../types/index.js';

export const PROTECTION_TOOLS: ToolDefinition[] = [
  {
    name: 'excel_protect_worksheet',
    description: 'Protect a worksheet with password',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet to protect',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Protection password',
        required: false,
      },
      {
        name: 'allowSelectLockedCells',
        type: 'boolean',
        description: 'Allow selecting locked cells',
        required: false,
      },
      {
        name: 'allowSelectUnlockedCells',
        type: 'boolean',
        description: 'Allow selecting unlocked cells',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_unprotect_worksheet',
    description: 'Remove protection from a worksheet',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Protection password',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_protect_cells',
    description: 'Lock or unlock specific cells',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'worksheet',
        type: 'string',
        description: 'Name of the worksheet',
        required: true,
      },
      {
        name: 'startCell',
        type: 'string',
        description: 'Start cell of range',
        required: true,
      },
      {
        name: 'endCell',
        type: 'string',
        description: 'End cell of range',
        required: true,
      },
      {
        name: 'locked',
        type: 'boolean',
        description: 'Lock (true) or unlock (false) cells',
        required: true,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_protect_workbook',
    description: 'Protect workbook structure',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Protection password',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
  {
    name: 'excel_unprotect_workbook',
    description: 'Remove workbook protection',
    parameters: [
      {
        name: 'filename',
        type: 'string',
        description: 'Name of the opened workbook',
        required: true,
      },
      {
        name: 'password',
        type: 'string',
        description: 'Protection password',
        required: false,
      },
    ],
    requiredPermissions: ['write'],
  },
];