/**
 * Excel Coordinates Utility
 * Centralizes cell address, column letter, column index, and range parsing
 * Single Responsibility: Coordinate transformations and range calculations
 */

export interface ParsedCellAddress {
  row: number;
  column: number;
  colLetter: string;
}

export interface ParsedCellRange {
  start: ParsedCellAddress;
  end: ParsedCellAddress;
}

/**
 * Convert column letter to 1-based index ('A' -> 1, 'Z' -> 26, 'AA' -> 27)
 */
export function columnLetterToNumber(column: string): number {
  const upper = column.toUpperCase();
  let result = 0;
  for (let i = 0; i < upper.length; i++) {
    const code = upper.charCodeAt(i);
    if (code >= 65 && code <= 90) {
      result = result * 26 + (code - 64);
    }
  }
  return result;
}

/**
 * Alias for columnLetterToNumber
 */
export const columnToNumber = columnLetterToNumber;

/**
 * Convert 1-based column index to letter (1 -> 'A', 26 -> 'Z', 27 -> 'AA')
 */
export function numberToColumn(num: number): string {
  let s = '';
  let n = num;
  while (n > 0) {
    const t = (n - 1) % 26;
    s = String.fromCharCode(65 + t) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Parse cell address string (e.g. 'A1', 'BC99')
 */
export function parseCellAddress(address: string): ParsedCellAddress | null {
  const match = address.trim().match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;
  const colLetter = match[1].toUpperCase();
  const row = parseInt(match[2], 10);
  const column = columnLetterToNumber(colLetter);
  return { row, column, colLetter };
}

/**
 * Parse cell range string (e.g. 'A1:C10') or two separate address strings
 */
export function parseCellRange(startOrRange: string, maybeEnd?: string): ParsedCellRange | null {
  if (maybeEnd) {
    const start = parseCellAddress(startOrRange);
    const end = parseCellAddress(maybeEnd);
    if (!start || !end) return null;
    return {
      start: {
        row: Math.min(start.row, end.row),
        column: Math.min(start.column, end.column),
        colLetter: numberToColumn(Math.min(start.column, end.column)),
      },
      end: {
        row: Math.max(start.row, end.row),
        column: Math.max(start.column, end.column),
        colLetter: numberToColumn(Math.max(start.column, end.column)),
      },
    };
  }

  const parts = startOrRange.split(':');
  if (parts.length === 2) {
    return parseCellRange(parts[0], parts[1]);
  }

  // Single cell as 1x1 range
  const single = parseCellAddress(startOrRange);
  if (!single) return null;
  return { start: single, end: single };
}
