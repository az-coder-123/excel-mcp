/**
 * Analysis Tool Handlers
 * Single Responsibility: Handle data analysis, statistics, filtering, and profiling operations
 */

import ExcelJS from 'exceljs';

interface FilterCondition {
  column: string;
  operator: string;
  value?: string;
  matchCase?: boolean;
}

export class AnalysisHandlers {
  private workbookManager: Map<string, ExcelJS.Workbook>;

  constructor(workbookManager: Map<string, ExcelJS.Workbook>) {
    this.workbookManager = workbookManager;
  }

  private getWorkbook(filename: string): ExcelJS.Workbook | undefined {
    return this.workbookManager.get(filename);
  }

  /**
   * Get column statistics
   */
  async getColumnStats(
    filename: string,
    worksheet: string,
    column: string,
    hasHeader: boolean = true
  ) {
    try {
      const excelFile = this.getWorkbook(filename);
      if (!excelFile) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const sheet = excelFile.getWorksheet(worksheet);
      if (!sheet) {
        return { success: false, error: `Worksheet "${worksheet}" not found` };
      }

      const colIndex = column.charCodeAt(0) - 65; // A=0, B=1, etc.
      const startRow = hasHeader ? 2 : 1;
      
      const values: string[] = [];
      let rowCount = 0;

      sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
        if (rowNumber >= startRow) {
          const cell = row.getCell(colIndex + 1);
          const value = cell.value;
          
          if (value !== null && value !== undefined && value !== '') {
            values.push(String(value));
          }
          rowCount++;
        }
      });

      // Calculate statistics
      const uniqueValues = new Set(values);
      const valueCounts = new Map<string, number>();
      
      values.forEach(v => {
        valueCounts.set(v, (valueCounts.get(v) || 0) + 1);
      });

      const totalCells = rowCount - (hasHeader ? 1 : 0);
      const missingCount = totalCells - values.length;
      const missingPercentage = ((missingCount / totalCells) * 100).toFixed(2);

      // Sort by count
      const sortedValues = Array.from(valueCounts.entries())
        .sort((a, b) => b[1] - a[1]);

      return {
        success: true,
        data: {
          column,
          totalCells,
          nonEmptyCells: values.length,
          missingCells: missingCount,
          missingPercentage: `${missingPercentage}%`,
          uniqueCount: uniqueValues.size,
          duplicates: values.length - uniqueValues.size,
          valueDistribution: sortedValues.map(([value, count]) => ({
            value,
            count,
            percentage: `${((count / values.length) * 100).toFixed(2)}%`,
          })),
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error in getColumnStats',
      };
    }
  }

  /**
   * Filter data based on conditions
   */
  async filterData(
    filename: string,
    worksheet: string,
    startCell: string,
    endCell: string,
    filters: FilterCondition[],
    hasHeader: boolean = true
  ) {
    try {
      const excelFile = this.getWorkbook(filename);
      if (!excelFile) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const sheet = excelFile.getWorksheet(worksheet);
      if (!sheet) {
        return { success: false, error: `Worksheet "${worksheet}" not found` };
      }

      // Parse range
      const startColMatch = startCell.match(/[A-Z]+/);
      const startCol = startColMatch?.[0] ? startColMatch[0].charCodeAt(0) - 65 : 0;
      const startRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
      const endColMatch = endCell.match(/[A-Z]+/);
      const endCol = endColMatch?.[0] ? endColMatch[0].charCodeAt(0) - 65 : 0;
      const endRow = parseInt(endCell.match(/\d+/)?.[0] || '1');

      const headerRow: string[] = [];
      if (hasHeader) {
        const header = sheet.getRow(startRow);
        for (let c = startCol; c <= endCol; c++) {
          headerRow.push(header.getCell(c + 1).value as string || `Column ${c}`);
        }
      }

      const filteredRows: Record<string, unknown>[] = [];
      const dataStartRow = hasHeader ? startRow + 1 : startRow;

      sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
        if (rowNumber >= dataStartRow && rowNumber <= endRow) {
          let matchesAllFilters = true;

          for (const filter of filters) {
            const filterCol = filter.column.charCodeAt(0) - 65;
            const cellValue = row.getCell(filterCol + 1).value;
            const cellStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : '';
            const filterValue = filter.value || '';

            if (!this.evaluateFilter(cellStr, filter, filterValue)) {
              matchesAllFilters = false;
              break;
            }
          }

          if (matchesAllFilters) {
            const rowData: Record<string, unknown> = {};
            for (let c = startCol; c <= endCol; c++) {
              const cellValue = row.getCell(c + 1).value;
              const colLetter = String.fromCharCode(65 + c);
              rowData[colLetter] = cellValue;
            }
            filteredRows.push(rowData);
          }
        }
      });

      return {
        success: true,
        data: {
          headers: headerRow,
          totalRows: endRow - dataStartRow + 1,
          matchedRows: filteredRows.length,
          filteredRows,
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error in filterData',
      };
    }
  }

  /**
   * Group and aggregate data
   */
  async groupAggregate(
    filename: string,
    worksheet: string,
    startCell: string,
    endCell: string,
    groupByColumn: string,
    aggregateColumn: string,
    operation: string,
    hasHeader: boolean = true
  ) {
    try {
      const excelFile = this.getWorkbook(filename);
      if (!excelFile) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const sheet = excelFile.getWorksheet(worksheet);
      if (!sheet) {
        return { success: false, error: `Worksheet "${worksheet}" not found` };
      }

      const groupColIndex = groupByColumn.charCodeAt(0) - 65;
      const aggColIndex = aggregateColumn ? aggregateColumn.charCodeAt(0) - 65 : -1;
      
      const startRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
      const endRow = parseInt(endCell.match(/\d+/)?.[0] || '1');
      const dataStartRow = hasHeader ? startRow + 1 : startRow;

      const groups = new Map<string, unknown[]>();

      sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
        if (rowNumber >= dataStartRow && rowNumber <= endRow) {
          const groupValue = row.getCell(groupColIndex + 1).value;
          const groupKey = String(groupValue !== null && groupValue !== undefined ? groupValue : '');

          if (!groups.has(groupKey)) {
            groups.set(groupKey, []);
          }

          if (aggColIndex >= 0) {
            const aggValue = row.getCell(aggColIndex + 1).value;
            if (aggValue !== null && aggValue !== undefined && aggValue !== '') {
              groups.get(groupKey)!.push(aggValue);
            }
          } else {
            // Count operation
            groups.get(groupKey)!.push(1);
          }
        }
      });

      // Calculate aggregates
      const results: Record<string, unknown>[] = [];
      for (const [key, values] of groups.entries()) {
        let result: unknown;
        
        switch (operation) {
          case 'count':
            result = values.length;
            break;
          case 'sum':
            result = values.reduce((sum: number, v) => sum + (Number(v) || 0), 0);
            break;
          case 'avg':
            const sum = values.reduce((s: number, v) => s + (Number(v) || 0), 0);
            result = sum / values.length;
            break;
          case 'max':
            result = Math.max(...values.map(v => Number(v) || 0));
            break;
          case 'min':
            result = Math.min(...values.map(v => Number(v) || 0));
            break;
          default:
            result = values.length;
        }

        results.push({
          [groupByColumn]: key,
          [`${operation}${aggregateColumn ? `_${aggregateColumn}` : ''}`]: result,
          count: values.length,
        });
      }

      return {
        success: true,
        data: {
          groupByColumn,
          aggregateColumn: aggregateColumn || 'count',
          operation,
          groups: results.sort((a, b) => (b.count as number) - (a.count as number)),
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error in groupAggregate',
      };
    }
  }

  /**
   * Profile data for a range
   */
  async profileData(
    filename: string,
    worksheet: string,
    startCell: string,
    endCell: string,
    hasHeader: boolean = true
  ) {
    try {
      const excelFile = this.getWorkbook(filename);
      if (!excelFile) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const sheet = excelFile.getWorksheet(worksheet);
      if (!sheet) {
        return { success: false, error: `Worksheet "${worksheet}" not found` };
      }

      const startColMatch = startCell.match(/[A-Z]+/);
      const startCol = startColMatch?.[0] ? startColMatch[0].charCodeAt(0) - 65 : 0;
      const startRow = parseInt(startCell.match(/\d+/)?.[0] || '1');
      const endColMatch = endCell.match(/[A-Z]+/);
      const endCol = endColMatch?.[0] ? endColMatch[0].charCodeAt(0) - 65 : 0;
      const endRow = parseInt(endCell.match(/\d+/)?.[0] || '1');

      const columnProfiles: Record<string, unknown>[] = [];

      for (let col = startCol; col <= endCol; col++) {
        const colLetter = String.fromCharCode(65 + col);
        const values: string[] = [];
        const types = new Set<string>();
        let nullCount = 0;
        let totalCells = 0;

        sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
          if (rowNumber >= startRow && rowNumber <= endRow) {
            totalCells++;
            const cell = row.getCell(col + 1);
            const value = cell.value;

            if (value === null || value === undefined || value === '') {
              nullCount++;
            } else {
              const strValue = String(value);
              values.push(strValue);
              types.add(typeof value);
            }
          }
        });

        // Detect patterns
        const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        const phonePattern = /^[\d\s\-\+\(\)]+$/;
        const datePattern = /^\d{4}-\d{2}-\d{2}$|^\d{2}\/\d{2}\/\d{4}$/;

        const emails = values.filter(v => emailPattern.test(v)).length;
        const phones = values.filter(v => phonePattern.test(v)).length;
        const dates = values.filter(v => datePattern.test(v)).length;

        columnProfiles.push({
          column: colLetter,
          header: hasHeader ? sheet.getRow(startRow).getCell(col + 1).value : '',
          totalCells,
          nullCount,
          nullPercentage: `${((nullCount / totalCells) * 100).toFixed(2)}%`,
          uniqueValues: new Set(values).size,
          dataTypes: Array.from(types),
          detectedPatterns: {
            email: emails > 0 ? `${emails} found` : 'none',
            phone: phones > 0 ? `${phones} found` : 'none',
            date: dates > 0 ? `${dates} found` : 'none',
          },
        });
      }

      return {
        success: true,
        data: {
          worksheet,
          range: `${startCell}:${endCell}`,
          hasHeader,
          columnProfiles,
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error in profileData',
      };
    }
  }

  /**
   * Advanced search
   */
  async search(
    filename: string,
    worksheet: string,
    searchQuery: string,
    searchType: string = 'contains',
    columnRange: string = '',
    matchCase: boolean = false,
    maxResults: number = 100
  ) {
    try {
      const excelFile = this.getWorkbook(filename);
      if (!excelFile) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const sheet = excelFile.getWorksheet(worksheet);
      if (!sheet) {
        return { success: false, error: `Worksheet "${worksheet}" not found` };
      }

      let startCol = 0;
      let endCol = 26; // Default to all columns (A-Z)

      if (columnRange) {
        const match = columnRange.match(/^([A-Z]+):([A-Z]+)$/);
        if (match && match[1] && match[2]) {
          startCol = match[1].charCodeAt(0) - 65;
          endCol = match[2].charCodeAt(0) - 65;
        } else {
          startCol = columnRange.charCodeAt(0) - 65;
          endCol = startCol;
        }
      }

      const results: Record<string, unknown>[] = [];
      const query = matchCase ? searchQuery : searchQuery.toLowerCase();

      sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
        if (results.length >= maxResults) {
          return false;
        }

        for (let col = startCol; col <= endCol; col++) {
          const cell = row.getCell(col + 1);
          const value = cell.value;

          if (value !== null && value !== undefined) {
            const cellStr = String(value);
            const cellValue = matchCase ? cellStr : cellStr.toLowerCase();

            if (this.matchesSearch(cellValue, query, searchType)) {
              results.push({
                address: `${String.fromCharCode(65 + col)}${rowNumber}`,
                value: cellStr,
                row: rowNumber,
                column: String.fromCharCode(65 + col),
              });
              
              if (results.length >= maxResults) {
                return false;
              }
            }
          }
        }
        return true;
      });

      return {
        success: true,
        data: {
          searchQuery,
          searchType,
          totalMatches: results.length,
          maxResults,
          matches: results,
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error in search',
      };
    }
  }

  /**
   * Compare two ranges
   */
  async compareRanges(
    filename: string,
    worksheet1: string,
    range1Start: string,
    range1End: string,
    worksheet2: string = '',
    range2Start: string,
    range2End: string,
    compareType: string = 'values'
  ) {
    try {
      const excelFile = this.getWorkbook(filename);
      if (!excelFile) {
        return { success: false, error: `Workbook "${filename}" not found` };
      }

      const sheet1 = excelFile.getWorksheet(worksheet1);
      if (!sheet1) {
        return { success: false, error: `Worksheet "${worksheet1}" not found` };
      }

      const sheet2 = worksheet2 ? excelFile.getWorksheet(worksheet2) : sheet1;
      if (!sheet2) {
        return { success: false, error: `Worksheet "${worksheet2}" not found` };
      }

      // Parse ranges
      const parseCell = (cell: string) => {
        const colMatch = cell.match(/[A-Z]+/);
        const col = colMatch?.[0] ? colMatch[0].charCodeAt(0) - 65 : 0;
        const row = parseInt(cell.match(/\d+/)?.[0] || '1');
        return { col, row };
      };

      const r1Start = parseCell(range1Start);
      const r1End = parseCell(range1End);
      const r2Start = parseCell(range2Start);
      const r2End = parseCell(range2End);

      const differences: Record<string, unknown>[] = [];

      for (let r = 0; r <= Math.max(r1End.row - r1Start.row, r2End.row - r2Start.row); r++) {
        for (let c = 0; c <= Math.max(r1End.col - r1Start.col, r2End.col - r2Start.col); c++) {
          const cell1 = sheet1.getCell(r1Start.row + r, r1Start.col + c + 1).value;
          const cell2 = sheet2.getCell(r2Start.row + r, r2Start.col + c + 1).value;

          const isDifferent = cell1 !== cell2;

          if (isDifferent) {
            differences.push({
              address1: `${String.fromCharCode(65 + r1Start.col + c)}${r1Start.row + r}`,
              address2: `${String.fromCharCode(65 + r2Start.col + c)}${r2Start.row + r}`,
              value1: cell1,
              value2: cell2,
            });
          }
        }
      }

      return {
        success: true,
        data: {
          range1: `${range1Start}:${range1End}`,
          range2: `${range2Start}:${range2End}`,
          compareType,
          totalCellsCompared: differences.length,
          differentCells: differences.length,
          differences,
          matchPercentage: `${((1 - differences.length / (differences.length || 1)) * 100).toFixed(2)}%`,
        },
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error in compareRanges',
      };
    }
  }

  // Helper methods
  private evaluateFilter(cellValue: string, filter: FilterCondition, filterValue: string): boolean {
    const value = filter.matchCase ? cellValue : cellValue.toLowerCase();
    const compareValue = filter.matchCase ? filterValue : filterValue.toLowerCase();

    switch (filter.operator) {
      case 'equals':
        return value === compareValue;
      case 'notEquals':
        return value !== compareValue;
      case 'contains':
        return value.includes(compareValue);
      case 'notContains':
        return !value.includes(compareValue);
      case 'startsWith':
        return value.startsWith(compareValue);
      case 'endsWith':
        return value.endsWith(compareValue);
      case 'greaterThan':
        return parseFloat(value) > parseFloat(compareValue);
      case 'lessThan':
        return parseFloat(value) < parseFloat(compareValue);
      case 'greaterThanOrEqual':
        return parseFloat(value) >= parseFloat(compareValue);
      case 'lessThanOrEqual':
        return parseFloat(value) <= parseFloat(compareValue);
      case 'isBlank':
        return value === '';
      case 'isNotBlank':
        return value !== '';
      default:
        return false;
    }
  }

  private matchesSearch(cellValue: string, query: string, searchType: string): boolean {
    switch (searchType) {
      case 'exact':
        return cellValue === query;
      case 'contains':
        return cellValue.includes(query);
      case 'startsWith':
        return cellValue.startsWith(query);
      case 'endsWith':
        return cellValue.endsWith(query);
      case 'regex':
        try {
          const regex = new RegExp(query, 'i');
          return regex.test(cellValue);
        } catch {
          return false;
        }
      default:
        return cellValue.includes(query);
    }
  }
}