/**
 * Excel Formula Analyzer - Analyzes and explains Excel formulas
 * Single Responsibility: Formula parsing, dependency tracking, and explanation
 */

import ExcelJS from 'exceljs';

export interface FormulaCell {
  address: string;
  formula: string;
  value: unknown;
  type: 'formula' | 'error' | 'value';
}

export interface FormulaInfo {
  formula: string;
  functions: string[];
  cellReferences: string[];
  namedRanges: string[];
  complexity: {
    depth: number;
    functionCount: number;
    referenceCount: number;
  };
  error?: string;
}

export interface DependencyInfo {
  cell: string;
  formula: string;
  precedents: string[];
  dependents: string[];
  isCircular: boolean;
  circularChain?: string[];
}

export interface FormulaExplanation {
  description: string;
  steps: string[];
  functions: Array<{ name: string; description: string }>;
  inputParameters: string[];
  outputType: string;
  suggestions: string[];
}

export class ExcelFormulaAnalyzer {
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;

  constructor(activeWorkbooks: Map<string, ExcelJS.Workbook>) {
    this.activeWorkbooks = activeWorkbooks;
  }

  /**
   * List all formulas in a worksheet
   */
  public async listFormulas(
    filename: string,
    worksheetName: string
  ): Promise<Array<FormulaCell>> {
    const workbook = this.activeWorkbooks.get(filename);
    if (!workbook) {
      throw new Error(`Workbook "${filename}" not found`);
    }

    const worksheet = workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }

    const formulas: FormulaCell[] = [];

    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.formula) {
          formulas.push({
            address: cell.address,
            formula: cell.formula,
            value: cell.result,
            type: cell.type === ExcelJS.ValueType.Error ? 'error' : 'formula',
          });
        }
      });
    });

    return formulas;
  }

  /**
   * Analyze a specific formula
   */
  public async analyzeFormula(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<FormulaInfo> {
    const workbook = this.activeWorkbooks.get(filename);
    if (!workbook) {
      throw new Error(`Workbook "${filename}" not found`);
    }

    const worksheet = workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }

    const cell = worksheet.getCell(cellAddress);
    if (!cell.formula) {
      throw new Error(`Cell ${cellAddress} does not contain a formula`);
    }

    const formulaStr = cell.formula;
    const functions = this.extractFunctions(formulaStr);
    const cellReferences = this.extractCellReferences(formulaStr);
    const namedRanges = this.extractNamedRanges(formulaStr);
    const complexity = this.calculateComplexity(formulaStr);

    const result: FormulaInfo = {
      formula: formulaStr,
      functions,
      cellReferences,
      namedRanges,
      complexity,
    };

    if (cell.type === ExcelJS.ValueType.Error) {
      result.error = String(cell.value);
    }

    return result;
  }

  /**
   * Get dependencies for a cell
   */
  public async getDependencies(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<DependencyInfo> {
    const workbook = this.activeWorkbooks.get(filename);
    if (!workbook) {
      throw new Error(`Workbook "${filename}" not found`);
    }

    const worksheet = workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${worksheetName}" not found`);
    }

    // Build dependency graph
    const graph = this.buildDependencyGraph(worksheet);
    
    // Get info for this cell
    const info = graph.get(cellAddress);
    if (!info) {
      throw new Error(`No formula found in cell ${cellAddress}`);
    }

    // Check for circular references
    const circularCheck = this.checkCircularReferences(graph, cellAddress);

    return {
      cell: cellAddress,
      formula: info.formula,
      precedents: info.precedents,
      dependents: info.dependents,
      isCircular: circularCheck.isCircular,
      circularChain: circularCheck.chain,
    };
  }

  /**
   * Explain a formula in Vietnamese
   */
  public async explainFormula(
    filename: string,
    worksheetName: string,
    cellAddress: string
  ): Promise<FormulaExplanation> {
    const info = await this.analyzeFormula(filename, worksheetName, cellAddress);
    
    const description = this.generateDescription(info);
    const steps = this.generateSteps(info);
    const functionDescriptions = this.getFunctionDescriptions(info.functions);
    const suggestions = this.generateSuggestions(info);

    const result: FormulaExplanation = {
      description,
      steps,
      functions: functionDescriptions,
      inputParameters: info.cellReferences,
      outputType: this.inferOutputType(info),
      suggestions,
    };

    return result;
  }

  /**
   * Audit all formulas in worksheet
   */
  public async auditFormulas(
    filename: string,
    worksheetName: string
  ): Promise<{
    totalFormulas: number;
    errors: Array<{ cell: string; formula: string; error: string }>;
    warnings: Array<{ cell: string; message: string }>;
    statistics: {
      mostUsedFunctions: Array<{ function: string; count: number }>;
      averageComplexity: number;
      maxComplexity: number;
    };
  }> {
    const formulas = await this.listFormulas(filename, worksheetName);
    
    const errors: Array<{ cell: string; formula: string; error: string }> = [];
    const warnings: Array<{ cell: string; message: string }> = [];
    const functionCounts = new Map<string, number>();
    let totalComplexity = 0;
    let maxComplexity = 0;

    for (const formula of formulas) {
      // Check for errors
      if (formula.type === 'error') {
        errors.push({
          cell: formula.address,
          formula: formula.formula,
          error: String(formula.value),
        });
      }

      // Analyze complexity
      const complexity = this.calculateComplexity(formula.formula);
      totalComplexity += complexity.depth;
      maxComplexity = Math.max(maxComplexity, complexity.depth);

      // Count functions
      const functions = this.extractFunctions(formula.formula);
      for (const func of functions) {
        const count = functionCounts.get(func) || 0;
        functionCounts.set(func, count + 1);
      }

      // Check for potential issues
      const potentialIssues = this.checkPotentialIssues(formula.formula);
      for (const issue of potentialIssues) {
        warnings.push({
          cell: formula.address,
          message: issue,
        });
      }
    }

    // Sort most used functions
    const mostUsedFunctions = Array.from(functionCounts.entries())
      .map(([func, count]) => ({ function: func, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    return {
      totalFormulas: formulas.length,
      errors,
      warnings,
      statistics: {
        mostUsedFunctions,
        averageComplexity: formulas.length > 0 ? totalComplexity / formulas.length : 0,
        maxComplexity,
      },
    };
  }

  // Private helper methods

  private buildDependencyGraph(
    worksheet: ExcelJS.Worksheet
  ): Map<string, { formula: string; precedents: string[]; dependents: string[] }> {
    const graph = new Map<string, { formula: string; precedents: string[]; dependents: string[] }>();

    // Collect all formulas
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.formula) {
          const references = this.extractCellReferences(cell.formula);
          graph.set(cell.address, {
            formula: cell.formula,
            precedents: references,
            dependents: [],
          });
        }
      });
    });

    // Link dependents
    for (const [address, info] of graph.entries()) {
      for (const ref of info.precedents) {
        const dependent = graph.get(ref);
        if (dependent) {
          dependent.dependents.push(address);
        }
      }
    }

    return graph;
  }

  private checkCircularReferences(
    graph: Map<string, { precedents: string[] }>,
    startCell: string
  ): { isCircular: boolean; chain?: string[] } {
    const visited = new Set<string>();
    const chain: string[] = [];

    const traverse = (cell: string): boolean => {
      if (visited.has(cell)) {
        const cycleStart = chain.indexOf(cell);
        if (cycleStart !== -1) {
          return true;
        }
        return false;
      }

      visited.add(cell);
      chain.push(cell);

      const info = graph.get(cell);
      if (info) {
        for (const precedent of info.precedents) {
          if (traverse(precedent)) {
            return true;
          }
        }
      }

      chain.pop();
      return false;
    };

    const isCircular = traverse(startCell);
    
    if (isCircular) {
      // Extract circular chain
      const cycleStart = chain.indexOf(chain[chain.length - 1]);
      return { isCircular: true, chain: chain.slice(cycleStart) };
    }

    return { isCircular: false };
  }

  private extractFunctions(formula: string): string[] {
    const functionPattern = /[A-Z][A-Z0-9\.]*/g;
    const matches = formula.match(functionPattern) || [];
    
    // Remove Excel keywords and operators
    const keywords = ['IF', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE'];
    return matches
      .map(m => m.toUpperCase())
      .filter(f => !keywords.includes(f))
      .filter((f, i, arr) => arr.indexOf(f) === i); // Remove duplicates
  }

  private extractCellReferences(formula: string): string[] {
    // Match cell references like A1, $A$1, A1:B10
    const patterns = [
      /\$?[A-Z]+\$?\d+/g, // Single cell (A1, $A$1)
      /\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+/g, // Range (A1:B10)
    ];

    const references: string[] = [];
    for (const pattern of patterns) {
      const matches = formula.match(pattern) || [];
      references.push(...matches);
    }

    return [...new Set(references)].sort();
  }

  private extractNamedRanges(formula: string): string[] {
    // Match named ranges (alphanumeric without cell reference pattern)
    const pattern = /[A-Z][A-Z0-9_]+/g;
    const matches = formula.match(pattern) || [];
    
    const cellRefPattern = /^[A-Z]+\d+$/;
    return matches
      .filter(m => !cellRefPattern.test(m))
      .filter((m, i, arr) => arr.indexOf(m) === i);
  }

  private calculateComplexity(formula: string): { depth: number; functionCount: number; referenceCount: number } {
    const functions = this.extractFunctions(formula);
    const references = this.extractCellReferences(formula);
    
    let depth = 0;
    let currentDepth = 0;
    
    for (const char of formula) {
      if (char === '(') {
        currentDepth++;
        depth = Math.max(depth, currentDepth);
      } else if (char === ')') {
        currentDepth--;
      }
    }

    return {
      depth,
      functionCount: functions.length,
      referenceCount: references.length,
    };
  }

  private checkPotentialIssues(formula: string): string[] {
    const issues: string[] = [];

    // Check for hard-coded values in formulas
    if (/\d+(?![A-Z])/g.test(formula)) {
      issues.push('Contains hard-coded numeric values - consider using named cells');
    }

    // Check for volatile functions
    const volatileFunctions = ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT'];
    const functions = this.extractFunctions(formula);
    const hasVolatile = functions.some(f => volatileFunctions.includes(f));
    if (hasVolatile) {
      issues.push('Contains volatile functions - may slow down calculations');
    }

    // Check for nested IFs
    const ifCount = functions.filter(f => f === 'IF').length;
    if (ifCount > 3) {
      issues.push('Many nested IF statements - consider using lookup table or CHOOSE');
    }

    return issues;
  }

  private generateDescription(info: FormulaInfo): string {
    if (info.functions.length === 0) {
      return `Công thức này tính toán giá trị từ các tham chiếu: ${info.cellReferences.join(', ')}`;
    }

    const mainFunction = info.functions[0];
    const description = this.getFunctionVietnameseName(mainFunction);
    
    return `Công thức này sử dụng hàm ${mainFunction} để ${description}. Nó tham chiếu đến ${info.cellReferences.length} cell: ${info.cellReferences.join(', ')}`;
  }

  private generateSteps(info: FormulaInfo): string[] {
    const steps: string[] = [];
    
    steps.push(`1. Đọc giá trị từ các cell: ${info.cellReferences.join(', ')}`);
    
    for (const func of info.functions) {
      steps.push(`2. Thực hiện hàm ${func}`);
    }
    
    steps.push(`3. Tính toán và trả về kết quả`);
    
    return steps;
  }

  private getFunctionDescriptions(functions: string[]): Array<{ name: string; description: string }> {
    return functions.map(func => ({
      name: func,
      description: this.getFunctionVietnameseName(func),
    }));
  }

  private getFunctionVietnameseName(func: string): string {
    const functionMap: Record<string, string> = {
      SUM: 'tính tổng các giá trị',
      AVERAGE: 'tính giá trị trung bình',
      COUNT: 'đếm số lượng giá trị',
      MAX: 'tìm giá trị lớn nhất',
      MIN: 'tìm giá trị nhỏ nhất',
      IF: 'kiểm tra điều kiện và trả về giá trị khác nhau',
      VLOOKUP: 'tìm kiếm giá trị trong bảng và trả về kết quả',
      HLOOKUP: 'tìm kiếm theo hàng ngang',
      INDEX: 'trả về giá trị tại vị trí cụ thể',
      MATCH: 'tìm vị trí của giá trị trong danh sách',
      CONCATENATE: 'nối các chuỗi lại với nhau',
      LEFT: 'trích xuất các ký tự bên trái',
      RIGHT: 'trích xuất các ký tự bên phải',
      MID: 'trích xuất các ký tự ở giữa',
      TRIM: 'xóa khoảng trắng thừa',
      UPPER: 'chuyển thành chữ hoa',
      LOWER: 'chuyển thành chữ thường',
      PROPER: 'viết hoa chữ cái đầu tiên',
      NOW: 'lấy ngày và giờ hiện tại',
      TODAY: 'lấy ngày hiện tại',
      DATE: 'tạo ngày từ các tham số',
      TIME: 'tạo thời gian từ các tham số',
      ROUND: 'làm tròn số',
      ROUNDUP: 'làm tròn lên',
      ROUNDDOWN: 'làm tròn xuống',
      INT: 'lấy phần nguyên',
      ABS: 'lấy giá trị tuyệt đối',
      POWER: 'tính lũy thừa',
      SQRT: 'tính căn bậc hai',
      AND: 'kiểm tra tất cả điều kiện đều đúng',
      OR: 'kiểm tra ít nhất một điều kiện đúng',
      NOT: 'đảo ngược kết quả logic',
      COUNTIF: 'đếm ô thỏa mãn điều kiện',
      SUMIF: 'tính tổng ô thỏa mãn điều kiện',
      AVERAGEIF: 'tính trung bình ô thỏa mãn điều kiện',
      COUNTA: 'đếm ô không rỗng',
      COUNTBLANK: 'đếm ô rỗng',
      OFFSET: 'trả về tham chiếu bị dịch chuyển',
      INDIRECT: 'trả về tham chiếu từ chuỗi văn bản',
      PMT: 'tính toán thanh toán khoản vay',
      NPV: 'tính giá trị hiện tại ròng',
      IRR: 'tính tỷ suất hoàn vốn nội bộ',
    };

    return functionMap[func] || `thực hiện hàm ${func}`;
  }

  private inferOutputType(info: FormulaInfo): string {
    const functions = info.functions;
    
    if (functions.includes('SUM') || functions.includes('AVERAGE') || functions.includes('MAX') || functions.includes('MIN')) {
      return 'Số (Number)';
    }
    
    if (functions.includes('COUNT') || functions.includes('COUNTIF') || functions.includes('COUNTA')) {
      return 'Số nguyên (Integer)';
    }
    
    if (functions.includes('CONCATENATE') || functions.includes('LEFT') || functions.includes('RIGHT') || functions.includes('MID')) {
      return 'Chuỗi (Text)';
    }
    
    if (functions.includes('IF') || functions.includes('AND') || functions.includes('OR')) {
      return 'Logic (Boolean)';
    }
    
    if (functions.includes('DATE') || functions.includes('NOW') || functions.includes('TODAY')) {
      return 'Ngày (Date)';
    }
    
    return 'Giá trị (Value)';
  }

  private generateSuggestions(info: FormulaInfo): string[] {
    const suggestions: string[] = [];

    // Check for complexity
    if (info.complexity.depth > 5) {
      suggestions.push('Công thức phức tạp - cân nhắc chia thành các cell trung gian');
    }

    // Check for many references
    if (info.complexity.referenceCount > 10) {
      suggestions.push('Nhiều tham chiếu - cân nhắc dùng named ranges hoặc table');
    }

    // Check for nested IFs
    const ifCount = info.functions.filter(f => f === 'IF').length;
    if (ifCount > 3) {
      suggestions.push('Nhiều hàm IF lồng nhau - cân nhắc dùng IFS hoặc VLOOKUP');
    }

    // Check for VLOOKUP optimization
    if (info.functions.includes('VLOOKUP')) {
      suggestions.push('VLOOKUP có thể chậm - cân nhắc dùng INDEX-MATCH thay thế');
    }

    if (suggestions.length === 0) {
      suggestions.push('Công thức nhìn tốt, không có vấn đề rõ ràng');
    }

    return suggestions;
  }
}
