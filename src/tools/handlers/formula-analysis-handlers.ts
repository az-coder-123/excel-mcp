/**
 * Formula Analysis Handlers
 * Single Responsibility: Handle formula analysis tool requests
 */

import { OperationResult } from '../../types/index.js';
import { ExcelFormulaAnalyzer } from '../../services/excel-formula-analyzer.js';

export class FormulaAnalysisHandlers {
  private formulaAnalyzer: ExcelFormulaAnalyzer;

  constructor(formulaAnalyzer: ExcelFormulaAnalyzer) {
    this.formulaAnalyzer = formulaAnalyzer;
  }

  /**
   * List all formulas in worksheet
   */
  public async handleListFormulas(args: {
    filename: string;
    worksheet: string;
  }): Promise<OperationResult<any>> {
    try {
      const formulas = await this.formulaAnalyzer.listFormulas(
        args.filename,
        args.worksheet
      );

      return {
        success: true,
        data: {
          total: formulas.length,
          formulas: formulas.map(f => ({
            cell: f.address,
            formula: f.formula,
            value: f.value,
            type: f.type,
          })),
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Analyze a specific formula
   */
  public async handleAnalyzeFormula(args: {
    filename: string;
    worksheet: string;
    cellAddress: string;
  }): Promise<OperationResult<any>> {
    try {
      const info = await this.formulaAnalyzer.analyzeFormula(
        args.filename,
        args.worksheet,
        args.cellAddress
      );

      return {
        success: true,
        data: {
          cell: args.cellAddress,
          formula: info.formula,
          functions: info.functions,
          cellReferences: info.cellReferences,
          namedRanges: info.namedRanges,
          complexity: info.complexity,
          error: info.error,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Get dependencies for a cell
   */
  public async handleGetDependencies(args: {
    filename: string;
    worksheet: string;
    cellAddress: string;
  }): Promise<OperationResult<any>> {
    try {
      const deps = await this.formulaAnalyzer.getDependencies(
        args.filename,
        args.worksheet,
        args.cellAddress
      );

      return {
        success: true,
        data: {
          cell: deps.cell,
          formula: deps.formula,
          precedents: deps.precedents,
          dependents: deps.dependents,
          isCircular: deps.isCircular,
          circularChain: deps.circularChain,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Trace precedents (data sources)
   */
  public async handleTracePrecedents(args: {
    filename: string;
    worksheet: string;
    cellAddress: string;
  }): Promise<OperationResult<any>> {
    try {
      const deps = await this.formulaAnalyzer.getDependencies(
        args.filename,
        args.worksheet,
        args.cellAddress
      );

      return {
        success: true,
        data: {
          cell: args.cellAddress,
          formula: deps.formula,
          precedents: deps.precedents,
          count: deps.precedents.length,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Trace dependents (cells using this value)
   */
  public async handleTraceDependents(args: {
    filename: string;
    worksheet: string;
    cellAddress: string;
  }): Promise<OperationResult<any>> {
    try {
      const deps = await this.formulaAnalyzer.getDependencies(
        args.filename,
        args.worksheet,
        args.cellAddress
      );

      return {
        success: true,
        data: {
          cell: args.cellAddress,
          dependents: deps.dependents,
          count: deps.dependents.length,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Check for circular references
   */
  public async handleCheckCircular(args: {
    filename: string;
    worksheet: string;
  }): Promise<OperationResult<any>> {
    try {
      const workbook = this.formulaAnalyzer['activeWorkbooks'].get(args.filename);
      if (!workbook) {
        return { success: false, error: `Workbook "${args.filename}" not found` };
      }

      const worksheet = workbook.getWorksheet(args.worksheet);
      if (!worksheet) {
        return { success: false, error: `Worksheet "${args.worksheet}" not found` };
      }

      // Build dependency graph
      const graph = this.formulaAnalyzer['buildDependencyGraph'](worksheet);
      
      // Check all cells for circular references
      const circularCells: Array<{ cell: string; chain: string[] }> = [];
      
      for (const [cell] of graph.keys()) {
        const check = this.formulaAnalyzer['checkCircularReferences'](graph, cell);
        if (check.isCircular && check.chain) {
          circularCells.push({ cell, chain: check.chain });
        }
      }

      return {
        success: true,
        data: {
          hasCircularReferences: circularCells.length > 0,
          circularCells,
          totalChecked: graph.size,
          circularCount: circularCells.length,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Explain a formula in Vietnamese
   */
  public async handleExplainFormula(args: {
    filename: string;
    worksheet: string;
    cellAddress: string;
  }): Promise<OperationResult<any>> {
    try {
      const explanation = await this.formulaAnalyzer.explainFormula(
        args.filename,
        args.worksheet,
        args.cellAddress
      );

      return {
        success: true,
        data: {
          cell: args.cellAddress,
          formula: (await this.formulaAnalyzer.analyzeFormula(
            args.filename,
            args.worksheet,
            args.cellAddress
          )).formula,
          description: explanation.description,
          steps: explanation.steps,
          functions: explanation.functions,
          inputParameters: explanation.inputParameters,
          outputType: explanation.outputType,
          suggestions: explanation.suggestions,
        },
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }

  /**
   * Audit all formulas in worksheet
   */
  public async handleAuditFormulas(args: {
    filename: string;
    worksheet: string;
  }): Promise<OperationResult<any>> {
    try {
      const audit = await this.formulaAnalyzer.auditFormulas(
        args.filename,
        args.worksheet
      );

      return {
        success: true,
        data: audit,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return { success: false, error: message };
    }
  }
}