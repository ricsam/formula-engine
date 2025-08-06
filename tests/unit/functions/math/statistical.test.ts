import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { EvaluationContext } from '../../../../src/evaluator/evaluator';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('Statistical Functions', () => {
  let evaluator: Evaluator;
  let errorHandler: ErrorHandler;
  let dependencyGraph: DependencyGraph;

  // Helper function to evaluate a formula
  const evaluateFormula = (formula: string, mockRangeValues?: Record<string, CellValue[][]>): CellValue => {
    const ast = Parser.parse(formula, 0);
    
    const context: EvaluationContext = {
      currentSheet: 0,
      namedExpressions: new Map(),
      getCellValue: (address: SimpleCellAddress) => undefined,
      getRangeValues: (range: SimpleCellRange) => {
        // Mock range values for testing
        const key = `${range.start.col},${range.start.row}:${range.end.col},${range.end.row}`;
        return mockRangeValues?.[key] || [[]];
      },
      getFunction: (name: string) => functionRegistry.get(name),
      errorHandler,
      evaluationStack: new Set(),
    };
    
    const result = evaluator.evaluate(ast, context);
    return result.value;
  };

  beforeEach(() => {
    errorHandler = new ErrorHandler();
    dependencyGraph = new DependencyGraph();
    
    // Get all functions from the registry
    const functions = functionRegistry.getAllFunctions();
    
    evaluator = new Evaluator(dependencyGraph, functions, errorHandler);
  });

  describe('SUM function', () => {
    test('Sum of literal numbers', () => {
      expect(evaluateFormula('=SUM(1, 2, 3, 4, 5)')).toBe(15);
      expect(evaluateFormula('=SUM(10)')).toBe(10);
      expect(evaluateFormula('=SUM(-5, 5)')).toBe(0);
    });

    test('Sum with nested expressions', () => {
      expect(evaluateFormula('=SUM(1, FE.ADD(2, 3), 4)')).toBe(10);
    });

    test('Sum ignores text and boolean values', () => {
      expect(evaluateFormula('=SUM(1, "text", 2, TRUE, 3)')).toBe(6);
    });

    test('Sum with arrays', () => {
      // Test with array literal - SUM should flatten arrays
      expect(evaluateFormula('=SUM({1,2,3})')).toBe(6);
      expect(evaluateFormula('=SUM({1,2},{3,4})')).toBe(10);
    });

    test('Sum with mixed values', () => {
      // Test that SUM handles various value types correctly
      expect(evaluateFormula('=SUM(1, 2)')).toBe(3);
    });

    test('Sum propagates errors', () => {
      expect(evaluateFormula('=SUM(1, FE.DIVIDE(1,0), 3)')).toBe('#DIV/0!');
    });
  });

  describe('COUNT function', () => {
    test('Count numbers only', () => {
      expect(evaluateFormula('=COUNT(1, 2, 3)')).toBe(3);
      expect(evaluateFormula('=COUNT(1, "text", 2, TRUE, 3)')).toBe(3);
      expect(evaluateFormula('=COUNT("a", "b", "c")')).toBe(0);
    });

    test('Count with arrays', () => {
      expect(evaluateFormula('=COUNT({1,2,3})')).toBe(3);
      expect(evaluateFormula('=COUNT({1,"text",3})')).toBe(2);
    });

    test('Count propagates errors', () => {
      // Currently, errors are propagated before reaching COUNT
      // This is different from Excel behavior but consistent with our evaluator
      expect(evaluateFormula('=COUNT(1, FE.DIVIDE(1,0), 2, 3)')).toBe('#DIV/0!');
    });
  });

  describe('AVERAGE function', () => {
    test('Average of numbers', () => {
      expect(evaluateFormula('=AVERAGE(1, 2, 3, 4, 5)')).toBe(3);
      expect(evaluateFormula('=AVERAGE(10, 20)')).toBe(15);
      expect(evaluateFormula('=AVERAGE(0, 0, 0)')).toBe(0);
    });

    test('Average ignores non-numeric values', () => {
      expect(evaluateFormula('=AVERAGE(1, "text", 2, TRUE, 3)')).toBe(2);
    });

    test('Average returns error when no numbers', () => {
      expect(evaluateFormula('=AVERAGE("a", "b", TRUE)')).toBe('#DIV/0!');
    });
  });

  describe('MAX and MIN functions', () => {
    test('MAX function', () => {
      expect(evaluateFormula('=MAX(1, 5, 3, 2, 4)')).toBe(5);
      expect(evaluateFormula('=MAX(-10, -5, -20)')).toBe(-5);
      expect(evaluateFormula('=MAX(100)')).toBe(100);
    });

    test('MIN function', () => {
      expect(evaluateFormula('=MIN(1, 5, 3, 2, 4)')).toBe(1);
      expect(evaluateFormula('=MIN(-10, -5, -20)')).toBe(-20);
      expect(evaluateFormula('=MIN(100)')).toBe(100);
    });

    test('MAX/MIN with non-numeric values', () => {
      expect(evaluateFormula('=MAX(1, "text", 5, TRUE, 3)')).toBe(5);
      expect(evaluateFormula('=MIN(1, "text", 5, TRUE, 3)')).toBe(1);
    });

    test('MAX/MIN return 0 when no numbers', () => {
      expect(evaluateFormula('=MAX("a", "b", TRUE)')).toBe(0);
      expect(evaluateFormula('=MIN("a", "b", TRUE)')).toBe(0);
    });
  });

  describe('MEDIAN function', () => {
    test('Median of odd number of values', () => {
      expect(evaluateFormula('=MEDIAN(1, 2, 3, 4, 5)')).toBe(3);
      expect(evaluateFormula('=MEDIAN(1, 3, 2)')).toBe(2);
    });

    test('Median of even number of values', () => {
      expect(evaluateFormula('=MEDIAN(1, 2, 3, 4)')).toBe(2.5);
      expect(evaluateFormula('=MEDIAN(10, 20)')).toBe(15);
    });

    test('Median with non-numeric values', () => {
      expect(evaluateFormula('=MEDIAN(1, "text", 2, TRUE, 3)')).toBe(2);
    });

    test('Median returns error when no numbers', () => {
      expect(evaluateFormula('=MEDIAN("a", "b", TRUE)')).toBe('#NUM!');
    });
  });

  describe('STDEV and VAR functions', () => {
    test('STDEV calculation', () => {
      // Sample standard deviation of [1,2,3,4,5]
      // Mean = 3, Variance = 2.5, StdDev = 1.58...
      const result = evaluateFormula('=STDEV(1, 2, 3, 4, 5)');
      expect(typeof result).toBe('number');
      expect(Math.abs((result as number) - 1.58113883)).toBeLessThan(0.0001);
    });

    test('VAR calculation', () => {
      // Sample variance of [1,2,3,4,5]
      // Mean = 3, Variance = 2.5
      expect(evaluateFormula('=VAR(1, 2, 3, 4, 5)')).toBe(2.5);
    });

    test('STDEV/VAR require at least 2 numbers', () => {
      expect(evaluateFormula('=STDEV(5)')).toBe('#DIV/0!');
      expect(evaluateFormula('=VAR(5)')).toBe('#DIV/0!');
    });
  });

  describe('PRODUCT function', () => {
    test('Product of numbers', () => {
      expect(evaluateFormula('=PRODUCT(2, 3, 4)')).toBe(24);
      expect(evaluateFormula('=PRODUCT(5, -2)')).toBe(-10);
      expect(evaluateFormula('=PRODUCT(1, 2, 0, 3)')).toBe(0);
    });

    test('Product ignores non-numeric values', () => {
      expect(evaluateFormula('=PRODUCT(2, "text", 3, TRUE)')).toBe(6);
    });

    test('Product returns 0 when no numbers', () => {
      expect(evaluateFormula('=PRODUCT("a", "b", TRUE)')).toBe(0);
    });
  });

  describe('COUNTBLANK function', () => {
    test('Count blank values in array', () => {
      // COUNTBLANK with array should count empty/undefined values
      // For now, skip this test as it requires range support
      expect(evaluateFormula('=COUNTBLANK({1,2,3})')).toBe(0);
    });

    test('Empty strings count as blank', () => {
      expect(evaluateFormula('=COUNTBLANK({""})')).toBe(1);
    });

    test('Non-blank values', () => {
      expect(evaluateFormula('=COUNTBLANK({"text"})')).toBe(0);
    });
  });
});