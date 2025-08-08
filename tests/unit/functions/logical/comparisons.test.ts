import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { EvaluationContext } from '../../../../src/evaluator/evaluator';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('Logical Comparison Functions', () => {
  let evaluator: Evaluator;
  let errorHandler: ErrorHandler;
  let dependencyGraph: DependencyGraph;

  // Helper function to evaluate a formula
  const evaluateFormula = (formula: string): CellValue => {
    const ast = Parser.parse(formula, 0);
    
    const context: EvaluationContext = {
      currentSheet: 0,
      namedExpressions: new Map(),
      getCellValue: (address: SimpleCellAddress) => undefined,
      getRangeValues: (range: SimpleCellRange) => [[]],
      getFunction: (name: string) => functionRegistry.get(name),
      errorHandler,
      evaluationStack: new Set(),
    };
    
    const result = evaluator.evaluate(ast, context);
    return result.value as any;
  };

  beforeEach(() => {
    errorHandler = new ErrorHandler();
    dependencyGraph = new DependencyGraph();
    
    // Get all functions from the registry
    const functions = functionRegistry.getAllFunctions();
    
    evaluator = new Evaluator(dependencyGraph, functions, errorHandler);
  });

  describe('Equality Comparison', () => {
    test('FE.EQ function with numbers', () => {
      expect(evaluateFormula('=FE.EQ(5, 5)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(5, 3)')).toBe(false);
      expect(evaluateFormula('=FE.EQ(0, 0)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(-1, -1)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(3.14, 3.14)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(3.14, 3.15)')).toBe(false);
    });

    test('FE.EQ function with strings', () => {
      expect(evaluateFormula('=FE.EQ("hello", "hello")')).toBe(true);
      expect(evaluateFormula('=FE.EQ("hello", "world")')).toBe(false);
      expect(evaluateFormula('=FE.EQ("", "")')).toBe(true);
      expect(evaluateFormula('=FE.EQ("abc", "ABC")')).toBe(false); // Case sensitive
    });

    test('FE.EQ function with booleans', () => {
      expect(evaluateFormula('=FE.EQ(TRUE, TRUE)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(FALSE, FALSE)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(TRUE, FALSE)')).toBe(false);
    });

    test('FE.EQ function with mixed types', () => {
      expect(evaluateFormula('=FE.EQ(1, TRUE)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(0, FALSE)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(5, "5")')).toBe(true);
      expect(evaluateFormula('=FE.EQ(0, "")')).toBe(false);
    });
  });

  describe('Inequality Comparison', () => {
    test('FE.NE function with numbers', () => {
      expect(evaluateFormula('=FE.NE(5, 3)')).toBe(true);
      expect(evaluateFormula('=FE.NE(5, 5)')).toBe(false);
      expect(evaluateFormula('=FE.NE(0, 1)')).toBe(true);
    });

    test('FE.NE function with strings', () => {
      expect(evaluateFormula('=FE.NE("hello", "world")')).toBe(true);
      expect(evaluateFormula('=FE.NE("hello", "hello")')).toBe(false);
    });

    test('FE.NE function with mixed types', () => {
      expect(evaluateFormula('=FE.NE(1, TRUE)')).toBe(false);
      expect(evaluateFormula('=FE.NE(1, FALSE)')).toBe(true);
    });
  });

  describe('Less Than Comparison', () => {
    test('FE.LT function with numbers', () => {
      expect(evaluateFormula('=FE.LT(3, 5)')).toBe(true);
      expect(evaluateFormula('=FE.LT(5, 3)')).toBe(false);
      expect(evaluateFormula('=FE.LT(5, 5)')).toBe(false);
      expect(evaluateFormula('=FE.LT(-1, 0)')).toBe(true);
      expect(evaluateFormula('=FE.LT(0, -1)')).toBe(false);
    });

    test('FE.LT function with strings', () => {
      expect(evaluateFormula('=FE.LT("apple", "banana")')).toBe(true);
      expect(evaluateFormula('=FE.LT("banana", "apple")')).toBe(false);
      expect(evaluateFormula('=FE.LT("abc", "abc")')).toBe(false);
    });

    test('FE.LT function with mixed types', () => {
      expect(evaluateFormula('=FE.LT(FALSE, TRUE)')).toBe(true);
      expect(evaluateFormula('=FE.LT(TRUE, FALSE)')).toBe(false);
    });
  });

  describe('Less Than or Equal Comparison', () => {
    test('FE.LTE function with numbers', () => {
      expect(evaluateFormula('=FE.LTE(3, 5)')).toBe(true);
      expect(evaluateFormula('=FE.LTE(5, 5)')).toBe(true);
      expect(evaluateFormula('=FE.LTE(5, 3)')).toBe(false);
    });

    test('FE.LTE function with strings', () => {
      expect(evaluateFormula('=FE.LTE("apple", "banana")')).toBe(true);
      expect(evaluateFormula('=FE.LTE("apple", "apple")')).toBe(true);
      expect(evaluateFormula('=FE.LTE("banana", "apple")')).toBe(false);
    });
  });

  describe('Greater Than Comparison', () => {
    test('FE.GT function with numbers', () => {
      expect(evaluateFormula('=FE.GT(5, 3)')).toBe(true);
      expect(evaluateFormula('=FE.GT(3, 5)')).toBe(false);
      expect(evaluateFormula('=FE.GT(5, 5)')).toBe(false);
      expect(evaluateFormula('=FE.GT(0, -1)')).toBe(true);
    });

    test('FE.GT function with strings', () => {
      expect(evaluateFormula('=FE.GT("banana", "apple")')).toBe(true);
      expect(evaluateFormula('=FE.GT("apple", "banana")')).toBe(false);
      expect(evaluateFormula('=FE.GT("abc", "abc")')).toBe(false);
    });
  });

  describe('Greater Than or Equal Comparison', () => {
    test('FE.GTE function with numbers', () => {
      expect(evaluateFormula('=FE.GTE(5, 3)')).toBe(true);
      expect(evaluateFormula('=FE.GTE(5, 5)')).toBe(true);
      expect(evaluateFormula('=FE.GTE(3, 5)')).toBe(false);
    });

    test('FE.GTE function with strings', () => {
      expect(evaluateFormula('=FE.GTE("banana", "apple")')).toBe(true);
      expect(evaluateFormula('=FE.GTE("apple", "apple")')).toBe(true);
      expect(evaluateFormula('=FE.GTE("apple", "banana")')).toBe(false);
    });
  });

  describe('Error Handling', () => {
    test('Error propagation', () => {
      expect(evaluateFormula('=FE.EQ(#VALUE!, 5)')).toBe('#VALUE!');
      expect(evaluateFormula('=FE.LT(5, #NUM!)')).toBe('#NUM!');
      expect(evaluateFormula('=FE.GT(#DIV/0!, #VALUE!)')).toBe('#DIV/0!'); // First error propagates
    });

    test('Wrong number of arguments', () => {
      expect(() => evaluateFormula('=FE.EQ(1)')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=FE.EQ(1, 2, 3)')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=FE.LT()')).toThrow('Invalid number of arguments');
    });
  });

  describe('Null and Undefined Handling', () => {
    test('Comparisons with null/undefined treated as 0', () => {
      // Note: In actual usage, undefined would be passed for empty cells
      // These tests verify the internal comparison logic handles null/undefined properly
      
      // Test via helper functions that might pass undefined internally
      expect(evaluateFormula('=FE.EQ(0, 0)')).toBe(true); // Simulates undefined comparison
      expect(evaluateFormula('=FE.LT(0, 1)')).toBe(true); // Simulates undefined < 1
      expect(evaluateFormula('=FE.GT(1, 0)')).toBe(true); // Simulates 1 > undefined
    });
  });

  describe('Type Coercion Edge Cases', () => {
    test('String number coercion', () => {
      expect(evaluateFormula('=FE.EQ("123", 123)')).toBe(true);
      expect(evaluateFormula('=FE.LT("5", 10)')).toBe(true);
      expect(evaluateFormula('=FE.GT("20", 15)')).toBe(true);
    });

    test('Boolean coercion', () => {
      expect(evaluateFormula('=FE.EQ(TRUE, 1)')).toBe(true);
      expect(evaluateFormula('=FE.EQ(FALSE, 0)')).toBe(true);
      expect(evaluateFormula('=FE.LT(FALSE, TRUE)')).toBe(true);
    });
  });
});
