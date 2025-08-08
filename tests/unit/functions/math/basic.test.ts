import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { EvaluationContext } from '../../../../src/evaluator/evaluator';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('Basic Math Functions', () => {
  let evaluator: Evaluator;
  let errorHandler: ErrorHandler;
  let dependencyGraph: DependencyGraph;

  // Helper function to evaluate a formula
  const evaluateFormula = (formula: string): CellValue | CellValue[][] => {
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
    return result.value;
  };

  beforeEach(() => {
    errorHandler = new ErrorHandler();
    dependencyGraph = new DependencyGraph();
    
    // Get all functions from the registry
    const functions = functionRegistry.getAllFunctions();
    
    evaluator = new Evaluator(dependencyGraph, functions, errorHandler);
  });

  describe('Arithmetic Operators', () => {
    test('ADD function', () => {
      expect(evaluateFormula('=FE.ADD(2, 3)')).toBe(5);
      expect(evaluateFormula('=FE.ADD(-5, 10)')).toBe(5);
      expect(evaluateFormula('=FE.ADD(0, 0)')).toBe(0);
      expect(evaluateFormula('=FE.ADD(1.5, 2.5)')).toBe(4);
    });

    test('MINUS function', () => {
      expect(evaluateFormula('=FE.MINUS(10, 3)')).toBe(7);
      expect(evaluateFormula('=FE.MINUS(5, 10)')).toBe(-5);
      expect(evaluateFormula('=FE.MINUS(0, 0)')).toBe(0);
      expect(evaluateFormula('=FE.MINUS(3.5, 1.5)')).toBe(2);
    });

    test('MULTIPLY function', () => {
      expect(evaluateFormula('=FE.MULTIPLY(2, 3)')).toBe(6);
      expect(evaluateFormula('=FE.MULTIPLY(-5, 4)')).toBe(-20);
      expect(evaluateFormula('=FE.MULTIPLY(0, 100)')).toBe(0);
      expect(evaluateFormula('=FE.MULTIPLY(2.5, 4)')).toBe(10);
    });

    test('DIVIDE function', () => {
      expect(evaluateFormula('=FE.DIVIDE(10, 2)')).toBe(5);
      expect(evaluateFormula('=FE.DIVIDE(15, 3)')).toBe(5);
      expect(evaluateFormula('=FE.DIVIDE(-20, 4)')).toBe(-5);
      expect(evaluateFormula('=FE.DIVIDE(0, 5)')).toBe(0);
      expect(evaluateFormula('=FE.DIVIDE(7.5, 2.5)')).toBe(3);
    });

    test('DIVIDE by zero returns error', () => {
      expect(evaluateFormula('=FE.DIVIDE(10, 0)')).toBe('#DIV/0!');
      expect(evaluateFormula('=FE.DIVIDE(-5, 0)')).toBe('#DIV/0!');
    });

    test('POW function', () => {
      expect(evaluateFormula('=FE.POW(2, 3)')).toBe(8);
      expect(evaluateFormula('=FE.POW(5, 2)')).toBe(25);
      expect(evaluateFormula('=FE.POW(10, 0)')).toBe(1);
      expect(evaluateFormula('=FE.POW(4, 0.5)')).toBe(2);
      expect(evaluateFormula('=FE.POW(2, -2)')).toBe(0.25);
    });
  });

  describe('Unary Operators', () => {
    test('UMINUS function', () => {
      expect(evaluateFormula('=FE.UMINUS(5)')).toBe(-5);
      expect(evaluateFormula('=FE.UMINUS(-10)')).toBe(10);
      expect(evaluateFormula('=FE.UMINUS(0)')).toBe(0);
      expect(evaluateFormula('=FE.UMINUS(3.14)')).toBe(-3.14);
    });

    test('UPLUS function', () => {
      expect(evaluateFormula('=FE.UPLUS(5)')).toBe(5);
      expect(evaluateFormula('=FE.UPLUS(-10)')).toBe(-10);
      expect(evaluateFormula('=FE.UPLUS(0)')).toBe(0);
      expect(evaluateFormula('=FE.UPLUS(3.14)')).toBe(3.14);
    });

    test('UNARY_PERCENT function', () => {
      expect(evaluateFormula('=FE.UNARY_PERCENT(50)')).toBe(0.5);
      expect(evaluateFormula('=FE.UNARY_PERCENT(100)')).toBe(1);
      expect(evaluateFormula('=FE.UNARY_PERCENT(25)')).toBe(0.25);
      expect(evaluateFormula('=FE.UNARY_PERCENT(0)')).toBe(0);
      expect(evaluateFormula('=FE.UNARY_PERCENT(150)')).toBe(1.5);
    });
  });

  describe('Type Coercion', () => {
    test('String to number coercion', () => {
      expect(evaluateFormula('=FE.ADD("5", "3")')).toBe(8);
      expect(evaluateFormula('=FE.MULTIPLY("2", "4")')).toBe(8);
    });

    test('Boolean to number coercion', () => {
      expect(evaluateFormula('=FE.ADD(TRUE, 5)')).toBe(6);
      expect(evaluateFormula('=FE.ADD(FALSE, 10)')).toBe(10);
      expect(evaluateFormula('=FE.MULTIPLY(TRUE, 5)')).toBe(5);
      expect(evaluateFormula('=FE.MULTIPLY(FALSE, 5)')).toBe(0);
    });

    test('Invalid string returns error', () => {
      expect(evaluateFormula('=FE.ADD("abc", 5)')).toBe('#VALUE!');
      expect(evaluateFormula('=FE.MULTIPLY("xyz", 2)')).toBe('#VALUE!');
    });
  });

  describe('Error Propagation', () => {
    test('Error in arguments propagates', () => {
      expect(evaluateFormula('=FE.ADD(FE.DIVIDE(1, 0), 5)')).toBe('#DIV/0!');
      expect(evaluateFormula('=FE.MULTIPLY(#VALUE!, 2)')).toBe('#VALUE!');
    });

    test('Wrong number of arguments', () => {
      // Parser validates argument count and throws ParseError
      // This is expected behavior for built-in operators
      expect(() => evaluateFormula('=FE.ADD(1)')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=FE.ADD(1, 2, 3)')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=FE.UMINUS()')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=FE.UMINUS(1, 2)')).toThrow('Invalid number of arguments');
    });
  });
});