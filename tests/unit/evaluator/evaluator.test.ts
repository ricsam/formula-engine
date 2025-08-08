import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../src/evaluator/evaluator';
import { DependencyGraph } from '../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../src/evaluator/error-handler';
import type { 
  EvaluationContext, 
  FunctionDefinition,
  EvaluationResult 
} from '../../../src/evaluator/evaluator';
import type { 
  CellValue, 
  SimpleCellAddress, 
  SimpleCellRange,
  NamedExpression
} from '../../../src/core/types';
import type { ASTNode } from '../../../src/parser/ast';
import {
  createTestValueNode,
  createTestReferenceNode,
  createTestRangeNode,
  createTestFunctionNode,
  createTestBinaryOpNode,
  createTestUnaryOpNode,
  createTestArrayNode,
  createTestNamedExpressionNode,
  createTestErrorNode
} from './test-helpers';

describe('Evaluator', () => {
  let evaluator: Evaluator;
  let dependencyGraph: DependencyGraph;
  let errorHandler: ErrorHandler;
  let functions: Map<string, FunctionDefinition>;
  let context: EvaluationContext;

  beforeEach(() => {
    dependencyGraph = new DependencyGraph();
    errorHandler = new ErrorHandler();
    functions = new Map();
    
    // Add some test functions
    functions.set('SUM', {
      name: 'SUM',
      minArgs: 1,
      evaluate: ({ argValues: args }) => {
        let sum = 0;
        for (const arg of args) {
          if (Array.isArray(arg)) {
            // Handle arrays
            const flat = arg.flat();
            for (const val of flat) {
              if (typeof val === 'number') sum += val;
            }
          } else if (typeof arg === 'number') {
            sum += arg;
          }
        }
        return sum;
      }
    });

    functions.set('UPPER', {
      name: 'UPPER',
      minArgs: 1,
      maxArgs: 1,
      evaluate: ({ argValues: args }) => {
        const text = args[0];
        if (typeof text !== 'string') return '#VALUE!';
        return text.toUpperCase();
      }
    });

    evaluator = new Evaluator(dependencyGraph, functions, errorHandler);

    // Create a basic context
    context = {
      currentSheet: 0,
      currentCell: { sheet: 0, col: 0, row: 0 },
      namedExpressions: new Map(),
      getCellValue: (address: SimpleCellAddress) => {
        // Simple test data
        if (address.col === 0 && address.row === 0) return 10;
        if (address.col === 1 && address.row === 0) return 20;
        if (address.col === 2 && address.row === 0) return 30;
        if (address.col === 0 && address.row === 1) return 'Hello';
        if (address.col === 1 && address.row === 1) return 'World';
        return undefined;
      },
      getRangeValues: (range: SimpleCellRange) => {
        const result: CellValue[][] = [];
        for (let row = range.start.row; row <= range.end.row; row++) {
          const rowData: CellValue[] = [];
          for (let col = range.start.col; col <= range.end.col; col++) {
            rowData.push(context.getCellValue({ 
              sheet: range.start.sheet, 
              col, 
              row 
            }));
          }
          result.push(rowData);
        }
        return result;
      },
      getFunction: (name: string) => functions.get(name),
      errorHandler,
      evaluationStack: new Set()
    };
  });

  describe('Value nodes', () => {
    test('should evaluate number node', () => {
      const node = createTestValueNode(42, 'number');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(42);
      expect(result.isArrayResult).toBe(false);
      expect(result.dependencies.size).toBe(0);
    });

    test('should evaluate string node', () => {
      const node = createTestValueNode('test', 'string');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('test');
    });

    test('should evaluate boolean node', () => {
      const node = createTestValueNode(true, 'boolean');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(true);
    });

    test('should evaluate error node', () => {
      const node = createTestErrorNode('#DIV/0!');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#DIV/0!');
    });
  });

  describe('Reference nodes', () => {
    test('should evaluate cell reference', () => {
      const node = createTestReferenceNode(
        { sheet: 0, col: 0, row: 0 },
        { col: false, row: false }
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(10);
      expect(result.dependencies.has('0:0:0')).toBe(true);
    });

    test('should detect circular reference', () => {
      context.evaluationStack.add('0:0:0');
      
      const node = createTestReferenceNode(
        { sheet: 0, col: 0, row: 0 },
        { col: false, row: false }
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#CYCLE!');
    });
  });

  describe('Range nodes', () => {
    test('should evaluate range', () => {
      const node = createTestRangeNode({
        start: { sheet: 0, col: 0, row: 0 },
        end: { sheet: 0, col: 2, row: 0 }
      });

      const result = evaluator.evaluate(node, context);
      expect(result.value as unknown as number[][]).toEqual([[10, 20, 30]]);
      expect(result.isArrayResult).toBe(true);
      expect(result.arrayDimensions).toEqual({ rows: 1, cols: 3 });
      expect(result.dependencies.has('0:0:0:2:0')).toBe(true);
    });
  });

  describe('Function nodes', () => {
    test('should evaluate function call', () => {
      const node = createTestFunctionNode('SUM', [
        createTestValueNode(10, 'number'),
        createTestValueNode(20, 'number'),
        createTestValueNode(30, 'number')
      ]);

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(60);
    });

    test('should evaluate function with range argument', () => {
      const node = createTestFunctionNode('SUM', [
        createTestRangeNode({
          start: { sheet: 0, col: 0, row: 0 },
          end: { sheet: 0, col: 2, row: 0 }
        })
      ]);

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(60);
    });

    test('should handle unknown function', () => {
      const node = createTestFunctionNode('UNKNOWN_FUNC');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#NAME?');
    });

    test('should validate argument count', () => {
      const node = createTestFunctionNode('UPPER', [
        createTestValueNode('test', 'string'),
        createTestValueNode('extra', 'string')
      ]);

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#VALUE!');
    });

    test('should handle function errors', () => {
      const node = createTestFunctionNode('UPPER', [
        createTestValueNode(123, 'number')
      ]);

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#VALUE!');
    });
  });

  describe('Binary operations', () => {
    test('should evaluate addition', () => {
      const node = createTestBinaryOpNode(
        '+',
        createTestValueNode(10, 'number'),
        createTestValueNode(20, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(30);
    });

    test('should evaluate subtraction', () => {
      const node = createTestBinaryOpNode(
        '-',
        createTestValueNode(30, 'number'),
        createTestValueNode(10, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(20);
    });

    test('should evaluate multiplication', () => {
      const node = createTestBinaryOpNode(
        '*',
        createTestValueNode(5, 'number'),
        createTestValueNode(6, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(30);
    });

    test('should evaluate division', () => {
      const node = createTestBinaryOpNode(
        '/',
        createTestValueNode(20, 'number'),
        createTestValueNode(4, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(5);
    });

    test('should handle division by zero', () => {
      const node = createTestBinaryOpNode(
        '/',
        createTestValueNode(10, 'number'),
        createTestValueNode(0, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#DIV/0!');
    });

    test('should evaluate exponentiation', () => {
      const node = createTestBinaryOpNode(
        '^',
        createTestValueNode(2, 'number'),
        createTestValueNode(3, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(8);
    });

    test('should evaluate concatenation', () => {
      const node = createTestBinaryOpNode(
        '&',
        createTestValueNode('Hello', 'string'),
        createTestValueNode('World', 'string')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('HelloWorld');
    });

    test('should evaluate equality', () => {
      const node = createTestBinaryOpNode(
        '=',
        createTestValueNode(10, 'number'),
        createTestValueNode(10, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(true);
    });

    test('should evaluate inequality', () => {
      const node = createTestBinaryOpNode(
        '<>',
        createTestValueNode(10, 'number'),
        createTestValueNode(20, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(true);
    });

    test('should evaluate comparisons', () => {
      expect(evaluator.evaluate(createTestBinaryOpNode(
        '<',
        createTestValueNode(5, 'number'),
        createTestValueNode(10, 'number')
      ), context).value).toBe(true);

      expect(evaluator.evaluate(createTestBinaryOpNode(
        '<=',
        createTestValueNode(10, 'number'),
        createTestValueNode(10, 'number')
      ), context).value).toBe(true);

      expect(evaluator.evaluate(createTestBinaryOpNode(
        '>',
        createTestValueNode(20, 'number'),
        createTestValueNode(10, 'number')
      ), context).value).toBe(true);

      expect(evaluator.evaluate(createTestBinaryOpNode(
        '>=',
        createTestValueNode(10, 'number'),
        createTestValueNode(10, 'number')
      ), context).value).toBe(true);
    });

    test('should handle array operations', () => {
      const node = createTestBinaryOpNode(
        '+',
        createTestArrayNode([
          [
            createTestValueNode(1, 'number'),
            createTestValueNode(2, 'number')
          ]
        ]),
        createTestValueNode(10, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value as unknown as number[][]).toEqual([[11, 12]]);
    });
  });

  describe('Unary operations', () => {
    test('should evaluate negation', () => {
      const node = createTestUnaryOpNode(
        '-',
        createTestValueNode(42, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(-42);
    });

    test('should evaluate unary plus', () => {
      const node = createTestUnaryOpNode(
        '+',
        createTestValueNode(42, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(42);
    });

    test('should evaluate percent', () => {
      const node = createTestUnaryOpNode(
        '%',
        createTestValueNode(50, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(0.5);
    });
  });

  describe('Array nodes', () => {
    test('should evaluate array literal', () => {
      const node = createTestArrayNode([
        [createTestValueNode(1, 'number'), createTestValueNode(2, 'number')],
        [createTestValueNode(3, 'number'), createTestValueNode(4, 'number')]
      ]);

      const result = evaluator.evaluate(node, context);
      expect(result.value as unknown as number[][]).toEqual([[1, 2], [3, 4]]);
      expect(result.isArrayResult).toBe(true);
    });
  });

  describe('Named expressions', () => {
    test('should evaluate named expression', () => {
      context.namedExpressions.set('TaxRate', {
        name: 'TaxRate',
        expression: '0.08'
      });

      const node = createTestNamedExpressionNode('TaxRate');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe(0.08);
    });

    test('should handle missing named expression', () => {
      const node = createTestNamedExpressionNode('UnknownName');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#NAME?');
    });

    test('should detect circular reference in named expression', () => {
      context.namedExpressions.set('Circular', {
        name: 'Circular',
        expression: '=Circular'
      });
      context.evaluationStack.add('name:Circular');

      const node = createTestNamedExpressionNode('Circular');

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#CYCLE!');
    });
  });

  describe('Type coercion', () => {
    test('should coerce values to numbers', () => {
      expect(evaluator.coerceToNumber(42)).toBe(42);
      expect(evaluator.coerceToNumber('42')).toBe(42);
      expect(evaluator.coerceToNumber(true)).toBe(1);
      expect(evaluator.coerceToNumber(false)).toBe(0);
      expect(evaluator.coerceToNumber(undefined)).toBe(0);
      expect(evaluator.coerceToNumber('')).toBe(0);
      expect(evaluator.coerceToNumber('abc')).toBe('#VALUE!');
    });

    test('should coerce values to strings', () => {
      expect(evaluator.coerceToString('test')).toBe('test');
      expect(evaluator.coerceToString(42)).toBe('42');
      expect(evaluator.coerceToString(true)).toBe('TRUE');
      expect(evaluator.coerceToString(false)).toBe('FALSE');
      expect(evaluator.coerceToString(undefined)).toBe('');
    });

    test('should coerce values to booleans', () => {
      expect(evaluator.coerceToBoolean(true)).toBe(true);
      expect(evaluator.coerceToBoolean(false)).toBe(false);
      expect(evaluator.coerceToBoolean(1)).toBe(true);
      expect(evaluator.coerceToBoolean(0)).toBe(false);
      expect(evaluator.coerceToBoolean('TRUE')).toBe(true);
      expect(evaluator.coerceToBoolean('FALSE')).toBe(false);
      expect(evaluator.coerceToBoolean('text')).toBe(true);
      expect(evaluator.coerceToBoolean('')).toBe(false);
      expect(evaluator.coerceToBoolean(undefined)).toBe(false);
    });
  });

  describe('Error propagation', () => {
    test('should propagate errors through operations', () => {
      const node = createTestBinaryOpNode(
        '+',
        createTestErrorNode('#VALUE!'),
        createTestValueNode(10, 'number')
      );

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#VALUE!');
    });

    test('should propagate errors through functions', () => {
      const node = createTestFunctionNode('SUM', [
        createTestValueNode(10, 'number'),
        createTestErrorNode('#REF!'),
        createTestValueNode(20, 'number')
      ]);

      const result = evaluator.evaluate(node, context);
      expect(result.value).toBe('#REF!');
    });
  });

  describe('Cache management', () => {
    test('should clear evaluation cache', () => {
      // Evaluate something to potentially populate cache
      const node = createTestValueNode(42, 'number');

      evaluator.evaluate(node, context);
      
      // Clear cache should not throw
      expect(() => evaluator.clearCache()).not.toThrow();
    });
  });
});