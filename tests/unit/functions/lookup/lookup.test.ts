import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { EvaluationContext } from '../../../../src/evaluator/evaluator';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('Lookup Functions', () => {
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
    return result.value;
  };

  beforeEach(() => {
    errorHandler = new ErrorHandler();
    dependencyGraph = new DependencyGraph();
    const functions = functionRegistry.getAllFunctions();
    evaluator = new Evaluator(dependencyGraph, functions, errorHandler);
  });

  describe('INDEX function', () => {
    test('INDEX with 2D array', () => {
      // INDEX({1,2,3;4,5,6;7,8,9}, 2, 3) = 6
      expect(evaluateFormula('=INDEX({1,2,3;4,5,6;7,8,9}, 2, 3)')).toBe(6);
    });

    test('INDEX with 1D array (row)', () => {
      // INDEX({10,20,30}, 1, 2) = 20 (it's a 1x3 array)
      expect(evaluateFormula('=INDEX({10,20,30}, 1, 2)')).toBe(20);
    });

    test('INDEX with 1D array (column)', () => {
      // INDEX({10;20;30}, 2, 1) = 20 (it's a 3x1 array)
      expect(evaluateFormula('=INDEX({10;20;30}, 2, 1)')).toBe(20);
    });

    test('INDEX with single value', () => {
      // INDEX(42, 1, 1) = 42
      expect(evaluateFormula('=INDEX(42, 1, 1)')).toBe(42);
    });

    test('INDEX out of bounds', () => {
      expect(evaluateFormula('=INDEX({1,2,3}, 1, 4)')).toBe('#REF!');
      expect(evaluateFormula('=INDEX({1,2;3,4}, 3, 1)')).toBe('#REF!');
      expect(evaluateFormula('=INDEX({1,2;3,4}, 1, 3)')).toBe('#REF!');
    });

    test('INDEX with invalid arguments', () => {
      expect(evaluateFormula('=INDEX({1,2,3}, "text")')).toBe('#VALUE!');
      expect(evaluateFormula('=INDEX({1,2,3}, 1, "text")')).toBe('#VALUE!');
    });
  });

  describe('ROW/COLUMN/ROWS/COLUMNS/CHOOSE', () => {
    test('ROW no-arg returns current row (1-based)', () => {
      // Simulate current cell via evaluator context
      const ast = Parser.parse('=ROW()', 0);
      const context: EvaluationContext = {
        currentSheet: 0,
        currentCell: { sheet: 0, col: 0, row: 4 },
        namedExpressions: new Map(),
        getCellValue: () => undefined,
        getRangeValues: () => [[]],
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler,
        evaluationStack: new Set(),
      };
      const result = evaluator.evaluate(ast, context).value;
      expect(result).toBe(5);
    });

    test('COLUMN no-arg returns current column (1-based)', () => {
      const ast = Parser.parse('=COLUMN()', 0);
      const context: EvaluationContext = {
        currentSheet: 0,
        currentCell: { sheet: 0, col: 2, row: 2 },
        namedExpressions: new Map(),
        getCellValue: () => undefined,
        getRangeValues: () => [[]],
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler,
        evaluationStack: new Set(),
      };
      const result = evaluator.evaluate(ast, context).value;
      expect(result).toBe(3);
    });

    test('ROWS returns number of rows in array', () => {
      expect(evaluateFormula('=ROWS({1;2;3})')).toBe(3);
    });

    test('COLUMNS returns number of columns in array', () => {
      expect(evaluateFormula('=COLUMNS({1,2,3})')).toBe(3);
    });

    test('CHOOSE returns value by index', () => {
      expect(evaluateFormula('=CHOOSE(2, "a", "b", "c")')).toBe('b');
    });
  });

  describe('INDIRECT function', () => {
    test('should dereference single cell address', () => {
      // Mock getCellValue via context by setting a tiny sheet map
      const ast = Parser.parse('=INDIRECT("A1")', 0);
      const context: EvaluationContext = {
        currentSheet: 0,
        currentCell: { sheet: 0, col: 0, row: 0 },
        namedExpressions: new Map(),
        getCellValue: (addr) => (addr.col === 0 && addr.row === 0 ? 42 : undefined),
        getRangeValues: () => [[]],
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler,
        evaluationStack: new Set(),
      };
      const result = evaluator.evaluate(ast, context).value;
      expect(result).toBe(42);
    });

    test('should dereference range address and return 2D array', () => {
      const ast = Parser.parse('=INDIRECT("A1:B2")', 0);
      const context: EvaluationContext = {
        currentSheet: 0,
        currentCell: { sheet: 0, col: 0, row: 0 },
        namedExpressions: new Map(),
        getCellValue: () => undefined,
        getRangeValues: (range) => {
          return [[1, 2], [3, 4]];
        },
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler,
        evaluationStack: new Set(),
      };
      const result = evaluator.evaluate(ast, context).value as unknown as number[][];
      expect(result).toEqual([[1, 2], [3, 4]]);
    });

    test('invalid reference returns #REF!', () => {
      expect(evaluateFormula('=INDIRECT("not_an_addr")')).toBe('#REF!');
    });
  });

  describe('OFFSET function', () => {
    test('OFFSET with text reference single cell', () => {
      const ast = Parser.parse('=OFFSET("A1", 1, 2)', 0);
      const context: EvaluationContext = {
        currentSheet: 0,
        currentCell: { sheet: 0, col: 0, row: 0 },
        namedExpressions: new Map(),
        getCellValue: (addr) => (addr.col === 2 && addr.row === 1 ? 99 : undefined),
        getRangeValues: () => [[99]],
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler,
        evaluationStack: new Set(),
      };
      const result = evaluator.evaluate(ast, context).value;
      expect(result).toBe(99);
    });

    test('OFFSET with resize returns range', () => {
      const ast = Parser.parse('=OFFSET("A1", 0, 0, 2, 2)', 0);
      const context: EvaluationContext = {
        currentSheet: 0,
        currentCell: { sheet: 0, col: 0, row: 0 },
        namedExpressions: new Map(),
        getCellValue: () => undefined,
        getRangeValues: () => [[1,2],[3,4]],
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler,
        evaluationStack: new Set(),
      };
      const result = evaluator.evaluate(ast, context).value as unknown as number[][];
      expect(result).toEqual([[1,2],[3,4]]);
    });
  });

  describe('MATCH function', () => {
    test('MATCH exact match', () => {
      // MATCH(5, {1,3,5,7,9}, 0) = 3
      expect(evaluateFormula('=MATCH(5, {1,3,5,7,9}, 0)')).toBe(3);
      expect(evaluateFormula('=MATCH("b", {"a","b","c"}, 0)')).toBe(2);
    });

    test('MATCH not found', () => {
      expect(evaluateFormula('=MATCH(4, {1,3,5,7,9}, 0)')).toBe('#N/A');
    });

    test('MATCH less than or equal (default)', () => {
      // MATCH(6, {1,3,5,7,9}, 1) = 3 (5 is largest <= 6)
      expect(evaluateFormula('=MATCH(6, {1,3,5,7,9}, 1)')).toBe(3);
      expect(evaluateFormula('=MATCH(6, {1,3,5,7,9})')).toBe(3); // Default is 1
    });

    test('MATCH greater than or equal', () => {
      // MATCH(6, {9,7,5,3,1}, -1) = 1 (9 is first value >= 6 in descending array)
      expect(evaluateFormula('=MATCH(6, {9,7,5,3,1}, -1)')).toBe(1);
      // MATCH(0, {9,7,5,3,1}, -1) = 1 (all values are >= 0)
      expect(evaluateFormula('=MATCH(0, {9,7,5,3,1}, -1)')).toBe(1);
      // MATCH(10, {9,7,5,3,1}, -1) = #N/A (no values >= 10)
      expect(evaluateFormula('=MATCH(10, {9,7,5,3,1}, -1)')).toBe('#N/A');
    });
  });

  describe('VLOOKUP function', () => {
    test('VLOOKUP exact match', () => {
      // VLOOKUP(2, {1,"a";2,"b";3,"c"}, 2, FALSE) = "b"
      expect(evaluateFormula('=VLOOKUP(2, {1,"a";2,"b";3,"c"}, 2, FALSE)')).toBe("b");
    });

    test('VLOOKUP approximate match', () => {
      // VLOOKUP(2.5, {1,"a";2,"b";3,"c"}, 2, TRUE) = "b"
      expect(evaluateFormula('=VLOOKUP(2.5, {1,"a";2,"b";3,"c"}, 2, TRUE)')).toBe("b");
      expect(evaluateFormula('=VLOOKUP(2.5, {1,"a";2,"b";3,"c"}, 2)')).toBe("b"); // Default TRUE
    });

    test('VLOOKUP not found', () => {
      expect(evaluateFormula('=VLOOKUP(4, {1,"a";2,"b";3,"c"}, 2, FALSE)')).toBe('#N/A');
      expect(evaluateFormula('=VLOOKUP(0, {1,"a";2,"b";3,"c"}, 2, TRUE)')).toBe('#N/A');
    });

    test('VLOOKUP invalid column', () => {
      expect(evaluateFormula('=VLOOKUP(2, {1,"a";2,"b";3,"c"}, 3, FALSE)')).toBe('#REF!');
      expect(evaluateFormula('=VLOOKUP(2, {1,"a";2,"b";3,"c"}, 0, FALSE)')).toBe('#REF!');
    });
  });

  describe('HLOOKUP function', () => {
    test('HLOOKUP exact match', () => {
      // HLOOKUP("b", {"a","b","c";1,2,3}, 2, FALSE) = 2
      expect(evaluateFormula('=HLOOKUP("b", {"a","b","c";1,2,3}, 2, FALSE)')).toBe(2);
    });

    test('HLOOKUP approximate match', () => {
      // HLOOKUP(2.5, {1,2,3;"a","b","c"}, 2, TRUE) = "b"
      expect(evaluateFormula('=HLOOKUP(2.5, {1,2,3;"a","b","c"}, 2, TRUE)')).toBe("b");
    });

    test('HLOOKUP not found', () => {
      expect(evaluateFormula('=HLOOKUP("d", {"a","b","c";1,2,3}, 2, FALSE)')).toBe('#N/A');
    });
  });

  describe('XLOOKUP function', () => {
    test('XLOOKUP basic usage', () => {
      // XLOOKUP(2, {1,2,3}, {"a","b","c"}) = "b"
      expect(evaluateFormula('=XLOOKUP(2, {1,2,3}, {"a","b","c"})')).toBe("b");
    });

    test('XLOOKUP with if_not_found', () => {
      // XLOOKUP(4, {1,2,3}, {"a","b","c"}, "not found") = "not found"
      expect(evaluateFormula('=XLOOKUP(4, {1,2,3}, {"a","b","c"}, "not found")')).toBe("not found");
    });

    test('XLOOKUP exact match or next smallest', () => {
      // XLOOKUP(2.5, {1,2,3}, {"a","b","c"}, "not found", -1) = "b"
      expect(evaluateFormula('=XLOOKUP(2.5, {1,2,3}, {"a","b","c"}, "not found", -1)')).toBe("b");
    });

    test('XLOOKUP exact match or next largest', () => {
      // XLOOKUP(2.5, {1,2,3}, {"a","b","c"}, "not found", 1) = "c"
      expect(evaluateFormula('=XLOOKUP(2.5, {1,2,3}, {"a","b","c"}, "not found", 1)')).toBe("c");
    });

    test('XLOOKUP wildcard match', () => {
      // XLOOKUP("b*", {"apple","banana","cherry"}, {1,2,3}, 0, 2) = 2
      expect(evaluateFormula('=XLOOKUP("b*", {"apple","banana","cherry"}, {1,2,3}, 0, 2)')).toBe(2);
      expect(evaluateFormula('=XLOOKUP("?pple", {"apple","banana","cherry"}, {1,2,3}, 0, 2)')).toBe(1);
    });

    test('XLOOKUP reverse search', () => {
      // XLOOKUP(2, {1,2,2,3}, {"a","b","c","d"}, "not found", 0, -1) = "c" (last match)
      expect(evaluateFormula('=XLOOKUP(2, {1,2,2,3}, {"a","b","c","d"}, "not found", 0, -1)')).toBe("c");
    });

    test('XLOOKUP array size mismatch', () => {
      expect(evaluateFormula('=XLOOKUP(2, {1,2,3}, {"a","b"})')).toBe('#VALUE!');
    });
  });
});