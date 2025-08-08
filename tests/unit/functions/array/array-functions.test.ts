import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('Array functions: SORT/UNIQUE/SEQUENCE', () => {
  let evaluator: Evaluator;
  let errorHandler: ErrorHandler;
  let dependencyGraph: DependencyGraph;

  const evalFormula = (formula: string): any => {
    const ast = Parser.parse(formula, 0);
    const ctx = {
      currentSheet: 0,
      namedExpressions: new Map(),
      getCellValue: (_addr: SimpleCellAddress) => undefined,
      getRangeValues: (_range: SimpleCellRange) => [[]],
      getFunction: (name: string) => functionRegistry.get(name),
      errorHandler,
      evaluationStack: new Set<string>(),
    };
    return evaluator.evaluate(ast, ctx).value;
  };

  beforeEach(() => {
    errorHandler = new ErrorHandler();
    dependencyGraph = new DependencyGraph();
    evaluator = new Evaluator(dependencyGraph, functionRegistry.getAllFunctions(), errorHandler);
  });

  test('SEQUENCE basic usage', () => {
    expect(evalFormula('=SEQUENCE(3)')).toEqual([[1],[2],[3]]);
    expect(evalFormula('=SEQUENCE(2,3,10,2)')).toEqual([[10,12,14],[16,18,20]]);
  });

  test('SORT rows ascending/descending', () => {
    // Sort by second column asc
    expect(evalFormula('=SORT({3,2;1,5;2,4}, 2, 1)')).toEqual([[3,2],[2,4],[1,5]]);
    // Sort by first column desc
    expect(evalFormula('=SORT({3,2;1,5;2,4}, 1, -1)')).toEqual([[3,2],[2,4],[1,5]]);
  });

  test('SORT by columns (by_col=TRUE)', () => {
    // Sort columns by first row desc
    expect(evalFormula('=SORT({3,1,2;9,8,7}, 1, -1, TRUE)')).toEqual([[3,2,1],[9,7,8]]);
  });

  test('UNIQUE rows and exactly_once', () => {
    expect(evalFormula('=UNIQUE({1,2;1,2;2,3})')).toEqual([[1,2],[2,3]]);
    expect(evalFormula('=UNIQUE({1,2;1,2;2,3}, FALSE, TRUE)')).toEqual([[2,3]]);
  });

  test('UNIQUE by columns', () => {
    expect(evalFormula('=UNIQUE({1,1,2;2,2,3}, TRUE)')).toEqual([[1,2],[2,3]]);
  });
});


