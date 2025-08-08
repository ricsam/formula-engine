import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('IFERROR function', () => {
  let evaluator: Evaluator;
  let errorHandler: ErrorHandler;
  let dependencyGraph: DependencyGraph;

  const evalFormula = (formula: string): CellValue => {
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

  test('returns first arg if not error', () => {
    expect(evalFormula('=IFERROR(5/2, 0)')).toBe(2.5);
  });

  test('returns fallback when first arg is error', () => {
    expect(evalFormula('=IFERROR(5/0, 0)')).toBe(0);
    expect(evalFormula('=IFERROR(#REF!, "fallback")')).toBe('fallback');
  });
});


