import { test, expect, describe } from "bun:test";
import { parseFormula } from "../../src/parser/parser";
import { Evaluator } from "../../src/evaluator/evaluator";
import { DependencyGraph } from "../../src/evaluator/dependency-graph";
import { ErrorHandler } from "../../src/evaluator/error-handler";
import { functionRegistry } from "../../src/functions";
import type {
  EvaluationContext,
} from "../../src/evaluator/evaluator";
import type {
  CellValue,
  SimpleCellAddress,
  SimpleCellRange,
} from "../../src/core/types";

describe("AVERAGEIF Function Integration Test", () => {
  // Create a mock spreadsheet
  const spreadsheet = new Map<string, CellValue>([
    ["0:0:0", 10],      // A1 = 10
    ["0:0:1", 20],      // A2 = 20
    ["0:0:2", 30],      // A3 = 30
    ["0:0:3", 40],      // A4 = 40
    ["0:1:0", "Apple"], // B1 = "Apple"
    ["0:1:1", "Orange"], // B2 = "Orange"
    ["0:1:2", "Apple"], // B3 = "Apple"
    ["0:1:3", "Banana"], // B4 = "Banana"
  ]);

  const functions = functionRegistry.getAllFunctions();
  const dependencyGraph = new DependencyGraph();
  const errorHandler = new ErrorHandler();
  const evaluator = new Evaluator(dependencyGraph, functions, errorHandler);

  const evaluateFormula = (formula: string): CellValue => {
    const ast = parseFormula(formula, 0);
    
    const context: EvaluationContext = {
      currentSheet: 0,
      namedExpressions: new Map(),
      getCellValue: (address: SimpleCellAddress) => {
        const key = `${address.sheet}:${address.col}:${address.row}`;
        return spreadsheet.get(key);
      },
      getRangeValues: (range: SimpleCellRange) => {
        const result: CellValue[][] = [];
        for (let row = range.start.row; row <= range.end.row; row++) {
          const rowData: CellValue[] = [];
          for (let col = range.start.col; col <= range.end.col; col++) {
            const key = `${range.start.sheet}:${col}:${row}`;
            rowData.push(spreadsheet.get(key) || undefined);
          }
          result.push(rowData);
        }
        return result;
      },
      getFunction: (name: string) => functionRegistry.get(name),
      errorHandler,
      evaluationStack: new Set(),
    };
    
    const result = evaluator.evaluate(ast, context);
    return result.value as CellValue;
  };

  test("AVERAGEIF with exact match", () => {
    const result = evaluateFormula('=AVERAGEIF(B1:B4,"Apple",A1:A4)');
    expect(result).toBe(20); // Average of 10 and 30
  });

  test("AVERAGEIF with no third argument", () => {
    const result = evaluateFormula('=AVERAGEIF(A1:A4,">25")');
    expect(result).toBe(35); // Average of 30 and 40
  });

  test("AVERAGEIF with no matches", () => {
    const result = evaluateFormula('=AVERAGEIF(B1:B4,"Grape",A1:A4)');
    expect(result).toBe("#DIV/0!"); // No matches
  });

  test("AVERAGEIF with comparison operator", () => {
    const result = evaluateFormula('=AVERAGEIF(A1:A4,">=20",A1:A4)');
    expect(result).toBe(30); // Average of 20, 30, 40
  });
});
