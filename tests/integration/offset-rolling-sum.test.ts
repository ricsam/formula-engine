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

describe("OFFSET Rolling Sum Test", () => {
  // Create a mock spreadsheet with data in column B
  const spreadsheet = new Map<string, CellValue>([
    ["0:1:0", 10],  // B1 = 10
    ["0:1:1", 20],  // B2 = 20
    ["0:1:2", 30],  // B3 = 30
    ["0:1:3", 40],  // B4 = 40
    ["0:1:4", 50],  // B5 = 50
    ["0:1:5", 60],  // B6 = 60
    ["0:1:6", 70],  // B7 = 70
    ["0:1:7", 80],  // B8 = 80
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

  test("OFFSET basic functionality", () => {
    // Test basic OFFSET - should return B3 (30)
    const result = evaluateFormula('=OFFSET(B2,1,0)');
    expect(result).toBe(30);
  });

  test("OFFSET with height parameter - rolling sum of 5 items", () => {
    // Test the rolling sum case: SUM(OFFSET(B2,0,0,5,1))
    // Should return B2:B6 range which is 20+30+40+50+60 = 200
    const result = evaluateFormula('=SUM(OFFSET(B2,0,0,5,1))');
    expect(result).toBe(200);
  });

  test("OFFSET with height parameter - different starting point", () => {
    // Test rolling sum starting from B3: SUM(OFFSET(B3,0,0,3,1))
    // Should return B3:B5 range which is 30+40+50 = 120
    const result = evaluateFormula('=SUM(OFFSET(B3,0,0,3,1))');
    expect(result).toBe(120);
  });

  test("OFFSET single cell reference", () => {
    // Test single cell OFFSET
    const result = evaluateFormula('=OFFSET(B1,2,0)');
    expect(result).toBe(30); // B3
  });

  test("OFFSET range creation", () => {
    // Test that OFFSET creates a proper range
    const result = evaluateFormula('=SUM(OFFSET(B1,0,0,3,1))');
    expect(result).toBe(60); // B1:B3 = 10+20+30
  });

  test("Debug: What does OFFSET return?", () => {
    // Let's see what OFFSET actually returns for various cases
    console.log("=== OFFSET Debug Tests ===");
    
    // Test single cell OFFSET
    const singleCell = evaluateFormula('=OFFSET(B1,1,0)');
    console.log(`OFFSET(B1,1,0) = ${singleCell} (expected: 20)`);
    
    // Test range OFFSET
    const rangeResult = evaluateFormula('=OFFSET(B2,0,0,5,1)');
    console.log(`OFFSET(B2,0,0,5,1) = ${JSON.stringify(rangeResult)}`);
    
    // Test if we can sum the result
    const sumResult = evaluateFormula('=SUM(OFFSET(B2,0,0,5,1))');
    console.log(`SUM(OFFSET(B2,0,0,5,1)) = ${sumResult} (expected: 200)`);
    
    // Let's also try a simpler case
    const simpleRange = evaluateFormula('=OFFSET(B1,0,0,2,1)');
    console.log(`OFFSET(B1,0,0,2,1) = ${JSON.stringify(simpleRange)}`);
    
    const simpleSumResult = evaluateFormula('=SUM(OFFSET(B1,0,0,2,1))');
    console.log(`SUM(OFFSET(B1,0,0,2,1)) = ${simpleSumResult} (expected: 30)`);
  });

  test("OFFSET edge cases and error conditions", () => {
    // Test OFFSET with out-of-bounds references
    const outOfBounds = evaluateFormula('=OFFSET(B1,-2,0)');
    console.log(`Out of bounds OFFSET: ${outOfBounds}`);
    
    // Test OFFSET with zero height/width
    try {
      const zeroHeight = evaluateFormula('=OFFSET(B1,0,0,0,1)');
      console.log(`Zero height OFFSET: ${zeroHeight}`);
    } catch (e) {
      console.log(`Zero height OFFSET error: ${e}`);
    }
    
    // Test OFFSET with negative height/width
    try {
      const negativeHeight = evaluateFormula('=OFFSET(B1,0,0,-1,1)');
      console.log(`Negative height OFFSET: ${negativeHeight}`);
    } catch (e) {
      console.log(`Negative height OFFSET error: ${e}`);
    }
  });

  test("OFFSET with different reference styles", () => {
    // Test with different ways of referencing the starting cell
    const directRef = evaluateFormula('=SUM(OFFSET(B2,0,0,3,1))');
    expect(directRef).toBe(90); // B2:B4 = 20+30+40
    
    // Test with string reference (if supported)
    // const stringRef = evaluateFormula('=SUM(OFFSET("B2",0,0,3,1))');
    // console.log(`String ref result: ${stringRef}`);
  });
});
