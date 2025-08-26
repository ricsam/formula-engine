import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("MIN function", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  test("basic scalar arguments", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([["A1", "=MIN(6, 2, 8, 1, 9)"]])
    );

    expect(cell("A1")).toBe(1); // Minimum of 6, 2, 8, 1, 9 is 1
  });

  test("with cell references", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 30],
        ["A2", 10],
        ["A3", 20],
        ["B1", "=MIN(A1:A3)"],
      ])
    );

    expect(cell("B1")).toBe(10); // Minimum of 30, 10, 20 is 10
  });

  test("with structured references", () => {
    // Create a table with data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        // Table headers
        ["A1", "Name"],
        ["B1", "Value"],
        ["C1", "Count"],
        // Table data
        ["A2", "Item1"],
        ["B2", 100],
        ["C2", 5],
        ["A3", "Item2"],
        ["B3", 200],
        ["C3", 3],
        ["A4", "Item3"],
        ["B4", 150],
        ["C4", 7],
        ["D1", "=MIN(DataTable[Value])"], // Min the Value column
        ["D2", "=MIN(DataTable[Count])"], // Min the Count column
      ])
    );

    // Define the table
    engine.addTable({
      tableName: "DataTable",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 3 }, // 3 data rows
      numCols: 3, // 3 columns: Name, Value, Count
    });

    // ENGINE ISSUE: Structured references like DataTable[Value] not supported
    expect(cell("D1", true)).toBe(100); // Min of 100, 200, 150 is 100
    expect(cell("D2", true)).toBe(3); // Min of 5, 3, 7 is 3
  });

  test("with named expressions", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 30],
        ["A2", 10],
        ["A3", 20],
        ["B1", 15],
        ["B2", 5],
        ["C1", "=MIN(VALUES_A)"], // Min named range
        ["C2", "=MIN(VALUES_A, VALUES_B)"], // Min multiple named ranges
        ["C3", "=MIN(SINGLE_VALUE, 10)"], // Min named value with scalar
      ])
    );

    // Define named expressions
    engine.addNamedExpression({
      expression: "A1:A3",
      expressionName: "VALUES_A",
      sheetName,
    });
    engine.addNamedExpression({
      expression: "B1:B2",
      expressionName: "VALUES_B",
      sheetName,
    });
    engine.addNamedExpression({
      expression: "25",
      expressionName: "SINGLE_VALUE",
      sheetName,
    });

    // ENGINE ISSUE: Named expressions that evaluate to ranges not supported in function calls
    expect(cell("C1")).toBe(10); // Min of 30, 10, 20 is 10
    expect(cell("C2")).toBe(5); // Min of 30, 10, 20, 15, 5 is 5
    expect(cell("C3")).toBe(10); // Min of 25, 10 is 10
  });

  test("with cross-sheet references", () => {
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;

    // Set up data on Sheet1
    engine.setSheetContent(
      sheet1Name,
      new Map<string, SerializedCellValue>([
        ["A1", 100],
        ["A2", 200],
        ["A3", 300],
        ["B1", "=MIN(Sheet2!B1:B2)"], // Min range from Sheet2
        ["B2", "=MIN(A1:A3, Sheet2!B1:B2)"], // Min local and cross-sheet ranges
        ["B3", "=MIN(Sheet1!A1, Sheet2!B1, 25)"], // Mix of cross-sheet cells and scalar
      ])
    );

    // Set up data on Sheet2
    engine.setSheetContent(
      sheet2Name,
      new Map<string, SerializedCellValue>([
        ["B1", 50],
        ["B2", 75],
      ])
    );

    const cell = (sheetName: string, ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    // ENGINE ISSUE: Cross-sheet references like Sheet2!B1:B2 not supported
    expect(cell(sheet1Name, "B1")).toBe(50); // Min of 50, 75 is 50
    expect(cell(sheet1Name, "B2")).toBe(50); // Min of 100, 200, 300, 50, 75 is 50
    expect(cell(sheet1Name, "B3")).toBe(25); // Min of 100, 50, 25 is 25
  });

  test.skip("with 3D sheet references", () => {
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;
    const sheet3Name = engine.addSheet("Sheet3").name;

    // Set up same data on all sheets
    [sheet1Name, sheet2Name, sheet3Name].forEach((sheetName) => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
        ])
      );
    });

    // Create 3D reference formulas
    engine.setSheetContent(
      sheet1Name,
      new Map<string, SerializedCellValue>([
        ["B1", "=MIN(Sheet1:Sheet3!A1)"], // Min A1 across sheets 1-3
        ["B2", "=MIN(Sheet1:Sheet3!A1:A2)"], // Min A1:A2 across sheets 1-3
      ])
    );

    const cell = (sheetName: string, ref: string, debug?: boolean) =>
      engine.getCellValue(
        {
          sheetName,
          ...parseCellReference(ref),
        },
        debug
      );

    // ENGINE ISSUE: 3D references like Sheet1:Sheet3!A1 not supported
    expect(cell(sheet1Name, "B1", true)).toBe(10); // Min of 10, 10, 10 is 10
    expect(cell(sheet1Name, "B2", true)).toBe(10); // Min of all values is 10
  });

  test("with dynamic arrays (SEQUENCE)", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(3, 2, 10, 5)"], // Creates 2x3 array starting at 10, step 5
        ["D1", "=MIN(A1:B3)"], // Min the entire spilled array
        ["D2", "=MIN(A1:A3)"], // Min first column of spilled array
      ])
    );

    // SEQUENCE(3, 2, 10, 5) produces:
    // A1: 10, B1: 15
    // A2: 20, B2: 25
    // A3: 30, B3: 35
    // Min: 10
    expect(cell("D1")).toBe(10);
    expect(cell("D2")).toBe(10); // Min of 10, 20, 30 is 10
  });

  test("MIN used in dynamic array context", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        // Create multiple ranges to find min
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", 5],
        ["B2", 15],
        ["B3", 25],
        ["C1", 2],
        ["C2", 4],
        ["C3", 6],

        // Use MIN with SEQUENCE to create spilled results
        ["E1", "=MIN(A1:A3, SEQUENCE(3, 1, 100, 10))"], // MIN with dynamic array argument
        ["F1", "=SEQUENCE(3) + MIN(B1:B3)"], // MIN result used in array operation
      ])
    );

    // E1: MIN(A1:A3, SEQUENCE(3, 1, 100, 10))
    // A1:A3 = {10, 20, 30}
    // SEQUENCE(3, 1, 100, 10) = {100, 110, 120}
    // Min = 10
    expect(cell("E1")).toBe(10);

    // F1: SEQUENCE(3) + MIN(B1:B3) should spill
    // MIN(B1:B3) = min(5, 15, 25) = 5
    // SEQUENCE(3) = {1; 2; 3}
    // Result = {6; 7; 8}
    expect(cell("F1")).toBe(6); // 1 + 5
    expect(cell("F2")).toBe(7); // 2 + 5
    expect(cell("F3")).toBe(8); // 3 + 5
  });

  test("handling infinity", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=1/0"], // Positive infinity
        ["A2", "=-1/0"], // Negative infinity
        ["A3", 10],
        ["B1", "=MIN(A1, A3)"], // Min with positive infinity
        ["B2", "=MIN(A2, A3)"], // Min with negative infinity
        ["B3", "=MIN(A1, A2)"], // Min of both infinities
      ])
    );

    // ENGINE ISSUE: Division by zero (1/0) might not produce Infinity
    expect(cell("A1")).toBe("INFINITY");
    expect(cell("A2")).toBe("-INFINITY");

    // MIN with infinity
    expect(cell("B1")).toBe(10); // Min of Inf and 10 is 10
    expect(cell("B2")).toBe("-INFINITY"); // Min of -Inf and 10 is -Inf
    expect(cell("B3")).toBe("-INFINITY"); // Min of Inf and -Inf is -Inf
  });

  test("error handling", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "text"],
        ["A2", 10],
        ["A3", true],
        ["B1", "=MIN(A1, A2)"], // Mix of text and number
        ["B2", "=MIN(A2, A3)"], // Mix of number and boolean
        ["B3", "=MIN()"], // No arguments
      ])
    );

    // These should return errors due to non-numeric values
    expect(cell("B1")).toBe("#VALUE!");
    expect(cell("B2")).toBe("#VALUE!");

    // ENGINE ISSUE: MIN() with no arguments causes parse error instead of being handled by function
    expect(cell("B3")).toBe("#VALUE!"); // MIN with no arguments should return error
  });

  test("mixed argument types", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["B1", "=SEQUENCE(2, 1, 100, 50)"], // {100; 150}
        ["C1", "=MIN(5, A1:A2, B1:B2, 25)"], // Mix scalars, ranges, and dynamic arrays
      ])
    );

    // Min of 5, 10, 20, 100, 150, 25 is 5
    expect(cell("C1")).toBe(5);
  });

  test("MIN() with zero arguments", () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=MIN()"], // Should be allowed by parser
      ])
    );

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    // MIN with no arguments should return error
    expect(cell("A1")).toBe(FormulaError.VALUE);
  });

  test("single argument", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 42],
        ["B1", "=MIN(A1)"], // Single cell reference
        ["B2", "=MIN(42)"], // Single scalar
      ])
    );

    expect(cell("B1")).toBe(42); // Min of single value is the value itself
    expect(cell("B2")).toBe(42); // Min of single value is the value itself
  });

  test("negative numbers", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=MIN(-10, -5, -20, -1)"], // All negative
        ["A2", "=MIN(-10, 5, -20, 1)"], // Mix of positive and negative
        ["A3", "=MIN(0, -1, 1)"], // Including zero
      ])
    );

    expect(cell("A1")).toBe(-20); // Min of negatives
    expect(cell("A2")).toBe(-20); // Min including negatives
    expect(cell("A3")).toBe(-1); // Min with zero
  });

  test("decimal numbers", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=MIN(1.5, 2.7, 0.3, 1.9)"], // Decimal numbers
        ["A2", "=MIN(0.001, 0.002, 0.0005)"], // Small decimals
      ])
    );

    expect(cell("A1")).toBe(0.3);
    expect(cell("A2")).toBe(0.0005);
  });
});
