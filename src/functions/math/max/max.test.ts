import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("MAX function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  test("basic scalar arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "=MAX(6, 2, 8, 1, 9)"]])
    );

    expect(cell("A1")).toBe(9); // Maximum of 6, 2, 8, 1, 9 is 9
  });

  test("with cell references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 30],
        ["A2", 10],
        ["A3", 20],
        ["B1", "=MAX(A1:A3)"],
      ])
    );

    expect(cell("B1")).toBe(30); // Maximum of 30, 10, 20 is 30
  });

  test("with structured references", () => {
    // Create a table with data
    engine.setSheetContent(
      sheetAddress,
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
        ["D1", "=MAX(DataTable[Value])"], // Max the Value column
        ["D2", "=MAX(DataTable[Count])"], // Max the Count column
      ])
    );

    // Define the table
    engine.addTable({
      tableName: "DataTable",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "A1",
      numRows: { type: "number", value: 3 }, // 3 data rows
      numCols: 3, // 3 columns: Name, Value, Count
    });

    // ENGINE ISSUE: Structured references like DataTable[Value] not supported
    expect(cell("D1", true)).toBe(200); // Max of 100, 200, 150 is 200
    expect(cell("D2", true)).toBe(7); // Max of 5, 3, 7 is 7
  });

  test("with named expressions", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 30],
        ["A2", 10],
        ["A3", 20],
        ["B1", 15],
        ["B2", 5],
        ["C1", "=MAX(VALUES_A)"], // Max named range
        ["C2", "=MAX(VALUES_A, VALUES_B)"], // Max multiple named ranges
        ["C3", "=MAX(SINGLE_VALUE, 10)"], // Max named value with scalar
      ])
    );

    // Define named expressions
    engine.addNamedExpression({
      expression: "A1:A3",
      expressionName: "VALUES_A",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
    });
    engine.addNamedExpression({
      expression: "B1:B2",
      expressionName: "VALUES_B",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
    });
    engine.addNamedExpression({
      expression: "25",
      expressionName: "SINGLE_VALUE",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
    });

    // ENGINE ISSUE: Named expressions that evaluate to ranges not supported in function calls
    expect(cell("C1")).toBe(30); // Max of 30, 10, 20 is 30
    expect(cell("C2")).toBe(30); // Max of 30, 10, 20, 15, 5 is 30
    expect(cell("C3")).toBe(25); // Max of 25, 10 is 25
  });

  test("with cross-sheet references", () => {
    const sheet1Name = engine.addSheet({ workbookName, sheetName: "Sheet1" }).name;
    const sheet2Name = engine.addSheet({ workbookName, sheetName: "Sheet2" }).name;

    // Set up data on Sheet1
    engine.setSheetContent(
      { workbookName, sheetName: sheet1Name },
      new Map<string, SerializedCellValue>([
        ["A1", 100],
        ["A2", 200],
        ["A3", 300],
        ["B1", "=MAX(Sheet2!B1:B2)"], // Max range from Sheet2
        ["B2", "=MAX(A1:A3, Sheet2!B1:B2)"], // Max local and cross-sheet ranges
        ["B3", "=MAX(Sheet1!A1, Sheet2!B1, 25)"], // Mix of cross-sheet cells and scalar
      ])
    );

    // Set up data on Sheet2
    engine.setSheetContent(
      { workbookName, sheetName: sheet2Name },
      new Map<string, SerializedCellValue>([
        ["B1", 50],
        ["B2", 75],
      ])
    );

    const cell = (sheetName: string, ref: string) =>
      engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) });

    // ENGINE ISSUE: Cross-sheet references like Sheet2!B1:B2 not supported
    expect(cell(sheet1Name, "B1")).toBe(75); // Max of 50, 75 is 75
    expect(cell(sheet1Name, "B2")).toBe(300); // Max of 100, 200, 300, 50, 75 is 300
    expect(cell(sheet1Name, "B3")).toBe(100); // Max of 100, 50, 25 is 100
  });

  test.skip("with 3D sheet references", () => {
    const sheet1Name = engine.addSheet({ workbookName, sheetName: "Sheet1" }).name;
    const sheet2Name = engine.addSheet({ workbookName, sheetName: "Sheet2" }).name;
    const sheet3Name = engine.addSheet({ workbookName, sheetName: "Sheet3" }).name;

    // Set up same data on all sheets
    [sheet1Name, sheet2Name, sheet3Name].forEach((sheetName) => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
        ])
      );
    });

    // Create 3D reference formulas
    engine.setSheetContent(
      { workbookName, sheetName: sheet1Name },
      new Map<string, SerializedCellValue>([
        ["B1", "=MAX(Sheet1:Sheet3!A1)"], // Max A1 across sheets 1-3
        ["B2", "=MAX(Sheet1:Sheet3!A1:A2)"], // Max A1:A2 across sheets 1-3
      ])
    );

    const cell = (sheetName: string, ref: string, debug?: boolean) =>
      engine.getCellValue(
        {
          sheetName,
          workbookName,
          ...parseCellReference(ref),
        },
        debug
      );

    // ENGINE ISSUE: 3D references like Sheet1:Sheet3!A1 not supported
    expect(cell(sheet1Name, "B1", true)).toBe(10); // Max of 10, 10, 10 is 10
    expect(cell(sheet1Name, "B2", true)).toBe(20); // Max of all values is 20
  });

  test("with dynamic arrays (SEQUENCE)", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(3, 2, 10, 5)"], // Creates 2x3 array starting at 10, step 5
        ["D1", "=MAX(A1:B3)"], // Max the entire spilled array
        ["D2", "=MAX(A1:A3)"], // Max first column of spilled array
      ])
    );

    // SEQUENCE(3, 2, 10, 5) produces:
    // A1: 10, B1: 15
    // A2: 20, B2: 25
    // A3: 30, B3: 35
    // Max: 35
    expect(cell("D1")).toBe(35);
    expect(cell("D2")).toBe(30); // Max of 10, 20, 30 is 30
  });

  test("MAX used in dynamic array context", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Create multiple ranges to find max
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", 5],
        ["B2", 15],
        ["B3", 25],
        ["C1", 2],
        ["C2", 4],
        ["C3", 6],

        // Use MAX with SEQUENCE to create spilled results
        ["E1", "=MAX(A1:A3, SEQUENCE(3, 1, 100, 10))"], // MAX with dynamic array argument
        ["F1", "=SEQUENCE(3) + MAX(B1:B3)"], // MAX result used in array operation
      ])
    );

    // E1: MAX(A1:A3, SEQUENCE(3, 1, 100, 10))
    // A1:A3 = {10, 20, 30}
    // SEQUENCE(3, 1, 100, 10) = {100, 110, 120}
    // Max = 120
    expect(cell("E1")).toBe(120);

    // F1: SEQUENCE(3) + MAX(B1:B3) should spill
    // MAX(B1:B3) = max(5, 15, 25) = 25
    // SEQUENCE(3) = {1; 2; 3}
    // Result = {26; 27; 28}
    expect(cell("F1")).toBe(26); // 1 + 25
    expect(cell("F2")).toBe(27); // 2 + 25
    expect(cell("F3")).toBe(28); // 3 + 25
  });

  test("handling infinity", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=1/0"], // Positive infinity
        ["A2", "=-1/0"], // Negative infinity
        ["A3", 10],
        ["B1", "=MAX(A1, A3)"], // Max with positive infinity
        ["B2", "=MAX(A2, A3)"], // Max with negative infinity
        ["B3", "=MAX(A1, A2)"], // Max of both infinities
      ])
    );

    // ENGINE ISSUE: Division by zero (1/0) might not produce Infinity
    expect(cell("A1")).toBe("INFINITY");
    expect(cell("A2")).toBe("-INFINITY");

    // MAX with infinity
    expect(cell("B1")).toBe("INFINITY"); // Max of Inf and 10 is Inf
    expect(cell("B2")).toBe(10); // Max of -Inf and 10 is 10
    expect(cell("B3")).toBe("INFINITY"); // Max of Inf and -Inf is Inf
  });

  test("error handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "text"],
        ["A2", 10],
        ["A3", true],
        ["B1", "=MAX(A1, A2)"], // Mix of text and number
        ["B2", "=MAX(A2, A3)"], // Mix of number and boolean
        ["B3", "=MAX()"], // No arguments
      ])
    );

    // These should return errors due to non-numeric values
    expect(cell("B1")).toBe(FormulaError.VALUE);
    expect(cell("B2")).toBe(FormulaError.VALUE);

    // ENGINE ISSUE: MAX() with no arguments causes parse error instead of being handled by function
    expect(cell("B3")).toBe(FormulaError.VALUE); // MAX with no arguments should return error
  });

  test("mixed argument types", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["B1", "=SEQUENCE(2, 1, 100, 50)"], // {100; 150}
        ["C1", "=MAX(5, A1:A2, B1:B2, 25)"], // Mix scalars, ranges, and dynamic arrays
      ])
    );

    // Max of 5, 10, 20, 100, 150, 25 is 150
    expect(cell("C1")).toBe(150);
  });

  test("MAX() with zero arguments", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    const sheetName = engine.addSheet({ workbookName, sheetName: "Sheet1" }).name;

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "=MAX()"], // Should be allowed by parser
      ])
    );

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) });

    // MAX with no arguments should return error
    expect(cell("A1")).toBe(FormulaError.VALUE);
  });

  test("single argument", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 42],
        ["B1", "=MAX(A1)"], // Single cell reference
        ["B2", "=MAX(42)"], // Single scalar
      ])
    );

    expect(cell("B1")).toBe(42); // Max of single value is the value itself
    expect(cell("B2")).toBe(42); // Max of single value is the value itself
  });

  test("negative numbers", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=MAX(-10, -5, -20, -1)"], // All negative
        ["A2", "=MAX(-10, 5, -20, 1)"], // Mix of positive and negative
        ["A3", "=MAX(0, -1, 1)"], // Including zero
      ])
    );

    expect(cell("A1")).toBe(-1); // Max of negatives
    expect(cell("A2")).toBe(5); // Max including positives
    expect(cell("A3")).toBe(1); // Max with zero
  });

  test("decimal numbers", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=MAX(1.5, 2.7, 0.3, 1.9)"], // Decimal numbers
        ["A2", "=MAX(0.001, 0.002, 0.0005)"], // Small decimals
      ])
    );

    expect(cell("A1")).toBe(2.7);
    expect(cell("A2")).toBe(0.002);
  });
});
