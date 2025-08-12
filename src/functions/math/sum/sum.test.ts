import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("SUM function", () => {
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
      new Map<string, SerializedCellValue>([["A1", "=SUM(1, 2, 3)"]])
    );

    expect(cell("A1")).toBe(6);
  });

  test("with cell references", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", "=SUM(A1:A3)"],
      ])
    );

    expect(cell("B1")).toBe(60);
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
        ["D1", "=SUM(DataTable[Value])"], // Sum the Value column
        ["D2", "=SUM(DataTable[Count])"], // Sum the Count column
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
    expect(cell("D1", true)).toBe(450); // 100 + 200 + 150
    expect(cell("D2", true)).toBe(15); // 5 + 3 + 7
  });

  test("with named expressions", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", 5],
        ["B2", 15],
        ["C1", "=SUM(VALUES_A)"], // Sum named range
        ["C2", "=SUM(VALUES_A, VALUES_B)"], // Sum multiple named ranges
        ["C3", "=SUM(SINGLE_VALUE, 10)"], // Sum named value with scalar
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
    expect(cell("C1")).toBe(60); // 10 + 20 + 30
    expect(cell("C2")).toBe(80); // 10 + 20 + 30 + 5 + 15
    expect(cell("C3")).toBe(35); // 25 + 10
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
        ["B1", "=SUM(Sheet2!B1:B2)"], // Sum range from Sheet2
        ["B2", "=SUM(A1:A3, Sheet2!B1:B2)"], // Sum local and cross-sheet ranges
        ["B3", "=SUM(Sheet1!A1, Sheet2!B1, 25)"], // Mix of cross-sheet cells and scalar
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
    expect(cell(sheet1Name, "B1")).toBe(125); // 50 + 75
    expect(cell(sheet1Name, "B2")).toBe(725); // 100 + 200 + 300 + 50 + 75
    expect(cell(sheet1Name, "B3")).toBe(175); // 100 + 50 + 25
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
        ["B1", "=SUM(Sheet1:Sheet3!A1)"], // Sum A1 across sheets 1-3
        ["B2", "=SUM(Sheet1:Sheet3!A1:A2)"], // Sum A1:A2 across sheets 1-3
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
    expect(cell(sheet1Name, "B1", true)).toBe(30); // 10 + 10 + 10
    expect(cell(sheet1Name, "B2", true)).toBe(90); // (10+20) + (10+20) + (10+20)
  });

  test.skip("with cross-sheet ranges (engine feature needed)", () => {
    // Cross-sheet ranges like Sheet2!B1:B2 need engine implementation
    // SUM function supports this but engine parser may not handle the syntax
  });

  test.skip("with 3D sheet references (engine feature needed)", () => {
    // 3D references like SUM(Sheet1:Sheet3!A1) would need engine implementation
    // This allows summing the same cell/range across multiple sheets
    // Example: SUM(Sheet1:Sheet3!A1) sums A1 from Sheet1, Sheet2, and Sheet3
  });

  test("with dynamic arrays (SEQUENCE)", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(3, 2, 10, 5)"], // Creates 2x3 array starting at 10, step 5
        ["D1", "=SUM(A1:B3)"], // Sum the entire spilled array
        ["D2", "=SUM(A1:A3)"], // Sum first column of spilled array
      ])
    );

    // SEQUENCE(3, 2, 10, 5) produces:
    // A1: 10, B1: 15
    // A2: 20, B2: 25
    // A3: 30, B3: 35
    // Total: 10 + 15 + 20 + 25 + 30 + 35 = 135
    expect(cell("D1")).toBe(135);
    expect(cell("D2")).toBe(60); // 10 + 20 + 30
  });

  test("SUM used in dynamic array context", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        // Create multiple ranges to sum
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", 5],
        ["B2", 15],
        ["B3", 25],
        ["C1", 2],
        ["C2", 4],
        ["C3", 6],

        // Use SUM with SEQUENCE to create spilled results
        ["E1", "=SUM(A1:A3, SEQUENCE(3, 1, 100, 10))"], // SUM with dynamic array argument
        ["F1", "=SEQUENCE(3) + SUM(B1:B3)"], // SUM result used in array operation
      ])
    );

    // E1: SUM(A1:A3, SEQUENCE(3, 1, 100, 10))
    // A1:A3 = 10+20+30 = 60
    // SEQUENCE(3, 1, 100, 10) = {100; 110; 120} = 330
    // Total = 60 + 330 = 390
    expect(cell("E1")).toBe(390);

    // F1: SEQUENCE(3) + SUM(B1:B3) should spill
    // SUM(B1:B3) = 5+15+25 = 45
    // SEQUENCE(3) = {1; 2; 3}
    // Result = {46; 47; 48}
    expect(cell("F1")).toBe(46); // 1 + 45
    expect(cell("F2")).toBe(47); // 2 + 45
    expect(cell("F3")).toBe(48); // 3 + 45
  });

  test.skip("handling infinity", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=1/0"], // Positive infinity
        ["A2", "=-1/0"], // Negative infinity
        ["A3", 10],
        ["B1", "=SUM(A1, A3)"], // Sum with positive infinity
        ["B2", "=SUM(A2, A3)"], // Sum with negative infinity
        ["B3", "=SUM(A1, A2)"], // Sum of both infinities (should be NaN/error)
      ])
    );

    // ENGINE ISSUE: Division by zero (1/0) might not produce Infinity
    expect(cell("A1")).toBe(Infinity);
    expect(cell("A2")).toBe(-Infinity);

    // SUM with infinity should return infinity
    expect(cell("B1")).toBe(Infinity); // Inf + 10 = Inf
    expect(cell("B2")).toBe(-Infinity); // -Inf + 10 = -Inf

    // Sum of positive and negative infinity should be error (NaN)
    const result = cell("B3");
    expect(typeof result === "string" && result.includes("#")).toBe(true); // Expect some error
  });

  test.skip("error handling", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "text"],
        ["A2", 10],
        ["A3", true],
        ["B1", "=SUM(A1, A2)"], // Mix of text and number
        ["B2", "=SUM(A2, A3)"], // Mix of number and boolean
        ["B3", "=SUM()"], // No arguments
      ])
    );

    // These should return errors due to non-numeric values
    expect(cell("B1")).toBe(FormulaError.VALUE);
    expect(cell("B2")).toBe(FormulaError.VALUE);

    // ENGINE ISSUE: SUM() with no arguments causes parse error instead of being handled by function
    expect(cell("B3")).toBe(0); // SUM with no arguments typically returns 0
  });

  test("mixed argument types", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["B1", "=SEQUENCE(2, 1, 100, 50)"], // {100; 150}
        ["C1", "=SUM(5, A1:A2, B1:B2, 25)"], // Mix scalars, ranges, and dynamic arrays
      ])
    );

    // 5 + (10+20) + (100+150) + 25 = 5 + 30 + 250 + 25 = 310
    expect(cell("C1")).toBe(310);
  });

  test("SUM() with zero arguments", () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=SUM()"], // Should be allowed by parser
      ])
    );

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    expect(cell("A1")).toBe(0);
  });
});

/* ========================================================================
 * COMMENTED OUT TESTS - ENGINE FEATURES NEEDED
 *
 * These tests identify specific engine limitations that need to be addressed.
 * Each test is grouped by the type of engine feature required.
 * ======================================================================== */

/* 
// ==================== PARSER ISSUES ====================

describe("SUM function - Parser Issues", () => {



});

// ==================== CROSS-SHEET REFERENCE ISSUES ====================

describe("SUM function - Cross-Sheet References", () => {
  test("Cross-sheet ranges", () => {
    // ISSUE: Sheet2!B1:B2 syntax not supported
    // LOCATION: Parser - cross-sheet range parsing
    // EXPECTED: Should parse and evaluate cross-sheet ranges
    
    const engine = FormulaEngine.buildEmpty();
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;

    engine.setSheetContent(sheet2Name, new Map([
      ["B1", 50], ["B2", 75]
    ]));

    engine.setSheetContent(sheet1Name, new Map([
      ["A1", "=SUM(Sheet2!B1:B2)"],  // Should sum range from Sheet2
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName: sheet1Name, ...parseCellReference(ref) });

    expect(cell("A1")).toBe(125); // 50 + 75
  });

  test("Cross-sheet individual cell references", () => {
    // ISSUE: Sheet2!B1 individual cell references might not work
    // LOCATION: Parser - cross-sheet cell parsing
    // EXPECTED: Should parse and evaluate individual cross-sheet cells
    
    const engine = FormulaEngine.buildEmpty();
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;

    engine.setSheetContent(sheet1Name, new Map([["A1", 100]]));
    engine.setSheetContent(sheet2Name, new Map([["B1", 50]]));

    engine.setSheetContent(sheet1Name, new Map([
      ["C1", "=SUM(Sheet1!A1, Sheet2!B1, 25)"],  // Mix local, cross-sheet, and scalar
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName: sheet1Name, ...parseCellReference(ref) });

    expect(cell("C1")).toBe(175); // 100 + 50 + 25
  });
});

// ==================== 3D REFERENCE ISSUES ====================

describe("SUM function - 3D References", () => {
  test("3D cell references", () => {
    // ISSUE: Parser recognizes 3d-range nodes but evaluator doesn't handle them
    // LOCATION: FormulaEvaluator.evaluateNode() - missing case for "3d-range"
    // EXPECTED: Should sum same cell across multiple sheets
    
    const engine = FormulaEngine.buildEmpty();
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;
    const sheet3Name = engine.addSheet("Sheet3").name;

    // Set up same data on all sheets
    [sheet1Name, sheet2Name, sheet3Name].forEach(sheetName => {
      engine.setSheetContent(sheetName, new Map([["A1", 10]]));
    });

    engine.setSheetContent(sheet1Name, new Map([
      ["B1", "=SUM(Sheet1:Sheet3!A1)"],  // Sum A1 across sheets 1-3
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName: sheet1Name, ...parseCellReference(ref) });

    expect(cell("B1")).toBe(30); // 10 + 10 + 10
  });

  test("3D range references", () => {
    // ISSUE: Parser recognizes 3d-range with range nodes but evaluator doesn't handle them
    // LOCATION: FormulaEvaluator.evaluateNode() - missing case for "3d-range" with ranges
    // EXPECTED: Should sum same range across multiple sheets
    
    const engine = FormulaEngine.buildEmpty();
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;
    const sheet3Name = engine.addSheet("Sheet3").name;

    // Set up same data on all sheets
    [sheet1Name, sheet2Name, sheet3Name].forEach(sheetName => {
      engine.setSheetContent(sheetName, new Map([
        ["A1", 10], ["A2", 20]
      ]));
    });

    engine.setSheetContent(sheet1Name, new Map([
      ["B1", "=SUM(Sheet1:Sheet3!A1:A2)"],  // Sum A1:A2 across sheets 1-3
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName: sheet1Name, ...parseCellReference(ref) });

    expect(cell("B1")).toBe(90); // (10+20) + (10+20) + (10+20)
  });
});

// ==================== STRUCTURED REFERENCE ISSUES ====================

describe("SUM function - Structured References", () => {
  test("Table column references", () => {
    // ISSUE: DataTable[ColumnName] syntax not supported
    // LOCATION: Parser - structured reference parsing
    // EXPECTED: Should parse and resolve table column references
    
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(sheetName, new Map([
      ["A1", "Name"], ["B1", "Value"], ["C1", "Count"],
      ["A2", "Item1"], ["B2", 100], ["C2", 5],
      ["A3", "Item2"], ["B3", 200], ["C3", 3],
      ["A4", "Item3"], ["B4", 150], ["C4", 7],
    ]));

    engine.addTable({
      tableName: "DataTable",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 4 },
      numCols: 3,
    });

    engine.setSheetContent(sheetName, new Map([
      ["D1", "=SUM(DataTable[Value])"],  // Should sum Value column
      ["D2", "=SUM(DataTable[Count])"],  // Should sum Count column
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    expect(cell("D1")).toBe(450); // 100 + 200 + 150
    expect(cell("D2")).toBe(15);  // 5 + 3 + 7
  });

  test("Table area references", () => {
    // ISSUE: DataTable[#All], DataTable[#Data], etc. not supported
    // LOCATION: Parser - structured reference area parsing
    // EXPECTED: Should parse and resolve table area references
    
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(sheetName, new Map([
      ["A1", "Value"],
      ["A2", 10], ["A3", 20], ["A4", 30],
    ]));

    engine.addTable({
      tableName: "MyTable",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 4 },
      numCols: 1,
    });

    engine.setSheetContent(sheetName, new Map([
      ["B1", "=SUM(MyTable[#Data])"],    // Should sum data rows only (exclude header)
      ["B2", "=SUM(MyTable[#All])"],     // Should sum all including header if numeric
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    expect(cell("B1")).toBe(60);  // 10 + 20 + 30 (data only)
    expect(cell("B2")).toBe(60);  // Same (header "Value" is text, ignored)
  });
});

// ==================== NAMED EXPRESSION ISSUES ====================

describe("SUM function - Named Expression Issues", () => {
  test("Named expressions that evaluate to ranges", () => {
    // ISSUE: Named expressions like VALUES_A that resolve to A1:A3 don't work in function calls
    // LOCATION: Named expression resolution in function context
    // EXPECTED: SUM(VALUES_A) should expand to SUM(A1:A3)
    
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(sheetName, new Map([
      ["A1", 10], ["A2", 20], ["A3", 30],
      ["B1", 5], ["B2", 15],
    ]));

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

    engine.setSheetContent(sheetName, new Map([
      ["C1", "=SUM(VALUES_A)"],            // Should expand to SUM(A1:A3)
      ["C2", "=SUM(VALUES_A, VALUES_B)"],  // Should expand to SUM(A1:A3, B1:B2)
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    expect(cell("C1")).toBe(60);  // 10 + 20 + 30
    expect(cell("C2")).toBe(80);  // 10 + 20 + 30 + 5 + 15
  });

  test("Cross-sheet named expressions", () => {
    // ISSUE: Named expressions referencing other sheets might not work
    // LOCATION: Named expression resolution with cross-sheet references
    // EXPECTED: Named expressions should work across sheets
    
    const engine = FormulaEngine.buildEmpty();
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;

    engine.setSheetContent(sheet2Name, new Map([
      ["A1", 100], ["A2", 200],
    ]));

    engine.addNamedExpression({
      expression: "Sheet2!A1:A2",
      expressionName: "REMOTE_VALUES",
      sheetName: sheet1Name,
    });

    engine.setSheetContent(sheet1Name, new Map([
      ["B1", "=SUM(REMOTE_VALUES)"],  // Should sum Sheet2!A1:A2
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName: sheet1Name, ...parseCellReference(ref) });

    expect(cell("B1")).toBe(300); // 100 + 200
  });
});

// ==================== ADDITIONAL EDGE CASES ====================

describe("SUM function - Advanced Features", () => {
  test("Circular reference detection", () => {
    // ISSUE: Might not properly detect and handle circular references
    // LOCATION: Dependency tracking
    // EXPECTED: Should detect circular references and return appropriate error
    
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(sheetName, new Map([
      ["A1", "=SUM(B1, 10)"],
      ["B1", "=SUM(A1, 20)"],  // Circular reference
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    expect(cell("A1")).toBe(FormulaError.CIRCULAR_REF);
    expect(cell("B1")).toBe(FormulaError.CIRCULAR_REF);
  });

  test("Volatile functions in SUM arguments", () => {
    // ISSUE: Might not properly handle volatile functions like NOW(), RAND()
    // LOCATION: Volatility tracking and recalculation
    // EXPECTED: SUM containing volatile functions should recalculate appropriately
    
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(sheetName, new Map([
      ["A1", "=RAND()"],           // Volatile function
      ["B1", "=SUM(A1, 100)"],     // Should be marked volatile
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    const firstValue = cell("B1");
    // Force recalculation
    engine.reevaluate();
    const secondValue = cell("B1");

    expect(typeof firstValue).toBe("number");
    expect(typeof secondValue).toBe("number");
    expect(firstValue).not.toBe(secondValue); // Should be different due to RAND()
  });

  test("Array formula behavior", () => {
    // ISSUE: Might not properly handle array formulas (if supported)
    // LOCATION: Array formula evaluation
    // EXPECTED: Array formulas should work with SUM
    
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("Sheet1").name;

    engine.setSheetContent(sheetName, new Map([
      ["A1", 1], ["A2", 2], ["A3", 3],
      ["B1", 10], ["B2", 20], ["B3", 30],
    ]));

    // Array formula: SUM(A1:A3 * B1:B3) = SUM({10, 40, 90}) = 140
    engine.setSheetContent(sheetName, new Map([
      ["C1", "{=SUM(A1:A3*B1:B3)}"],  // Array formula syntax
    ]));

    const cell = (ref: string) =>
      engine.getCellValue({ sheetName, ...parseCellReference(ref) });

    expect(cell("C1")).toBe(140); // (1*10) + (2*20) + (3*30)
  });
});

*/
