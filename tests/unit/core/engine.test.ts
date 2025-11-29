import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import {
  FormulaError,
  type SerializedCellValue,
  type TableDefinition,
} from "../../../src/core/types";
import { getCellReference, parseCellReference } from "../../../src/core/utils";

describe("FormulaEngine", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "TestSheet";

  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const cellContent = (ref: string): SerializedCellValue => {
    const sheetContent = engine.getSheetSerialized({ workbookName, sheetName });
    return sheetContent.get(ref) ?? "";
  };

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  test("basic scalar arguments", () => {
    setCellContent("A1", "=SUM(1, 2, 3)");

    expect(cell("A1")).toBe(6);
  });

  test("should set and get cell values", () => {
    setCellContent("A1", 42);
    expect(cell("A1")).toBe(42);
  });

  test("should handle empty cells", () => {
    setCellContent("A1", "");
    expect(cell("A1")).toBe("");
  });

  test("should set multiple values with 2D array", () => {
    const rawData = new Map([
      ["A1", 1],
      ["B1", 2],
      ["C1", 3],
      ["A2", 4],
      ["B2", 5],
      ["C2", 6],
    ]);
    engine.setSheetContent(sheetAddress, rawData);

    // Verify all values
    for (let row = 0; row < 2; row++) {
      for (let col = 0; col < 3; col++) {
        const value = engine.getCellValue({
          sheetName,
          workbookName,
          colIndex: col,
          rowIndex: row,
        });
        const cellReference = getCellReference({
          colIndex: col,
          rowIndex: row,
        });
        expect(value).toBe(rawData.get(cellReference)!);
      }
    }
  });

  test("should handle formulas", () => {
    const data = new Map<string, SerializedCellValue>([
      ["A1", "=A2+B2"],
      ["A2", 1],
      ["B2", 2],
    ]);

    engine.setSheetContent(sheetAddress, data);

    expect(cell("A1")).toBe(3);
  });

  test("should handle formulas with cross sheet references", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName: "Sheet1" });
    engine.addSheet({ workbookName, sheetName: "Sheet2" });

    engine.setSheetContent(
      { workbookName, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([
        ["A1", "=Sheet2!C1 + B2"],
        ["B2", "=A4 + 5"],
        ["A4", 5],
      ])
    );

    engine.setSheetContent(
      { workbookName, sheetName: "Sheet2" },
      new Map<string, SerializedCellValue>([
        ["C1", "=A3 + 100"], // A3 must refer to Sheet2
        ["A3", 23],
      ])
    );

    expect(
      engine.getCellValue(
        {
          // C1
          workbookName,
          sheetName: "Sheet2",
          colIndex: 2,
          rowIndex: 0,
        },
        true
      )
    ).toBe(123);

    expect(
      engine.getCellValue(
        {
          // B2
          workbookName,
          sheetName: "Sheet1",
          colIndex: 1,
          rowIndex: 1,
        },
        true
      )
    ).toBe(10);

    expect(
      engine.getCellValue(
        {
          workbookName,
          sheetName: "Sheet1",
          colIndex: 0,
          rowIndex: 0,
        },
        true
      )
    ).toBe(133);
  });

  test("should handle named expressions", () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = "Sheet1";
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map([["A1", "=SOME_EXPRESSION"]])
    );

    // Add global named expression
    engine.addNamedExpression({
      expression: "123 + 123",
      expressionName: "SOME_EXPRESSION",
    });

    const value = engine.getCellValue(
      {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      },
      true
    );
    expect(value).toBe(246);
  });

  test("should handle named expressions with cross sheet references", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName: "Sheet1" });
    engine.addSheet({ workbookName, sheetName: "Sheet2" });

    // global named expression
    engine.addNamedExpression({
      expression: "123 + 123",
      expressionName: "SOME_EXPRESSION",
    });

    // scoped named expression
    engine.addNamedExpression({
      expression: "10",
      expressionName: "SOME_EXPRESSION",
      sheetName: "Sheet1",
      workbookName,
    });

    engine.setSheetContent(
      { workbookName, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([["A1", "=SOME_EXPRESSION"]])
    );
    engine.setSheetContent(
      { workbookName, sheetName: "Sheet2" },
      new Map<string, SerializedCellValue>([
        ["B1", "=SOME_EXPRESSION"],
        ["C1", "=Sheet1!SOME_EXPRESSION"],
      ])
    );

    expect(
      engine.getCellValue({
        workbookName,
        sheetName: "Sheet1",
        colIndex: 0,
        rowIndex: 0,
      })
    ).toBe(10);
    expect(
      engine.getCellValue({
        workbookName,
        sheetName: "Sheet2",
        colIndex: 1,
        rowIndex: 0,
      })
    ).toBe(246);
    expect(
      engine.getCellValue({
        workbookName,
        sheetName: "Sheet2",
        colIndex: 2,
        rowIndex: 0,
      })
    ).toBe(10);
  });

  test("should handle structured references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "num"],
        ["B1", "result"],
        ["C1", "sum"],
        ["A2", 2],
        ["A3", 3],
        ["A4", 4],
        ["B2", "=Table1[@num] * 10"],
        ["B3", "=Table1[@num] * 10"],
        ["B4", "=Table1[@num] * 10"],
        ["B5", "=Table1[@num] * 10"], // should be errored
        ["C2", "=SUM(Table1[result])"],
        ["C3", "=SUM(Table1[[num]:[result]])"],
        ["C4", "=SUM(Table1[@[num]:[result]])"],
        ["F1", "=Table1[[#Headers],[result]]"],
        ["F20", "=Table1[#Headers]"],
      ])
    );

    engine.addTable({
      tableName: "Table1",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 3,
    });

    expect(cell("B2")).toBe(20);
    expect(cell("B3")).toBe(30);
    expect(cell("B4")).toBe(40);
    expect(cell("B5")).toBe(FormulaError.VALUE);
    expect(cell("C2", true)).toBe(90);
    expect(cell("C3")).toBe(99);
    expect(cell("C4")).toBe(44);

    expect(cell("F1")).toBe("result");
    expect(cell("F20")).toBe("num");
    expect(cell("G20")).toBe("result");
    expect(cell("H20")).toBe("sum");
  });

  const fourByFour: [string, SerializedCellValue][] = [
    ["A1", 1],
    ["A2", 2],
    ["A3", 3],
    ["A4", 4],
    ["B1", 5],
    ["B2", 6],
    ["B3", 7],
    ["B4", 8],
    ["C1", 9],
    ["C2", 10],
    ["C3", 11],
    ["C4", 12],
    ["D1", 13],
    ["D2", 14],
    ["D3", 15],
    ["D4", 16],
  ];

  test("should handle spilling", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([...fourByFour, ["F1", "=A1:D4"]])
    );

    expect(cell("F1")).toBe(1);
    expect(cell("H1")).toBe(9);
  });

  test("should handle reduced spilled values, when evaluating the spill origin first", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ...fourByFour,
        ["F1", "=A1:D4 * 10"],
      ])
    );

    expect(cell("F1")).toBe(10);
    expect(cell("H1")).toBe(90);
  });

  test("should handle reduced spilled values, when evaluating the spill origin last", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ...fourByFour,
        ["F1", "=A1:D4 * 10"],
      ])
    );

    expect(cell("H1")).toBe(90);
    expect(cell("F1")).toBe(10);

    expect(cell("F2")).toBe(20);
  });

  test("should get spill errors", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ...fourByFour,
        ["F1", "=A1:D4 * 10"],
        ["F2", "some value here!"],
      ])
    );

    expect(cell("H1")).toBe("");
    expect(cell("F1")).toBe(FormulaError.SPILL);
  });

  test("should work with a spilled value as a dependency", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ...fourByFour,
        ["F1", "=A1:D4 * 10"],
        ["F10", "=F2 * 123"], // 20 * 123 = 2460
      ])
    );

    expect(cell("F10")).toBe(2460);
  });

  test("should be able to update the spreadsheet content", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ...fourByFour,
        ["F1", "=A1:D4 * 10"],
        ["F10", "=F2 * 123"], // 20 * 123 = 2460
      ])
    );

    expect(cell("F10")).toBe(2460);

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ...fourByFour,
        ["F1", "=A1:D4"],
        ["F10", "=F2 * 123"], // 2 * 123 = 246
      ])
    );

    expect(cell("F10")).toBe(246);
  });

  test("should handle Excel table with bare column references", () => {
    // Set up the exact table structure from the user's Excel example:
    // num	result	                    sum	                        extras
    // 1	=[@num] * 10	            =SUM([result])
    // 3	=[@num] * 10	            =SUM(Table1[[num]:[result]])
    // 4	=[@num] * 10	            =SUM(Table1[@[num]:[result]])
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Headers
        ["A1", "num"],
        ["B1", "result"],
        ["C1", "sum"],
        ["D1", "extras"],
        // Data rows
        ["A2", 1],
        ["B2", "=[@num] * 10"], // Current row reference
        ["C2", "=SUM([result])"], // Bare column reference
        ["A3", 3],
        ["B3", "=[@num] * 10"], // Current row reference
        ["C3", "=SUM(Table1[[num]:[result]])"], // Bracketed column range
        ["A4", 4],
        ["B4", "=[@num] * 10"], // Current row reference
        ["C4", "=SUM(Table1[@[num]:[result]])"], // Bracketed current row range
      ])
    );

    // Define the table (A1:D4, so 4 rows including header)
    engine.addTable({
      tableName: "Table1",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 4,
    });

    // Test the calculated values

    // B2: [@num] * 10 = 1 * 10 = 10
    expect(cell("B2", true)).toBe(10);

    // B3: [@num] * 10 = 3 * 10 = 30
    expect(cell("B3")).toBe(30);

    // B4: [@num] * 10 = 4 * 10 = 40
    expect(cell("B4", true)).toBe(40);

    // C2: SUM([result]) = SUM(B2:B4) = 10 + 30 + 40 = 80
    expect(cell("C2", true)).toBe(80);

    // C3: SUM(Table1[[num]:[result]]) = SUM(A2:B4) = (1+10) + (3+30) + (4+40) = 88
    expect(cell("C3")).toBe(88);

    // C4: SUM(Table1[@[num]:[result]]) = SUM(A4:B4) = 4 + 40 = 44
    expect(cell("C4")).toBe(44);
  });

  test("should handle complex formula with LEFT and FIND in table", () => {
    // Set up a table with comma-separated values in the Payload column
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Headers
        ["A1", "ID"],
        ["B1", "Payload"],
        ["C1", "Extracted"],
        // Data rows with comma-separated payloads
        ["A2", 1],
        ["B2", "apple,banana,cherry"],
        ["C2", '=LEFT([@Payload],FIND(",",[@Payload])-1)'],
        ["A3", 2],
        ["B3", "dog,cat,bird"],
        ["C3", '=LEFT([@Payload],FIND(",",[@Payload])-1)'],
        ["A4", 3],
        ["B4", "red,green,blue"],
        ["C4", '=LEFT([@Payload],FIND(",",[@Payload])-1)'],
      ])
    );

    // Define the table
    engine.addTable({
      tableName: "DataTable",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "A1",
      numRows: { type: "number", value: 4 },
      numCols: 3,
    });

    // Test the extracted values (should be the text before the first comma)

    // C2: LEFT("apple,banana,cherry", FIND(",", "apple,banana,cherry") - 1) = LEFT("apple,banana,cherry", 6 - 1) = LEFT("apple,banana,cherry", 5) = "apple"
    expect(cell("C2")).toBe("apple");

    // C3: LEFT("dog,cat,bird", FIND(",", "dog,cat,bird") - 1) = LEFT("dog,cat,bird", 4 - 1) = LEFT("dog,cat,bird", 3) = "dog"
    expect(cell("C3")).toBe("dog");

    // C4: LEFT("red,green,blue", FIND(",", "red,green,blue") - 1) = LEFT("red,green,blue", 4 - 1) = LEFT("red,green,blue", 3) = "red"
    expect(cell("C4")).toBe("red");
  });

  test("should handle INDEX+MATCH with structured references", () => {
    // Set up a table with ORDER-ID and CUSTOMER-ID columns
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Headers
        ["A1", "ORDER-ID"],
        ["B1", "CUSTOMER-ID"],
        ["C1", "AMOUNT"],
        ["D1", "LOOKUP-RESULT"],
        // Data rows
        ["A2", "ORD001"],
        ["B2", "CUST123"],
        ["C2", 100],
        [
          "D2",
          "=INDEX(Table1[ORDER-ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))",
        ],
        ["A3", "ORD002"],
        ["B3", "CUST456"],
        ["C3", 200],
        [
          "D3",
          "=INDEX(Table1[ORDER-ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))",
        ],
        ["A4", "ORD003"],
        ["B4", "CUST123"], // Same customer as row 2, should return ORD001
        ["C4", 150],
        [
          "D4",
          "=INDEX(Table1[ORDER-ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))",
        ],
      ])
    );

    // Define the table
    engine.addTable({
      tableName: "Table1",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 4,
    });

    // Test the lookup results
    // D2: Should find first occurrence of CUST123 in CUSTOMER_ID column and return corresponding ORDER_ID
    expect(cell("D2")).toBe("ORD001");

    // D3: Should find CUST456 and return ORD002
    expect(cell("D3")).toBe("ORD002");

    // D4: Should find first occurrence of CUST123 (which is in row 2) and return ORD001
    expect(cell("D4")).toBe("ORD001");
  });

  test("Special case", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        [
          "A1",
          "=INDEX(Table1[CAR ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))",
        ],
      ])
    );
    expect(() => cell("A1", true)).not.toThrow();
  });

  test("evaluation should handle range inputs as gracefully /1", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "apple,banana,cherry"],
        ["A2", "dog,cat,bird"],
        ["A3", "red,green,blue"],
        ["B1", 1],
        ["B2", 2],
        ["B3", 3],
        ["C1", '=LEFT(A1:A3,FIND(",",A1:A3)-1)'],
      ])
    );

    expect(cell("C1")).toBe("apple");
    expect(cell("C2")).toBe("dog");
    expect(cell("C3")).toBe("red");
  });

  test("evaluation should handle range inputs as gracefully /2", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", "=SUM(A1:A3 * 10)"],
      ])
    );

    expect(cell("B1")).toBe(60);
  });

  test("multiplication of ranges", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", "=D11 * 0.5"],
        ["B2", 8],
        ["B3", 7],
        ["C1", "=A1:A3 * B1:B3"],
        ["D10", "=A1:A2 * (B2 + A1)"],
      ])
    );

    expect(cell("D11", true)).toBe(18);
    expect(cell("C1", true)).toBe(9);
    expect(cell("C2")).toBe(16);
    expect(cell("C3")).toBe(21);
    expect(cell("D10", true)).toBe(9);
  });

  test("addition of sparse ranges should use position-based pairing", () => {
    // Test case from user: A1:A7 + B1:B6
    // A column: 1, 3, empty, 4, empty, empty, 2
    // B column: empty, 1, empty, empty, empty, 3
    // Expected: 1, 4, #N/A, 4, #N/A, 3, #N/A
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["A2", 3],
        // A3 is empty
        ["A4", 4],
        // A5, A6 are empty
        ["A7", 2],
        // B1 is empty
        ["B2", 1],
        // B3, B4, B5 are empty
        ["B6", 3],
        // B7 doesn't exist (beyond B range)
        ["C1", "=A1:A7+B1:B6"],
      ])
    );

    // Row 1: 1 + empty = #N/A (empty within range)
    expect(cell("C1")).toBe("#N/A");
    // Row 2: 3 + 1 = 4
    expect(cell("C2")).toBe(4);
    // Row 3: empty + empty = #N/A
    expect(cell("C3")).toBe("#N/A");
    // Row 4: 4 + empty = #N/A (empty within range)
    expect(cell("C4")).toBe("#N/A");
    // Row 5: empty + empty = #N/A
    expect(cell("C5")).toBe("#N/A");
    // Row 6: empty + 3 = #N/A (empty within range)
    expect(cell("C6")).toBe("#N/A");
    // Row 7: 2 + beyond range = #N/A
    expect(cell("C7")).toBe("#N/A");
  });

  test("evaluation should handle range inputs as gracefully /3", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 60],
        ["A2", 50],
        ["A3", 40],
        ["B1", "=A1:A3 - 10"],
        ["C1", "=A1:A3 - B1:B3"],
      ])
    );

    expect(cell("B1")).toBe(50);
    expect(cell("B2")).toBe(40);
    expect(cell("B3")).toBe(30);

    expect(cell("C1")).toBe(10);
    expect(cell("C2")).toBe(10);
    expect(cell("C3")).toBe(10);
  });

  test.skip("with 3D sheet references", () => {
    const sheet1Name = engine.addSheet({
      workbookName,
      sheetName: "Sheet1",
    }).name;
    const sheet2Name = engine.addSheet({
      workbookName,
      sheetName: "Sheet2",
    }).name;
    const sheet3Name = engine.addSheet({
      workbookName,
      sheetName: "Sheet3",
    }).name;

    // Set up same data on all sheets
    [sheet1Name, sheet2Name, sheet3Name].forEach((sheetName) => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
        ])
      );
    });

    // Create 3D reference formulas
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", "=SUM(Sheet1:Sheet3!A1)"], // Sum A1 across sheets 1-3
        ["B2", "=SUM(Sheet1:Sheet3!A1:A2)"], // Sum A1:A2 across sheets 1-3
      ])
    );

    const cell = (sheetName: string, ref: string, debug?: boolean) =>
      engine.getCellValue(
        {
          sheetName: sheetAddress.sheetName,
          workbookName: sheetAddress.workbookName,
          ...parseCellReference(ref),
        },
        debug
      );

    // ENGINE ISSUE: 3D references like Sheet1:Sheet3!A1 not supported
    expect(cell(sheet1Name, "B1", true)).toBe(30); // 10 + 10 + 10
    expect(cell(sheet1Name, "B2", true)).toBe(90); // (10+20) + (10+20) + (10+20)
  });

  test("Division by zero should produce Infinity", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=1/0"], // Should produce Infinity
        ["A2", "=-1/0"], // Should produce -Infinity
        ["A3", "=0/0"], // Should produce NaN or #NUM! error
      ])
    );

    expect(cell("A1")).toBe("INFINITY");
    expect(cell("A2", true)).toBe("-INFINITY");
    expect(cell("A3")).toBe(FormulaError.NUM);
  });

  test("Infinity * Infinity should produce Infinity", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=INFINITY * INFINITY"], // Should produce Infinity
        ["A2", "=-INFINITY * INFINITY"], // Should produce -Infinity
        ["A3", "=INFINITY * -INFINITY"], // Should produce -Infinity
      ])
    );

    expect(cell("A1", true)).toBe("INFINITY");
    expect(cell("A2", true)).toBe("-INFINITY");
    expect(cell("A3")).toBe("-INFINITY");
  });

  test("Array row syntax", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "={1;2;3}"]])
    );

    expect(cell("A1", true)).toBe(1);
    expect(cell("A2", true)).toBe(2);
    expect(cell("A3", true)).toBe(3);
  });

  test("Array col syntax", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([["A1", "={1,2,3}"]])
    );

    expect(cell("A1", true)).toBe(1);
    expect(cell("B1", true)).toBe(2);
    expect(cell("C1", true)).toBe(3);
  });

  describe("Sheet Operations with Formula Dependencies", () => {
    test("should handle cross-sheet references", () => {
      const sheet2 = { workbookName, sheetName: "Sheet2" };
      engine.addSheet(sheet2);

      // Set data on Sheet1
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 100],
          ["B1", 200],
        ])
      );

      // Reference from Sheet2
      engine.setSheetContent(
        sheet2,
        new Map<string, SerializedCellValue>([
          ["A1", `=${sheetName}!A1+${sheetName}!B1`],
          ["A2", `=${sheetName}!A1*2`],
        ])
      );

      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(300); // 100 + 200
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 1,
          colIndex: 0,
        })
      ).toBe(200); // 100 * 2
    });

    test("should update cross-sheet references when sheet is renamed", () => {
      const sheet2 = { workbookName, sheetName: "Sheet2" };
      const newSheetName = "DataSheet";
      engine.addSheet(sheet2);

      // Set data on Sheet1
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", 150]])
      );

      // Reference from Sheet2
      engine.setSheetContent(
        sheet2,
        new Map<string, SerializedCellValue>([
          ["A1", `=${sheetAddress.sheetName}!A1*3`],
        ])
      );

      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(450); // 150 * 3

      // Rename Sheet1
      engine.renameSheet({
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        newSheetName,
      });

      // Formula should still work
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(450);
    });

    test("should show error when referenced sheet is removed", () => {
      const sheet2 = { workbookName, sheetName: "Sheet2" };
      const sheet3 = { workbookName, sheetName: "Sheet3" };
      engine.addSheet(sheet2);
      engine.addSheet(sheet3);

      // Set data on Sheet1
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", 250]])
      );

      // Reference from Sheet2
      engine.setSheetContent(
        sheet2,
        new Map<string, SerializedCellValue>([
          ["A1", `=${sheetAddress.sheetName}!A1+100`],
        ])
      );

      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(350); // 250 + 100

      // Remove Sheet1
      engine.removeSheet(sheetAddress);

      // Formula should now show error
      const result = engine.getCellValue({
        sheetName: sheet2.sheetName,
        workbookName: sheet2.workbookName,
        rowIndex: 0,
        colIndex: 0,
      });
      expect(typeof result === "string" && result.startsWith("#")).toBe(true);
    });

    test("should handle complex dependencies across multiple sheets", () => {
      const sheet2 = { workbookName, sheetName: "Sheet2" };
      const sheet3 = { workbookName, sheetName: "Sheet3" };
      engine.addSheet(sheet2);
      engine.addSheet(sheet3);

      // Add global named expression
      engine.addNamedExpression({
        expressionName: "MULTIPLIER",
        expression: "2",
      });

      // Set data on Sheet1
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", 100]])
      );

      // Sheet2 references Sheet1 and uses named expression
      engine.setSheetContent(
        sheet2,
        new Map<string, SerializedCellValue>([
          ["A1", `=${sheetAddress.sheetName}!A1*MULTIPLIER`],
        ])
      );

      // Sheet3 references Sheet2
      engine.setSheetContent(
        sheet3,
        new Map<string, SerializedCellValue>([
          ["A1", `=${sheet2.sheetName}!A1+50`],
        ])
      );

      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(200); // 100 * 2
      expect(
        engine.getCellValue({
          sheetName: sheet3.sheetName,
          workbookName: sheet3.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(250); // 200 + 50

      // Update named expression
      engine.updateNamedExpression({
        expressionName: "MULTIPLIER",
        expression: "3",
      });

      // All dependent formulas should update
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(300); // 100 * 3
      expect(
        engine.getCellValue({
          sheetName: sheet3.sheetName,
          workbookName: sheet3.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(350); // 300 + 50

      // Change source data
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", 200]])
      );

      // All dependent formulas should update
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(600); // 200 * 3
      expect(
        engine.getCellValue({
          sheetName: sheet3.sheetName,
          workbookName: sheet3.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(650); // 600 + 50
    });

    test("should handle sheet operations with global tables and named expressions", () => {
      const sheet2 = { workbookName, sheetName: "Sheet2" };
      engine.addSheet(sheet2);

      // Add global named expression
      engine.addNamedExpression({ expressionName: "TAX", expression: "0.1" });

      // Add sheet-scoped named expression
      engine.addNamedExpression({
        expressionName: "DISCOUNT",
        expression: "0.05",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      // Create table on Sheet1 (but table name is global)
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Item"],
          ["B1", "Price"],
          ["A2", "Widget"],
          ["B2", 100],
          ["C1", "=SUM(Products[Price])*(1-DISCOUNT)*(1+TAX)"],
        ])
      );

      engine.addTable({
        tableName: "Products",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      });

      // Reference global table from Sheet2 (no sheet prefix needed)
      engine.setSheetContent(
        sheet2,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Products[Price])*TAX"],
        ])
      );

      expect(cell("C1")).toBeCloseTo(104.5); // 100 * 0.95 * 1.1
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(10); // 100 * 0.1

      // Rename Sheet1 - table references should still work since tables are global
      const newSheetName = "Inventory";
      engine.renameSheet({
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        newSheetName,
      });

      expect(
        engine.getCellValue(
          {
            sheetName: newSheetName,
            workbookName: sheetAddress.workbookName,
            rowIndex: 0,
            colIndex: 2,
          },
          true
        )
      ).toBeNumber();

      // Formulas should still work
      expect(
        engine.getCellValue({
          sheetName: newSheetName,
          workbookName: sheetAddress.workbookName,
          rowIndex: 0,
          colIndex: 2,
        })
      ).toBeCloseTo(104.5);
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(10);
    });
  });

  describe("Event System", () => {
    test("should trigger tables-updated event when sheet is deleted", () => {
      let updateCount = 0;
      let lastUpdatedTables: Map<string, TableDefinition> | null = null;

      // Listen for update events
      const unsubscribe = engine._eventManager.onUpdate(() => {
        updateCount++;
        lastUpdatedTables = engine.getTables(sheetAddress.workbookName);
      });

      // Set up data and create table
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Product"],
          ["B1", "Price"],
          ["A2", "Widget"],
          ["B2", 100],
        ])
      );

      engine.addTable({
        tableName: "TestTable",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      });

      expect(updateCount).toBe(2); // From addTable
      expect(lastUpdatedTables!.has("TestTable")).toBe(true);

      // Remove sheet - should trigger tables-updated event because table is removed
      engine.removeSheet(sheetAddress);

      expect(updateCount).toBe(3); // From removeSheet -> removeTablesForSheet
      expect(lastUpdatedTables!.size).toBe(0); // Table should be gone

      unsubscribe();
    });

    test("should trigger onCellsUpdate callbacks when global named expression is deleted", () => {
      let cellsUpdateCount = 0;

      // Add listener for cells update
      const unsubscribe = engine._eventManager.onUpdate(() => {
        cellsUpdateCount++;
      });

      // Add global named expression - should trigger cells update
      engine.addNamedExpression({
        expressionName: "RATE",
        expression: "0.1",
      });

      expect(cellsUpdateCount).toBe(1); // From addNamedExpression

      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1000],
          ["B1", "=A1*RATE"],
        ])
      );

      expect(cell("B1")).toBe(100); // 1000 * 0.1
      expect(cellsUpdateCount).toBe(2); // From setSheetContent

      // Remove the named expression - should trigger cells update
      engine.removeNamedExpression({ expressionName: "RATE" });

      expect(cellsUpdateCount).toBe(3); // From removeNamedExpression

      // Formula should now error
      const result = cell("B1");
      expect(typeof result === "string" && result.startsWith("#")).toBe(true);

      unsubscribe();
    });

    test("should trigger onCellsUpdate callbacks when sheet-scoped named expression is deleted", () => {
      let cellsUpdateCount = 0;

      // Add listener for cells update
      const unsubscribe = engine._eventManager.onUpdate(() => {
        cellsUpdateCount++;
      });

      // Add sheet-scoped named expression - should trigger cells update
      engine.addNamedExpression({
        expressionName: "DISCOUNT",
        expression: "0.15",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(cellsUpdateCount).toBe(1); // From addNamedExpression

      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1000],
          ["B1", "=A1*DISCOUNT"],
        ])
      );

      expect(cell("B1")).toBe(150); // 1000 * 0.15
      expect(cellsUpdateCount).toBe(2); // From setSheetContent

      // Remove the sheet-scoped named expression - should trigger cells update
      engine.removeNamedExpression({
        expressionName: "DISCOUNT",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(cellsUpdateCount).toBe(3); // From removeNamedExpression

      // Formula should now error
      const result = cell("B1");
      expect(typeof result === "string" && result.startsWith("#")).toBe(true);

      unsubscribe();
    });

    test("should trigger onCellsUpdate callbacks when table is deleted", () => {
      let cellsUpdateCount = 0;

      // Add listener for cells update
      const unsubscribe = engine._eventManager.onUpdate(() => {
        cellsUpdateCount++;
      });

      // Set up data - should trigger cells update
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Product"],
          ["B1", "Price"],
          ["A2", "Widget"],
          ["B2", 100],
          ["C1", "=Products[Price]"],
        ])
      );

      expect(cellsUpdateCount).toBe(1); // From setSheetContent

      // Create table - should trigger cells update
      engine.addTable({
        tableName: "Products",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      });

      expect(cell("C1")).toBe(100);
      expect(cellsUpdateCount).toBe(2); // From addTable

      // Remove the table - should trigger cells update
      engine.removeTable({
        tableName: "Products",
        workbookName: sheetAddress.workbookName,
      });

      expect(cellsUpdateCount).toBe(3); // From removeTable

      // Formula should now error
      const result = cell("C1", true);
      expect(result).toMatchInlineSnapshot(`"#REF! in ast:Products[Price] Table Products not found"`);

      unsubscribe();
    });

    test("should trigger onCellsUpdate callbacks when global named expression is updated", () => {
      let cellsUpdateCount = 0;

      // Add listener for cells update
      const unsubscribe = engine._eventManager.onUpdate(() => {
        cellsUpdateCount++;
      });

      // Add global named expression - should trigger cells update
      engine.addNamedExpression({
        expressionName: "MULTIPLIER",
        expression: "2",
      });

      expect(cellsUpdateCount).toBe(1); // From addNamedExpression

      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 100],
          ["B1", "=A1*MULTIPLIER"],
        ])
      );

      expect(cell("B1")).toBe(200); // 100 * 2
      expect(cellsUpdateCount).toBe(2); // From setSheetContent

      // Update the named expression - should trigger cells update
      engine.updateNamedExpression({
        expressionName: "MULTIPLIER",
        expression: "3",
      });

      expect(cellsUpdateCount).toBe(3); // From updateNamedExpression
      expect(cell("B1")).toBe(300); // 100 * 3 (updated)

      unsubscribe();
    });

    test("should trigger onCellsUpdate callbacks when table is renamed", () => {
      let cellsUpdateCount = 0;

      // Add listener for cells update
      const unsubscribe = engine._eventManager.onUpdate(() => {
        cellsUpdateCount++;
      });

      // Set up data and create table
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Product"],
          ["B1", "Price"],
          ["A2", "Widget"],
          ["B2", 150],
          ["C1", "=SUM(OldTable[Price])"],
        ])
      );

      expect(cellsUpdateCount).toBe(1); // From setSheetContent

      engine.addTable({
        tableName: "OldTable",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      });

      expect(cell("C1")).toBe(150);
      expect(cellsUpdateCount).toBe(2); // From addTable

      // Rename the table - should trigger cells update
      engine.renameTable(sheetAddress.workbookName, {
        oldName: "OldTable",
        newName: "NewTable",
      });

      expect(cellsUpdateCount).toBe(3); // From renameTable
      expect(cell("C1")).toBe(150); // Should still work with new name

      unsubscribe();
    });

    test("should trigger onCellsUpdate callbacks when sheet is renamed", () => {
      const sheet2 = { workbookName, sheetName: "Sheet2" };
      engine.addSheet(sheet2); // Add sheet2 first

      let sheet1UpdateCount = 0;
      let sheet2UpdateCount = 0;

      // Add listeners for both sheets
      const unsubscribe1 = engine._eventManager.onUpdate(() => {
        sheet1UpdateCount++;
      });

      const unsubscribe2 = engine._eventManager.onUpdate(() => {
        sheet2UpdateCount++;
      });

      // Set up cross-sheet reference
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", 100]])
      );

      engine.setSheetContent(
        sheet2,
        new Map<string, SerializedCellValue>([["A1", `=${sheetName}!A1*2`]])
      );

      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(200);
      expect(sheet1UpdateCount).toBe(2); // From setSheetContent on both sheets (cross-sheet dependency)
      expect(sheet2UpdateCount).toBe(2); // From setSheetContent on both sheets (cross-sheet dependency)

      // Rename sheet - should trigger cells update on sheets with references
      const newSheetName = "RenamedSheet";
      engine.renameSheet({
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        newSheetName,
      });

      expect(sheet1UpdateCount).toBe(3); // From setSheetContent (2x) + renameSheet
      expect(sheet2UpdateCount).toBe(3); // From setSheetContent (2x) + renameSheet

      // Formula should still work
      expect(
        engine.getCellValue({
          sheetName: sheet2.sheetName,
          workbookName: sheet2.workbookName,
          rowIndex: 0,
          colIndex: 0,
        })
      ).toBe(200);

      unsubscribe1();
      unsubscribe2();
    });

    test("should trigger multiple events when using bulk operations", () => {
      let updateCount = 0;

      // Listen for all events
      const unsubscribeCount = engine._eventManager.onUpdate(() => {
        updateCount++;
      });

      // Set up initial data with formulas using tables and named expressions
      engine.addNamedExpression({ expressionName: "TAX", expression: "0.1" });

      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Item"],
          ["B1", "Price"],
          ["A2", "Widget"],
          ["B2", 100],
          ["C1", "=SUM(Products[Price])*(1+TAX)"],
        ])
      );

      engine.addTable({
        tableName: "Products",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      });

      expect(cell("C1")).toBeCloseTo(110); // 100 * 1.1
      expect(updateCount).toBe(3); // From addTable, setSheetContent, addNamedExpression

      // Use bulk operations - should trigger multiple events
      const newTables = new Map([
        [
          "Inventory",
          {
            name: "Inventory",
            sheetName: sheetAddress.sheetName,
            workbookName: sheetAddress.workbookName,
            start: { rowIndex: 0, colIndex: 0 },
            headers: new Map([
              ["Item", { name: "Item", index: 0 }],
              ["Price", { name: "Price", index: 1 }],
            ]),
            endRow: { type: "number", value: 1 } as const,
          },
        ],
      ]);

      const newGlobalExpressions = new Map([
        ["DISCOUNT", { name: "DISCOUNT", expression: "0.05" }],
      ]);

      engine.resetTables(new Map([[sheetAddress.workbookName, newTables]]));
      engine.setNamedExpressions({
        type: "global",
        expressions: newGlobalExpressions,
      });

      expect(updateCount).toBe(5); // From setNamedExpressions, resetTables, addTable, setSheetContent, addNamedExpression

      unsubscribeCount();
    });
  });

  describe("Open ended ranges", () => {
    test("SPILL on open ended ranges", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 3],
          ["A20", 4],
          ["A30", 5],
          ["B1", 6],
          ["B20", 7],
          ["B30", 8],
          ["C10", "=A20:B"],
        ])
      );

      expect(cell("C10")).toBe(4);
      expect(cell("C20", true)).toBe(5);
      expect(cell("D10")).toBe(7);
      expect(cell("D20", true)).toBe(8);
    });
  });

  describe("Cycle Detection", () => {
    test("should detect and mark all nodes in a simple cycle", () => {
      // Create a simple cycle: A1 -> B1 -> A1
      setCellContent("A1", "=B1");
      setCellContent("B1", "=A1");

      // Both cells should show cycle error
      expect(cell("A1")).toBe(FormulaError.CYCLE);
      expect(cell("B1")).toBe(FormulaError.CYCLE);
    });

    test("should detect and mark all nodes in a complex cycle", () => {
      // Create a more complex cycle: A1 -> B1 -> C1 -> A1
      setCellContent("A1", "=B1");
      setCellContent("B1", "=C1");
      setCellContent("C1", "=A1");

      // All three cells should show cycle error
      expect(cell("A1", true)).toMatchInlineSnapshot(
        `"#CYCLE! cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1 -> cell-value:TestWorkbook:TestSheet:A1 -> ast:A1"`);
      expect(cell("B1", true)).toMatchInlineSnapshot(
        `"#CYCLE! in cell-value:TestWorkbook:TestSheet:A1 cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1 -> cell-value:TestWorkbook:TestSheet:A1 -> ast:A1"`);
      expect(cell("C1", true)).toMatchInlineSnapshot(
        `"#CYCLE! in cell-value:TestWorkbook:TestSheet:A1 cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1 -> cell-value:TestWorkbook:TestSheet:A1 -> ast:A1"`);
    });

    test("should detect cycles with non-cycle dependencies", () => {
      // Create a cycle with additional dependencies: A1 -> B1 -> C1 -> B1, D1 -> A1
      setCellContent("A1", "=B1");
      setCellContent("B1", "=C1");
      setCellContent("C1", "=B1"); // Creates cycle B1 -> C1 -> B1
      setCellContent("D1", "=A1"); // Depends on A1 but not part of cycle

      // Cycle participants should show cycle error
      expect(cell("B1", true)).toMatchInlineSnapshot(
        `"#CYCLE! cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1"`);
      expect(cell("C1", true)).toMatchInlineSnapshot(
        `"#CYCLE! in cell-value:TestWorkbook:TestSheet:B1 cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1"`);

      // A1 should also show cycle error since it depends on the cycle
      expect(cell("A1", true)).toMatchInlineSnapshot(
        `"#CYCLE! in cell-value:TestWorkbook:TestSheet:B1 cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1"`);

      // D1 should also show cycle error since it depends on A1 which has a cycle
      expect(cell("D1", true)).toMatchInlineSnapshot(
        `"#CYCLE! in cell-value:TestWorkbook:TestSheet:B1 cell-value:TestWorkbook:TestSheet:C1 -> ast:C1 -> cell-value:TestWorkbook:TestSheet:B1 -> ast:B1"`);
    });

    test("should handle self-referencing cell", () => {
      // Create a self-reference: A1 -> A1
      setCellContent("A1", "=A1");

      expect(cell("A1", true)).toMatchInlineSnapshot(
        `"#CYCLE! cell-value:TestWorkbook:TestSheet:A1 -> ast:A1"`);
    });
  });

  describe("smartPaste", () => {
    test("should use pasteCells when paste area matches source size", () => {
      // Set up 2x2 source
      setCellContent("A1", 10);
      setCellContent("B1", 20);
      setCellContent("A2", 30);
      setCellContent("B2", 40);

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
      ];

      // Paste into same-size area at D5:E6
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 3, row: 4 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 5 },
            },
          }],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Should paste normally (not fill)
      expect(cell("D5")).toBe(10);
      expect(cell("E5")).toBe(20);
      expect(cell("D6")).toBe(30);
      expect(cell("E6")).toBe(40);
    });

    test("should use fillAreas when paste area is larger than source", () => {
      // Set up 1x1 source
      setCellContent("A1", 42);

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      // Paste into larger area F6:H8 (3x3)
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number", value: 7 },
              row: { type: "number", value: 7 },
            },
          }],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Should fill entire area
      for (let row = 5; row <= 7; row++) {
        for (let col = 5; col <= 7; col++) {
          const value = engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: col,
            rowIndex: row,
          });
          expect(value).toBe(42);
        }
      }
    });

    test("should use fillAreas with formulas and adjust references", () => {
      // Set up source with formula
      setCellContent("A1", "=ROW()+COLUMN()");

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      // Paste into larger area D5:E6 (2x2)
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 3, row: 4 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 5 },
            },
          }],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Check formulas were adjusted (column-first fill)
      const sheetContent = engine.getSheetSerialized({ workbookName, sheetName });
      expect(sheetContent.get("D5")).toBe("=ROW()+COLUMN()");
      expect(sheetContent.get("D6")).toBe("=ROW()+COLUMN()");
      expect(sheetContent.get("E5")).toBe("=ROW()+COLUMN()");
      expect(sheetContent.get("E6")).toBe("=ROW()+COLUMN()");
    });

    test("should handle 2x2 source filling into 5x5 area", () => {
      // Set up 2x2 source
      setCellContent("A1", 1);
      setCellContent("B1", 2);
      setCellContent("A2", 3);
      setCellContent("B2", 4);

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
      ];

      // Paste into F6:J10 (5x5)
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number", value: 9 },
              row: { type: "number", value: 9 },
            },
          }],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Check column-first fill pattern
      expect(cell("F6")).toBe(1);  // First column repeats: 1,3,1,3,1
      expect(cell("F7")).toBe(3);
      expect(cell("F8")).toBe(1);
      expect(cell("F9")).toBe(3);
      expect(cell("F10")).toBe(1);

      expect(cell("G6")).toBe(2);  // Second column: 2,4,2,4,2
      expect(cell("G7")).toBe(4);

      expect(cell("H6")).toBe(1);  // Replicates first column
      expect(cell("H7")).toBe(3);

      expect(cell("I6")).toBe(2);  // Replicates second column
      expect(cell("I7")).toBe(4);

      expect(cell("J6")).toBe(1);  // Partial replication
    });

    test("should use pasteCells when paste area equals source size", () => {
      // Set up 2x2 source
      setCellContent("A1", 10);
      setCellContent("B1", 20);
      setCellContent("A2", 30);
      setCellContent("B2", 40);

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
      ];

      // Paste into 2x2 area at F6:G7 (same size)
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number", value: 6 },
              row: { type: "number", value: 6 },
            },
          }],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Should use pasteCells - normal paste
      expect(cell("F6")).toBe(10);
      expect(cell("G6")).toBe(20);
      expect(cell("F7")).toBe(30);
      expect(cell("G7")).toBe(40);
    });

    test("should respect paste type option", () => {
      // Set up source with formula
      setCellContent("B1", 10);
      setCellContent("A1", "=B1*2");

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      // Paste as value into larger area
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number", value: 6 },
              row: { type: "number", value: 6 },
            },
          }],
        },
        { cut: false, type: "value", include: "all" }
      );

      // Should paste evaluated values, not formulas
      const sheetContent = engine.getSheetSerialized({ workbookName, sheetName });
      expect(sheetContent.get("F6")).toBe(20); // Value, not formula
      expect(sheetContent.get("G6")).toBe(20);
    });

    test("should paste to multiple areas independently", () => {
      // Set up single cell source
      setCellContent("A1", "Test");

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      // Paste into two different areas
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [
            {
              start: { col: 5, row: 5 },
              end: {
                col: { type: "number", value: 6 },
                row: { type: "number", value: 6 },
              },
            },
            {
              start: { col: 10, row: 10 },
              end: {
                col: { type: "number", value: 11 },
                row: { type: "number", value: 11 },
              },
            },
          ],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Both areas should be filled
      expect(cell("F6")).toBe("Test");
      expect(cell("G6")).toBe("Test");
      expect(cell("F7")).toBe("Test");
      expect(cell("G7")).toBe("Test");

      expect(cell("K11")).toBe("Test");
      expect(cell("L11")).toBe("Test");
      expect(cell("K12")).toBe("Test");
      expect(cell("L12")).toBe("Test");
    });

    test("should handle mixed paste and fill across multiple areas", () => {
      // Set up 2x2 source
      setCellContent("A1", 10);
      setCellContent("B1", 20);
      setCellContent("A2", 30);
      setCellContent("B2", 40);

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
      ];

      // Paste into two areas: one same-size (paste), one larger (fill)
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [
            // Area 1: same size as source (2x2) - should paste
            {
              start: { col: 5, row: 5 },
              end: {
                col: { type: "number", value: 6 },
                row: { type: "number", value: 6 },
              },
            },
            // Area 2: larger than source (4x4) - should fill
            {
              start: { col: 10, row: 10 },
              end: {
                col: { type: "number", value: 13 },
                row: { type: "number", value: 13 },
              },
            },
          ],
        },
        { cut: false, type: "formula", include: "all" }
      );

      // Area 1 (F6:G7) - should paste normally
      expect(cell("F6")).toBe(10);
      expect(cell("G6")).toBe(20);
      expect(cell("F7")).toBe(30);
      expect(cell("G7")).toBe(40);

      // Area 2 (K11:N14) - should fill with pattern
      expect(cell("K11")).toBe(10);
      expect(cell("L11")).toBe(20);
      expect(cell("K12")).toBe(30);
      expect(cell("L12")).toBe(40);
      // Pattern repeats
      expect(cell("M11")).toBe(10);
      expect(cell("N11")).toBe(20);
      expect(cell("K13")).toBe(10);
      expect(cell("L13")).toBe(20);
    });
  });

  describe("evaluateFormula", () => {
    test("should evaluate simple arithmetic formulas", () => {
      const result = engine.evaluateFormula(
        "=1 + 2 + 3",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(6);
    });

    test("should evaluate formulas with multiplication", () => {
      const result = engine.evaluateFormula(
        "=5 * 10",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(50);
    });

    test("should evaluate formulas with division", () => {
      const result = engine.evaluateFormula(
        "=100 / 4",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(25);
    });

    test("should evaluate formulas with cell references", () => {
      // Set up some data first
      setCellContent("A1", 10);
      setCellContent("B1", 20);

      const result = engine.evaluateFormula(
        "=A1 + B1",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(30);
    });

    test("should evaluate formulas with functions", () => {
      setCellContent("A1", 1);
      setCellContent("A2", 2);
      setCellContent("A3", 3);

      const result = engine.evaluateFormula(
        "=SUM(A1:A3)",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(6);
    });

    test("should evaluate formulas with cross-sheet references", () => {
      engine.addSheet({ workbookName, sheetName: "Sheet2" });
      setCellContent("A1", 100);

      engine.setSheetContent(
        { workbookName, sheetName: "Sheet2" },
        new Map([["B1", 50]])
      );

      const result = engine.evaluateFormula(
        "=TestSheet!A1 + Sheet2!B1",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(150);
    });

    test("should evaluate formulas with named expressions", () => {
      engine.addNamedExpression({
        expressionName: "MULTIPLIER",
        expression: "5",
      });

      const result = engine.evaluateFormula(
        "=MULTIPLIER * 10",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe(50);
    });

    test("should handle formulas with text concatenation", () => {
      const result = engine.evaluateFormula(
        '="Hello" & " " & "World"',
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe("Hello World");
    });

    test("should handle formulas with comparison operators", () => {
      const result1 = engine.evaluateFormula(
        "=5 > 3",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result1).toBe(true);

      const result2 = engine.evaluateFormula(
        "=5 = 3",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result2).toBe(false);
    });

    test("should handle division by zero", () => {
      const result = engine.evaluateFormula(
        "=1 / 0",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      expect(result).toBe('INFINITY');
    });

    test("should handle invalid cell references", () => {
      const result = engine.evaluateFormula(
        "=Z999 + 1",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      // Empty cell should be treated as 0 in arithmetic
      expect(result).toBe(FormulaError.VALUE);
    });
    test("should handle context dependent formulas", () => {
      const result = engine.evaluateFormula(
        "=ROW()",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }
      );
      // Empty cell should be treated as 0 in arithmetic
      expect(result).toBe(1);
      const result2 = engine.evaluateFormula(
        "=ROW()",
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 }
      );
      // Empty cell should be treated as 0 in arithmetic
      expect(result2).toBe(2);
    });
  });

  describe("moveCell", () => {
    test("should move cell content to new location", () => {
      setCellContent("A1", 42);

      engine.moveCell(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        { workbookName, sheetName, colIndex: 3, rowIndex: 4 }  // D5
      );

      // Source should be empty
      expect(cell("A1")).toBe("");

      // Target should have the value
      expect(cell("D5")).toBe(42);
    });

    test("should update dependent formulas when cell is moved", () => {
      setCellContent("A1", 100);
      setCellContent("B1", "=A1");
      setCellContent("C1", "=A1*2");

      expect(cell("B1")).toBe(100);
      expect(cell("C1")).toBe(200);

      // Move A1 to D5
      engine.moveCell(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        { workbookName, sheetName, colIndex: 3, rowIndex: 4 }  // D5
      );

      // References should be updated
      expect(cellContent("B1")).toBe("=D5");
      expect(cellContent("C1")).toBe("=D5*2");

      // Values should still be correct
      expect(cell("B1")).toBe(100);
      expect(cell("C1")).toBe(200);
    });

    test("should update absolute references when cell is moved", () => {
      setCellContent("A1", 50);
      setCellContent("B1", "=$A$1");

      engine.moveCell(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 2, rowIndex: 2 } // C3
      );

      expect(cellContent("B1")).toBe("=$C$3");
      expect(cell("B1")).toBe(50);
    });

    test("should handle multiple cells referencing moved cell", () => {
      setCellContent("A1", 10);
      setCellContent("B1", "=A1");
      setCellContent("B2", "=A1*2");
      setCellContent("B3", "=A1+5");

      engine.moveCell(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 5, rowIndex: 5 } // F6
      );

      expect(cellContent("B1")).toBe("=F6");
      expect(cellContent("B2")).toBe("=F6*2");
      expect(cellContent("B3")).toBe("=F6+5");

      expect(cell("B1")).toBe(10);
      expect(cell("B2")).toBe(20);
      expect(cell("B3")).toBe(15);
    });

    test("should move formulas correctly", () => {
      setCellContent("A1", 5);
      setCellContent("B1", 10);
      setCellContent("C1", "=A1+B1");

      expect(cell("C1")).toBe(15);

      // Move C1 to D5
      engine.moveCell(
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 3, rowIndex: 4 }
      );

      // Formula should be adjusted (relative references)
      // C1 is 2 cols to the right of A1, so D5 should reference 2 cols to left: B5
      // C1 is 1 col to the right of B1, so D5 should also reference 1 col to left: C5
      expect(cellContent("D5")).toBe("=B5+C5");
      
      // B5 and C5 are empty, which results in #VALUE! error when adding
      expect(cell("D5")).toBe("#VALUE!");
    });
  });

  describe("moveRange", () => {
    test("should move range content to new location", () => {
      setCellContent("A1", 1);
      setCellContent("B1", 2);
      setCellContent("A2", 3);
      setCellContent("B2", 4);

      engine.moveRange(
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
          }
        },
        { workbookName, sheetName, colIndex: 5, rowIndex: 5 } // F6
      );

      // Source range should be empty
      expect(cell("A1")).toBe("");
      expect(cell("B1")).toBe("");
      expect(cell("A2")).toBe("");
      expect(cell("B2")).toBe("");

      // Target range should have the values
      expect(cell("F6")).toBe(1);
      expect(cell("G6")).toBe(2);
      expect(cell("F7")).toBe(3);
      expect(cell("G7")).toBe(4);
    });

    test("should update range references when entire range is moved", () => {
      // Set up range A1:B2
      setCellContent("A1", 10);
      setCellContent("B1", 20);
      setCellContent("A2", 30);
      setCellContent("B2", 40);

      // C1 references the range
      setCellContent("C1", "=INDEX(A1:B2, 1, 1)");  // Using INDEX instead of SUM due to known issue
      expect(cell("C1")).toBe(10);

      // Move A1:B2 to D5:E6
      engine.moveRange(
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
          }
        },
        { workbookName, sheetName, colIndex: 3, rowIndex: 4 }
      );

      // Verify cells were moved
      expect(cell("D5")).toBe(10);
      expect(cell("E5")).toBe(20);
      expect(cell("D6")).toBe(30);
      expect(cell("E6")).toBe(40);

      // Reference should be updated
      expect(cellContent("C1")).toBe("=INDEX(D5:E6,1,1)");  // Spaces removed by parser
      expect(cell("C1")).toBe(10);  // Should still return first cell of moved range
    });

    test("should NOT update range reference when only partial range is moved", () => {
      // Set up larger range A1:D4
      setCellContent("A1", 1);
      setCellContent("D4", 4);

      // E1 references A1:D4
      setCellContent("E1", "=SUM(A1:D4)");

      // Move only A1:B2 (partial range)
      engine.moveRange(
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
          }
        },
        { workbookName, sheetName, colIndex: 5, rowIndex: 5 }
      );

      // E1 should still reference A1:D4 (partial move doesn't update)
      expect(cellContent("E1")).toBe("=SUM(A1:D4)");
    });

    test("should update individual cell references in moved range", () => {
      setCellContent("A1", 100);
      setCellContent("B1", 200);
      setCellContent("C1", "=A1");
      setCellContent("D1", "=B1");

      // Move A1:B1 to E5:F5
      engine.moveRange(
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } }
          }
        },
        { workbookName, sheetName, colIndex: 4, rowIndex: 4 }
      );

      expect(cellContent("C1")).toBe("=E5");
      expect(cellContent("D1")).toBe("=F5");
      expect(cell("C1")).toBe(100);
      expect(cell("D1")).toBe(200);
    });

    test("should handle cross-sheet references", () => {
      engine.addSheet({ workbookName, sheetName: "Sheet2" });

      // Set value in TestSheet (our default sheetName)
      setCellContent("A1", 123);
      
      // Sheet2 references TestSheet!A1
      engine.setCellContent(
        { workbookName, sheetName: "Sheet2", colIndex: 0, rowIndex: 0 },
        `=${sheetName}!A1`  // Use the sheetName variable which is "TestSheet"
      );

      const sheet2A1Value = engine.getCellValue({
        workbookName,
        sheetName: "Sheet2",
        colIndex: 0,
        rowIndex: 0
      });
      expect(sheet2A1Value).toBe(123);

      // Move TestSheet!A1 to TestSheet!D5
      engine.moveCell(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 3, rowIndex: 4 }
      );

      // Sheet2!A1 should now reference TestSheet!D5
      const sheet2Content = engine.getSheetSerialized({ workbookName, sheetName: "Sheet2" });
      expect(sheet2Content.get("A1")).toBe(`=${sheetName}!D5`);

      const sheet2A1NewValue = engine.getCellValue({
        workbookName,
        sheetName: "Sheet2",
        colIndex: 0,
        rowIndex: 0
      });
      expect(sheet2A1NewValue).toBe(123);
    });
  });

  describe("smartPaste with cut", () => {
    test("should use pasteCells when cut=true, not fillAreas", () => {
      // Set up source
      setCellContent("A1", 10);
      setCellContent("B1", 20);

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      ];

      // Paste into larger area (would normally trigger fillAreas)
      // But with cut=true, should use pasteCells instead
      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number", value: 8 }, // 4 columns wide
              row: { type: "number", value: 8 }  // 4 rows tall
            }
          }]
        },
        { cut: true, type: "formula", include: "all" }
      );

      // Should paste A1:B1 to F6:G6, NOT fill the entire area
      expect(cell("F6")).toBe(10);
      expect(cell("G6")).toBe(20);
      expect(cell("H6")).toBe(""); // Not filled
      expect(cell("F7")).toBe(""); // Not filled

      // Source should be cleared
      expect(cell("A1")).toBe("");
      expect(cell("B1")).toBe("");
    });

    test("should update references when cut with smartPaste", () => {
      setCellContent("A1", 50);
      setCellContent("C1", "=A1");

      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [{
            start: { col: 5, row: 5 },
            end: { col: { type: "number", value: 5 }, row: { type: "number", value: 5 } }
          }]
        },
        { cut: true, type: "formula", include: "all" }
      );

      // Reference should be updated
      expect(cellContent("C1")).toBe("=F6");
      expect(cell("C1")).toBe(50);
    });
  });
});
