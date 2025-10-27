import { beforeEach, describe, expect, test } from "bun:test";
import {
  FormulaError,
  type SerializedCellValue,
  type TableDefinition,
} from "../../../src/core/types";
import { getCellReference, parseCellReference } from "../../../src/core/utils";
import { FormulaEngine } from "../../../src/core/engine";
import { visualizeSpreadsheet } from "../../../src/core/utils/spreadsheet-visualizer";

describe("Reproduce issue with evalution order", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "TestSheet";

  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

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

  test("should resolve correctly", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(10)"],
        ["B1", "=A:A"],
        ["C1", "=B:B"],
        ["D1", "=C:C"],
        ["E1", "=D:D"],
        ["F1", "=E:E"],
        ["G1", "=INDEX(F:F, MATCH(5, F:F, 0))"],
      ])
    );

    expect(cell("G1", true)).toBe(5);
  });

  test("complex cross-workbook and cross-sheet MATCH with spilling and concatenation", () => {
    // Create two workbooks
    const workbook1Name = "Workbook1";
    const workbook2Name = "Workbook2";

    // Create a new engine for this test
    const testEngine = FormulaEngine.buildEmpty();
    testEngine.addWorkbook(workbook1Name);
    testEngine.addWorkbook(workbook2Name);

    // Add sheets to Workbook1: Sheet1, Sheet2, Sheet3
    testEngine.addSheet({ workbookName: workbook1Name, sheetName: "Sheet1" });
    testEngine.addSheet({ workbookName: workbook1Name, sheetName: "Sheet2" });
    testEngine.addSheet({ workbookName: workbook1Name, sheetName: "Sheet3" });

    // Add Sheet1 to Workbook2
    testEngine.addSheet({ workbookName: workbook2Name, sheetName: "Sheet1" });

    // Step 4: Set up [Workbook2]Sheet1 with some data
    // A1 has "some_", B1 has "value", C1 has concatenation of A1 and B1
    testEngine.setSheetContent(
      { workbookName: workbook2Name, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([
        ["A1", "some_"],
        ["B1", "value"],
        ["C1", "=A1&B1"], // "some_value"
        ["A2", "some_other_"],
        ["B2", "value"],
        ["C2", "=A2&B2"], // "some_other_value"
      ])
    );

    // Step 3: Sheet3 cell A1 has a reference to [Workbook2]Sheet1!A:INFINITY which spills
    // This creates a spilled range that includes all columns from A onwards
    testEngine.setSheetContent(
      { workbookName: workbook1Name, sheetName: "Sheet3" },
      new Map<string, SerializedCellValue>([
        ["A1", "=[Workbook2]Sheet1!A:INFINITY"], // This will spill across all columns
      ])
    );

    // Step 2: Sheet2 cell A1 has MATCH looking for "some_other_value" in Sheet3!C:C
    testEngine.setSheetContent(
      { workbookName: workbook1Name, sheetName: "Sheet2" },
      new Map<string, SerializedCellValue>([
        ["A1", '=MATCH("some_other_value", Sheet3!C:C, 0)'],
      ])
    );

    // Step 1: Sheet1 cell A1 has MATCH looking for "some_value" in Sheet2!A:A
    testEngine.setSheetContent(
      { workbookName: workbook1Name, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([["A1", "=MATCH(2, Sheet2!A:A, 0)"]])
    );

    // Helper function to get cell value
    const getCellValue = (workbook: string, sheet: string, ref: string) =>
      testEngine.getCellValue({
        workbookName: workbook,
        sheetName: sheet,
        ...parseCellReference(ref),
      });

    // Sheet1!A1 should find value 2 at position 1 in Sheet2!A:A
    expect(getCellValue(workbook1Name, "Sheet1", "A1")).toBe(1);
  });

  test("SUM, AVERAGE, COUNT on spilled range with aggregation", () => {
    // Create a spilled range using SEQUENCE
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // A1 will spill values 1, 2, 3, 4, 5 down column A
        ["A1", "=C1:C5"],

        // B1: SUM the spilled range in column A
        ["B1", "=SUM(A1:A5)"],

        // B2: AVERAGE the spilled range in column A
        ["B2", "=AVERAGE(A1:A5)"],

        // B3: COUNT the spilled range in column A
        ["B3", "=COUNT(A1:A5)"],

        // B4: SUM all the results from B1, B2, B3
        ["B4", "=SUM(B1:B3)"],

        ["C1", 1],
        ["C2", 2],
        ["C3", 3],
        ["C4", 4],
        ["C5", 5],
      ])
    );

    const value = cell("B4");

    expect(
      engine._dependencyManager.getDependencyTree(
        engine._dependencyManager.getCellValueNode(`cell-value:TestWorkbook:TestSheet:B4`)
      )
    ).toMatchSnapshot();
    // B4: SUM(15, 3, 5) = 23
    expect(value).toBe(23);

    // Verify the spilled values
    expect(cell("A1")).toBe(1);
    expect(cell("A2")).toBe(2);
    expect(cell("A3")).toBe(3);
    expect(cell("A4")).toBe(4);
    expect(cell("A5")).toBe(5);

    // Verify the calculations
    // B1: SUM(1, 2, 3, 4, 5) = 15
    expect(cell("B1")).toBe(15);

    // B2: AVERAGE(1, 2, 3, 4, 5) = 3
    expect(cell("B2")).toBe(3);

    // B3: COUNT(1, 2, 3, 4, 5) = 5
    expect(cell("B3")).toBe(5);
  });

  test("SUM, AVERAGE, COUNT on spilled range with aggregation /2", () => {
    // Create a spilled range using SEQUENCE
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=C1:C2"],
        ["B1", "=SUM(A1:A2)"],
        ["B2", 5],
        ["D1", "=SUM(B1:B2)"],
        ["C1", "=SEQUENCE(2)"],
      ])
    );

    const value = cell("D1");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 5,
        numCols: 5,
        sheetName: sheetAddress.sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C   | D   | E  
      -----+-----+-----+-----+-----+----
         1 | 1   | 3   | 1   | 8   |    
         2 | 2   | 5   | 2   |     |    
         3 |     |     |     |     |    
         4 |     |     |     |     |    
         5 |     |     |     |     |    
      "
    `);

    expect(value).toBe(1 + 2 + 5);
  });

  test("SUM, AVERAGE, COUNT on cross-sheet spilled range with aggregation", () => {
    // Create a new engine with multiple sheets for this test
    const testEngine = FormulaEngine.buildEmpty();
    testEngine.addWorkbook(workbookName);
    testEngine.addSheet({ workbookName, sheetName: "Sheet1" });
    testEngine.addSheet({ workbookName, sheetName: "Sheet2" });

    const getCellValue = (sheet: string, ref: string) =>
      testEngine.getCellValue({
        workbookName,
        sheetName: sheet,
        ...parseCellReference(ref),
      });

    // Sheet1: Source data with SEQUENCE
    testEngine.setSheetContent(
      { workbookName, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(10)"], // Spills 1-10 down column A
      ])
    );

    // Sheet2: Reference Sheet1's spilled range and do calculations
    testEngine.setSheetContent(
      { workbookName, sheetName: "Sheet2" },
      new Map<string, SerializedCellValue>([
        // Reference the spilled range from Sheet1
        ["A1", "=Sheet1!A:A"],

        // B1: SUM the cross-sheet spilled range
        ["B1", "=SUM(A:A)"],

        // B2: AVERAGE the cross-sheet spilled range
        ["B2", "=AVERAGE(A:A)"],

        // B3: COUNT the cross-sheet spilled range
        ["B3", "=COUNT(A:A)"],

        // B4: SUM all the results
        ["B4", "=SUM(B1:B3)"],
      ])
    );

    // B4: SUM(55, 5.5, 10) = 70.5
    expect(getCellValue("Sheet2", "B4")).toBe(70.5);

    // Verify Sheet1 spilled values
    expect(getCellValue("Sheet1", "A1")).toBe(1);
    expect(getCellValue("Sheet1", "A10")).toBe(10);

    // Verify Sheet2 has the spilled values from Sheet1
    expect(getCellValue("Sheet2", "A1")).toBe(1);
    expect(getCellValue("Sheet2", "A10")).toBe(10);

    // Verify the calculations on Sheet2
    // B1: SUM(1 to 10) = 55
    expect(getCellValue("Sheet2", "B1")).toBe(55);

    // B2: AVERAGE(1 to 10) = 5.5
    expect(getCellValue("Sheet2", "B2")).toBe(5.5);

    // B3: COUNT(1 to 10) = 10
    expect(getCellValue("Sheet2", "B3")).toBe(10);
  });
});
