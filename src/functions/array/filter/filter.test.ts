import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("FILTER function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const setContent = (content: Array<[string, SerializedCellValue]>) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(content)
    );
  };

  /** Reads a block of cells as rows of values. */
  const block = (start: string, rows: number, cols: number) => {
    const { colIndex, rowIndex } = parseCellReference(start);
    return Array.from({ length: rows }, (_, y) =>
      Array.from({ length: cols }, (_, x) =>
        engine.getCellValue({
          sheetName,
          workbookName,
          colIndex: colIndex + x,
          rowIndex: rowIndex + y,
        })
      )
    );
  };

  /** Regions in A2:A6, units in B2:B6 and prices in C2:C6. */
  const sales: Array<[string, SerializedCellValue]> = [
    ["A1", "Region"],
    ["B1", "Units"],
    ["C1", "Price"],
    ["A2", "North"],
    ["B2", 10],
    ["C2", 1.5],
    ["A3", "South"],
    ["B3", 20],
    ["C3", 2.5],
    ["A4", "North"],
    ["B4", 30],
    ["C4", 3.5],
    ["A5", "East"],
    ["B5", 40],
    ["C5", 4.5],
    ["A6", "North"],
    ["B6", 50],
    ["C6", 5.5],
  ];

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("selecting rows", () => {
    test("returns the rows whose include value is true", () => {
      setContent([...sales, ["E1", '=FILTER(A2:C6, A2:A6="North")']]);

      expect(block("E1", 3, 3)).toEqual([
        ["North", 10, 1.5],
        ["North", 30, 3.5],
        ["North", 50, 5.5],
      ]);
      expect(cell("E4")).toBe("");
    });

    test("filters one column by another", () => {
      setContent([...sales, ["E1", "=FILTER(B2:B6, C2:C6>3)"]]);

      expect(block("E1", 3, 1)).toEqual([[30], [40], [50]]);
    });

    test("follows edits of the data and the condition", () => {
      setContent([
        ...sales,
        ["G1", "North"],
        ["E1", "=FILTER(B2:B6, A2:A6=G1)"],
      ]);
      expect(block("E1", 3, 1)).toEqual([[10], [30], [50]]);

      setCellContent("A3", "North");
      expect(block("E1", 4, 1)).toEqual([[10], [20], [30], [50]]);

      setCellContent("G1", "East");
      expect(cell("E1")).toBe(40);
      expect(cell("E2")).toBe("");

      setCellContent("B5", 45);
      expect(cell("E1")).toBe(45);
    });

    test("returns a one-cell selection as a single value", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A6="South")'],
        ["E2", "=FILTER(5, TRUE)"],
      ]);

      expect(cell("E1")).toBe(20);
      expect(cell("E2")).toBe(5);
    });

    test("tests include values like IF conditions", () => {
      setContent([
        ...sales,
        // 1 and "x" are true; 0, FALSE and "" are false
        ["E1", '=FILTER(B2:B6, {1;0;"x";FALSE;""})'],
      ]);

      expect(block("E1", 2, 1)).toEqual([[10], [30]]);
    });

    test("does not select rows whose include cell is empty", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", true],
        // B2 is empty
        ["B3", 1],
        ["D1", "=FILTER(A1:A3, B1:B3)"],
      ]);

      expect(block("D1", 2, 1)).toEqual([[1], [3]]);
    });

    test("combines conditions with IF", () => {
      setContent([
        ...sales,
        // North and more than 20 units
        ["E1", '=FILTER(A2:B6, IF(A2:A6="North", B2:B6>20))'],
        // South or East
        ["E5", '=FILTER(A2:A6, IF(A2:A6="South", TRUE, A2:A6="East"))'],
      ]);

      expect(block("E1", 2, 2)).toEqual([
        ["North", 30],
        ["North", 50],
      ]);
      expect(block("E5", 2, 1)).toEqual([["South"], ["East"]]);
    });

    test("filters open-ended ranges by their populated cells", () => {
      setContent([...sales, ["E1", '=FILTER(B:B, A:A="North")']]);
      expect(block("E1", 3, 1)).toEqual([[10], [30], [50]]);

      setCellContent("A9", "North");
      setCellContent("B9", 90);
      expect(block("E1", 4, 1)).toEqual([[10], [30], [50], [90]]);
    });

    test("filters a computed array", () => {
      setContent([
        ["A1", "=FILTER(SEQUENCE(6), SEQUENCE(6)>4)"],
        ["C1", "=FILTER(SEQUENCE(3, 2), {TRUE;FALSE;TRUE})"],
      ]);

      expect(block("A1", 2, 1)).toEqual([[5], [6]]);
      expect(block("C1", 2, 2)).toEqual([
        [1, 2],
        [5, 6],
      ]);
    });
  });

  describe("selecting columns", () => {
    test("returns the columns whose include value is true", () => {
      setContent([...sales, ["E1", "=FILTER(A1:C3, {TRUE,FALSE,TRUE})"]]);

      expect(block("E1", 3, 2)).toEqual([
        ["Region", "Price"],
        ["North", 1.5],
        ["South", 2.5],
      ]);
    });

    test("filters columns with a condition on a header row", () => {
      setContent([
        ...sales,
        ['E1', '=FILTER(A1:C6, A1:C1<>"Units")'],
        ["H1", '=SUM(FILTER(A2:C6, A1:C1="Price"))'],
      ]);

      expect(block("E1", 2, 2)).toEqual([
        ["Region", "Price"],
        ["North", 1.5],
      ]);
      expect(cell("H1")).toBe(17.5);
    });

    test("filters the columns of an open-ended range", () => {
      setContent([
        ...sales,
        ["E1", "=SUM(FILTER(A:C, {FALSE,TRUE,FALSE}))"],
      ]);

      expect(cell("E1")).toBe(150);
    });
  });

  describe("empty selections", () => {
    test("returns #CALC! without if_empty", () => {
      setContent([...sales, ["E1", '=FILTER(B2:B6, A2:A6="West")']]);

      expect(cell("E1")).toBe(FormulaError.CALC);
    });

    test("returns if_empty when nothing is selected", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A6="West", "none")'],
        ["E2", '=FILTER(B2:B6, A2:A6="West", 0)'],
        ["E3", '=FILTER(B2:B6, A2:A6="West", {1,2})'],
      ]);

      expect(cell("E1")).toBe("none");
      expect(cell("E2")).toBe(0);
      expect(block("E3", 1, 2)).toEqual([[1, 2]]);
    });

    test("evaluates if_empty only when nothing is selected", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A6="South", NOSUCHFN())'],
        ["E2", '=FILTER(B2:B6, A2:A6="West", NOSUCHFN())'],
      ]);

      expect(cell("E1")).toBe(20);
      expect(cell("E2")).toBe(FormulaError.NAME);
    });

    test("keeps the default for an omitted if_empty", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A6="West",)'],
        ["E2", '=FILTER(B2:B6, A2:A6="South",)'],
      ]);

      expect(cell("E1")).toBe(FormulaError.CALC);
      expect(cell("E2")).toBe(20);
    });

    test("switches between a result and if_empty as the data changes", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A6="West", "none")'],
      ]);
      expect(cell("E1")).toBe("none");

      setCellContent("A5", "West");
      expect(cell("E1")).toBe(40);
    });
  });

  describe("errors", () => {
    test("returns an error in include", () => {
      setContent([
        ...sales,
        ["E1", "=FILTER(B2:B6, {TRUE;FALSE;NOSUCHFN();FALSE;TRUE})"],
        ["F1", '=FILTER(B2:B6, B2:B6/{1;1;0;1;"x"}>0)'],
      ]);

      expect(cell("E1")).toBe(FormulaError.NAME);
      expect(cell("F1")).toBe(FormulaError.VALUE);
    });

    test("keeps errors of the array in their cells", () => {
      setContent([
        ["A1", 1],
        ["A2", '=1+"x"'],
        ["A3", 3],
        ["D1", "=FILTER(A1:A3, {TRUE;TRUE;FALSE})"],
        ["E1", "=FILTER(A1:A3, {TRUE;FALSE;TRUE})"],
      ]);

      expect(block("D1", 2, 1)).toEqual([[1], [FormulaError.VALUE]]);
      expect(block("E1", 2, 1)).toEqual([[1], [3]]);
    });

    test("returns errors of single-value arguments", () => {
      setContent([
        ["A1", "=FILTER(NOSUCHFN(), TRUE)"],
        ["A2", '=FILTER(B1:B2, 1+"x")'],
      ]);

      expect(cell("A1")).toBe(FormulaError.NAME);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("returns #VALUE! when include does not fit the array", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A5="North")'],
        ["E2", '=FILTER(A2:C6, A2:C6="North")'],
        ["E3", "=FILTER(A2:C6, {TRUE,FALSE})"],
        ["E4", '=FILTER(B:B, A2:A6="North")'],
      ]);

      expect(cell("E1")).toBe(FormulaError.VALUE);
      expect(cell("E2")).toBe(FormulaError.VALUE);
      expect(cell("E3")).toBe(FormulaError.VALUE);
      expect(cell("E4")).toBe(FormulaError.VALUE);
    });

    test("returns #VALUE! for conditions combined by multiplication", () => {
      // Booleans are not numbers in this engine, so TRUE*TRUE is #VALUE!.
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, (A2:A6="North")*(B2:B6>20))'],
      ]);

      expect(cell("E1")).toBe(FormulaError.VALUE);
    });

    test("requires array and include", () => {
      setContent([
        ["A1", "=FILTER(B1:B3)"],
        ["A2", "=FILTER(, {TRUE})"],
        ["A3", "=FILTER(B1:B3,)"],
        ["A4", "=FILTER(B1:B3, TRUE, 1, 2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
      expect(cell("A4")).toBe(FormulaError.VALUE);
    });
  });

  describe("consumers", () => {
    test("is read whole by aggregates", () => {
      setContent([
        ...sales,
        ["E1", '=SUM(FILTER(B2:B6, A2:A6="North"))'],
        ["E2", '=AVERAGE(FILTER(B2:B6, A2:A6="North"))'],
        ["E3", '=STDEV.S(FILTER(B2:B6, A2:A6="North"))'],
        ["E4", '=COUNTA(FILTER(A2:C6, A2:A6="North"))'],
        ["E5", '=ROWS(FILTER(A2:C6, A2:A6="North"))'],
        ["E6", '=MAX(FILTER(C2:C6, A2:A6<>"North"))'],
      ]);

      expect(cell("E1")).toBe(90);
      expect(cell("E2")).toBe(30);
      expect(cell("E3")).toBe(20);
      expect(cell("E4")).toBe(9);
      expect(cell("E5")).toBe(3);
      expect(cell("E6")).toBe(4.5);
    });

    test("leaves the empty cells of a selected range out of aggregates", () => {
      setContent([
        ["A1", "x"],
        ["A2", "x"],
        ["A3", "x"],
        ["B1", 1],
        // B2 is empty
        ["B3", 3],
        ["D1", '=COUNTA(FILTER(B1:B3, A1:A3="x"))'],
        ["D2", '=AVERAGE(FILTER(B1:B3, A1:A3="x"))'],
      ]);

      expect(cell("D1")).toBe(2);
      expect(cell("D2")).toBe(2);
    });

    test("feeds other array functions", () => {
      setContent([
        ...sales,
        ["E1", '=UNIQUE(FILTER(A2:A6, B2:B6>15))'],
        ["G1", '=TRANSPOSE(FILTER(B2:B6, A2:A6="North"))'],
        ["G3", '=INDEX(FILTER(A2:C6, A2:A6="North"), 2, 3)'],
      ]);

      expect(block("E1", 3, 1)).toEqual([["South"], ["North"], ["East"]]);
      expect(block("G1", 1, 3)).toEqual([[10, 30, 50]]);
      expect(cell("G3")).toBe(3.5);
    });

    test("returns #SPILL! when the spill area is blocked", () => {
      setContent([
        ...sales,
        ["E1", '=FILTER(B2:B6, A2:A6="North")'],
        ["E2", "block"],
      ]);

      expect(cell("E1")).toBe(FormulaError.SPILL);

      setCellContent("E2", undefined);
      expect(block("E1", 3, 1)).toEqual([[10], [30], [50]]);
    });
  });
});
