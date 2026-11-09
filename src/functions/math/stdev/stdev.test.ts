import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("STDEV.S function", () => {
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

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  test("returns the sample standard deviation", () => {
    setContent([
      ["A1", "=STDEV.S(2, 4, 4, 4, 5, 5, 7, 9)"],
      ["A2", "=STDEV.S(1, 2, 3, 4)"],
      ["A3", "=STDEV.S(5, 5, 5)"],
      ["A4", "=STDEV.S(-1, 1)"],
    ]);

    expect(cell("A1")).toBeCloseTo(Math.sqrt(32 / 7), 14);
    expect(cell("A2")).toBeCloseTo(Math.sqrt(5 / 3), 14);
    expect(cell("A3")).toBe(0);
    expect(cell("A4")).toBeCloseTo(Math.SQRT2, 14);
  });

  test("matches Excel's result for a range of measurements", () => {
    // Excel: STDEV.S(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299)
    // is 27.46392...
    setContent([
      ["A1", 1345],
      ["A2", 1301],
      ["A3", 1368],
      ["A4", 1322],
      ["A5", 1310],
      ["A6", 1370],
      ["A7", 1318],
      ["A8", 1350],
      ["A9", 1303],
      ["A10", 1299],
      ["B1", "=STDEV.S(A1:A10)"],
    ]);

    expect(cell("B1")).toBeCloseTo(27.463915719843, 10);
  });

  test("stays accurate for numbers with a large common offset", () => {
    setContent([
      ["A1", "=STDEV.S(1000000004, 1000000007, 1000000013, 1000000016)"],
    ]);

    expect(cell("A1")).toBeCloseTo(Math.sqrt(30), 9);
  });

  test("combines arguments, ranges and spilled results", () => {
    setContent([
      ["A1", 2],
      ["A2", 4],
      ["B1", 4],
      ["B2", 4],
      ["C1", "=STDEV.S(A1:A2, B1:B2, 5, 5, {7, 9})"],
      ["C2", "=STDEV.S(SEQUENCE(4))"],
    ]);

    expect(cell("C1")).toBeCloseTo(Math.sqrt(32 / 7), 14);
    expect(cell("C2")).toBeCloseTo(Math.sqrt(5 / 3), 14);
  });

  test("skips text, logical values and empty cells", () => {
    setContent([
      ["A1", 1],
      ["A2", "text"],
      ["A3", true],
      // A4 is empty
      ["A5", 3],
      ["A6", '=""'],
      ["B1", "=STDEV.S(A1:A6)"],
      ["B2", '=STDEV.S(1, 3, TRUE, "5")'],
    ]);

    expect(cell("B1")).toBeCloseTo(Math.SQRT2, 14);
    // Excel would count the directly supplied TRUE and "5"
    expect(cell("B2")).toBeCloseTo(Math.SQRT2, 14);
  });

  test("recalculates when a value changes", () => {
    setContent([
      ["A1", 1],
      ["A2", 3],
      ["B1", "=STDEV.S(A1:A3)"],
    ]);
    expect(cell("B1")).toBeCloseTo(Math.SQRT2, 14);

    setCellContent("A3", 5);
    expect(cell("B1")).toBe(2);
  });

  test("returns #DIV/0! for fewer than two numbers", () => {
    setContent([
      ["A1", 7],
      ["A2", "x"],
      ["B1", "=STDEV.S(A1)"],
      ["B2", "=STDEV.S(A1:A3)"],
      ["B3", "=STDEV.S(A2:A3)"],
      ["B4", "=STDEV.S(5)"],
    ]);

    expect(cell("B1")).toBe(FormulaError.DIV0);
    expect(cell("B2")).toBe(FormulaError.DIV0);
    expect(cell("B3")).toBe(FormulaError.DIV0);
    expect(cell("B4")).toBe(FormulaError.DIV0);
  });

  test("propagates errors, also from inside a range", () => {
    setContent([
      ["A1", 1],
      ["A2", '=1+"x"'],
      ["A3", 3],
      ["B1", "=STDEV.S(A1:A3)"],
      ["B2", "=STDEV.S(1, 2, NOSUCHFN())"],
    ]);

    expect(cell("B1")).toBe(FormulaError.VALUE);
    expect(cell("B2")).toBe(FormulaError.NAME);
  });

  test("returns #NUM! for an infinity", () => {
    setContent([
      ["A1", "=STDEV.S(1, 2, INFINITY)"],
      ["A2", "=STDEV.S(1, 1/0)"],
    ]);

    expect(cell("A1")).toBe(FormulaError.NUM);
    expect(cell("A2")).toBe(FormulaError.NUM);
  });

  test("works over an open-ended range", () => {
    setContent([
      ["A1", 2],
      ["A2", 4],
      ["A3", 4],
      ["A4", 4],
      ["A5", 5],
      ["A6", 5],
      ["A7", 7],
      ["A8", 9],
      ["B1", "=STDEV.S(A:A)"],
    ]);

    expect(cell("B1")).toBeCloseTo(Math.sqrt(32 / 7), 14);
  });

  test("is also available under its older name STDEV", () => {
    setContent([["A1", "=STDEV(1, 2, 3, 4)"]]);

    expect(cell("A1")).toBeCloseTo(Math.sqrt(5 / 3), 14);
  });

  test("requires an argument", () => {
    setContent([["A1", "=STDEV.S()"]]);

    expect(cell("A1")).toBe(FormulaError.VALUE);
  });
});
