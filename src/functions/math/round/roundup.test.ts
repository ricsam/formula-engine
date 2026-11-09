import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ROUNDUP function", () => {
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

  test("rounds away from zero", () => {
    setContent([
      ["A1", "=ROUNDUP(3.2, 0)"],
      ["A2", "=ROUNDUP(76.9, 0)"],
      ["A3", "=ROUNDUP(3.14159, 3)"],
      ["A4", "=ROUNDUP(-3.14159, 1)"],
      ["A5", "=ROUNDUP(31415.92654, -2)"],
      ["A6", "=ROUNDUP(0.001, 0)"],
      ["A7", "=ROUNDUP(-0.001, 0)"],
    ]);

    expect(cell("A1")).toBe(4);
    expect(cell("A2")).toBe(77);
    expect(cell("A3")).toBe(3.142);
    expect(cell("A4")).toBe(-3.2);
    expect(cell("A5")).toBe(31500);
    expect(cell("A6")).toBe(1);
    expect(cell("A7")).toBe(-1);
  });

  test("leaves numbers that already fit unchanged", () => {
    setContent([
      ["A1", "=ROUNDUP(5, 0)"],
      ["A2", "=ROUNDUP(2.5, 1)"],
      ["A3", "=ROUNDUP(-1200, -2)"],
      ["A4", "=ROUNDUP(0, 2)"],
    ]);

    expect(cell("A1")).toBe(5);
    expect(cell("A2")).toBe(2.5);
    expect(cell("A3")).toBe(-1200);
    expect(cell("A4")).toBe(0);
  });

  test("works on the decimal value, as Excel does", () => {
    // In binary, 0.1+0.2 is 0.30000000000000004 and 1.1*1.1 is
    // 1.2100000000000002, which a binary ceiling would round up.
    setContent([
      ["A1", "=ROUNDUP(0.1+0.2, 1)"],
      ["A2", "=ROUNDUP(1.1*1.1, 2)"],
      ["A3", "=ROUNDUP(0.30000000000001, 1)"],
    ]);

    expect(cell("A1")).toBe(0.3);
    expect(cell("A2")).toBe(1.21);
    expect(cell("A3")).toBe(0.4);
  });

  test("rounds a power of ten up for a rounding position above the number", () => {
    setContent([
      ["A1", "=ROUNDUP(5, -1)"],
      ["A2", "=ROUNDUP(5, -3)"],
      ["A3", "=ROUNDUP(-0.5, -2)"],
    ]);

    expect(cell("A1")).toBe(10);
    expect(cell("A2")).toBe(1000);
    expect(cell("A3")).toBe(-100);
  });

  test("returns #NUM! when the result is too large for a number", () => {
    setContent([
      ["A1", "=ROUNDUP(5, -400)"],
      ["A2", "=ROUNDUP(5, -INFINITY)"],
      ["A3", "=ROUNDUP(0, -INFINITY)"],
    ]);

    expect(cell("A1")).toBe(FormulaError.NUM);
    expect(cell("A2")).toBe(FormulaError.NUM);
    expect(cell("A3")).toBe(0);
  });

  test("accepts only numbers and propagates errors", () => {
    setContent([
      ["A1", '=ROUNDUP("3.2", 0)'],
      ["A2", "=ROUNDUP(TRUE, 0)"],
      ["A3", "=ROUNDUP(B1, 0)"],
      ["A4", "=ROUNDUP(NOSUCHFN(), 0)"],
      ["A5", "=ROUNDUP(3.2)"],
      ["A6", "=ROUNDUP(INFINITY, 0)"],
    ]);

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
    expect(cell("A3")).toBe(FormulaError.VALUE);
    expect(cell("A4")).toBe(FormulaError.NAME);
    expect(cell("A5")).toBe(FormulaError.VALUE);
    expect(cell("A6")).toBe("INFINITY");

    setCellContent("B1", 7.01);
    expect(cell("A3")).toBe(8);
  });

  test("reads an omitted num_digits as 0", () => {
    setCellContent("A1", "=ROUNDUP(7.01,)");
    expect(cell("A1")).toBe(8);
  });

  test("rounds arrays element by element", () => {
    setContent([
      ["A1", 1.01],
      ["A2", -2.01],
      ["A3", 3],
      ["C1", "=ROUNDUP(A1:A3, 0)"],
      ["D1", "=SUM(ROUNDUP(A1:A3, 0))"],
    ]);

    expect(cell("C1")).toBe(2);
    expect(cell("C2")).toBe(-3);
    expect(cell("C3")).toBe(3);
    expect(cell("D1")).toBe(2);
  });
});
