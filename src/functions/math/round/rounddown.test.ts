import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ROUNDDOWN function", () => {
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

  test("rounds toward zero", () => {
    setContent([
      ["A1", "=ROUNDDOWN(3.2, 0)"],
      ["A2", "=ROUNDDOWN(76.9, 0)"],
      ["A3", "=ROUNDDOWN(3.14159, 3)"],
      ["A4", "=ROUNDDOWN(-3.14159, 1)"],
      ["A5", "=ROUNDDOWN(31415.92654, -2)"],
      ["A6", "=ROUNDDOWN(-3.9, 0)"],
      ["A7", "=ROUNDDOWN(0.999, 0)"],
    ]);

    expect(cell("A1")).toBe(3);
    expect(cell("A2")).toBe(76);
    expect(cell("A3")).toBe(3.141);
    expect(cell("A4")).toBe(-3.1);
    expect(cell("A5")).toBe(31400);
    expect(cell("A6")).toBe(-3);
    expect(cell("A7")).toBe(0);
  });

  test("differs from INT for negative numbers", () => {
    setContent([
      ["A1", "=ROUNDDOWN(-8.9, 0)"],
      ["A2", "=INT(-8.9)"],
    ]);

    expect(cell("A1")).toBe(-8);
    expect(cell("A2")).toBe(-9);
  });

  test("works on the decimal value, as Excel does", () => {
    // In binary, 4.35*100 is 434.99999999999994 and 0.7+0.1 is
    // 0.7999999999999999, which a binary truncation would round down.
    setContent([
      ["A1", "=ROUNDDOWN(4.35*100, 0)"],
      ["A2", "=ROUNDDOWN(0.7+0.1, 1)"],
      ["A3", "=ROUNDDOWN(2.675, 2)"],
    ]);

    expect(cell("A1")).toBe(435);
    expect(cell("A2")).toBe(0.8);
    expect(cell("A3")).toBe(2.67);
  });

  test("returns 0 for a rounding position above the number", () => {
    setContent([
      ["A1", "=ROUNDDOWN(5, -1)"],
      ["A2", "=ROUNDDOWN(-999, -3)"],
      ["A3", "=ROUNDDOWN(5, -400)"],
    ]);

    // toBe distinguishes 0 from -0
    expect(cell("A1")).toBe(0);
    expect(cell("A2")).toBe(0);
    expect(cell("A3")).toBe(0);
  });

  test("accepts only numbers and propagates errors", () => {
    setContent([
      ["A1", '=ROUNDDOWN("3.2", 0)'],
      ["A2", "=ROUNDDOWN(3.2, TRUE)"],
      ["A3", "=ROUNDDOWN(B1, 0)"],
      ["A4", "=ROUNDDOWN(3.2, NOSUCHFN())"],
      ["A5", "=ROUNDDOWN(3.2, 1, 2)"],
      ["A6", "=ROUNDDOWN(-INFINITY, 0)"],
    ]);

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
    expect(cell("A3")).toBe(FormulaError.VALUE);
    expect(cell("A4")).toBe(FormulaError.NAME);
    expect(cell("A5")).toBe(FormulaError.VALUE);
    expect(cell("A6")).toBe("-INFINITY");

    setCellContent("B1", 7.99);
    expect(cell("A3")).toBe(7);
  });

  test("rounds arrays element by element", () => {
    setContent([
      ["A1", 1.99],
      ["A2", -2.99],
      ["C1", "=ROUNDDOWN(A1:A2, {0,1})"],
    ]);

    expect(cell("C1")).toBe(1);
    expect(cell("D1")).toBe(1.9);
    expect(cell("C2")).toBe(-2);
    expect(cell("D2")).toBe(-2.9);
  });
});
