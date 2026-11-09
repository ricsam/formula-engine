import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ROUND function", () => {
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

  describe("single numbers", () => {
    test("rounds to decimal places", () => {
      setContent([
        ["A1", "=ROUND(2.15, 1)"],
        ["A2", "=ROUND(2.149, 1)"],
        ["A3", "=ROUND(-1.475, 2)"],
        ["A4", "=ROUND(21.5, 0)"],
        ["A5", "=ROUND(3.14159, 3)"],
        ["A6", "=ROUND(7, 2)"],
      ]);

      expect(cell("A1")).toBe(2.2);
      expect(cell("A2")).toBe(2.1);
      expect(cell("A3")).toBe(-1.48);
      expect(cell("A4")).toBe(22);
      expect(cell("A5")).toBe(3.142);
      expect(cell("A6")).toBe(7);
    });

    test("rounds halves away from zero", () => {
      setContent([
        ["A1", "=ROUND(0.5, 0)"],
        ["A2", "=ROUND(1.5, 0)"],
        ["A3", "=ROUND(2.5, 0)"],
        ["A4", "=ROUND(-0.5, 0)"],
        ["A5", "=ROUND(-2.5, 0)"],
        ["A6", "=ROUND(0.4999, 0)"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(2);
      expect(cell("A3")).toBe(3);
      expect(cell("A4")).toBe(-1);
      expect(cell("A5")).toBe(-3);
      expect(cell("A6")).toBe(0);
    });

    test("rounds the decimal value, as Excel does, not the binary one", () => {
      // 2.675, 1.005 and 8.325 are stored slightly below their decimal value,
      // so Math.round(x * 100) / 100 rounds them down.
      setContent([
        ["A1", "=ROUND(2.675, 2)"],
        ["A2", "=ROUND(1.005, 2)"],
        ["A3", "=ROUND(8.325, 2)"],
        ["A4", "=ROUND(0.285, 2)"],
        ["A5", "=ROUND(5.015, 2)"],
        ["A6", "=ROUND(-1.005, 2)"],
      ]);

      expect(cell("A1")).toBe(2.68);
      expect(cell("A2")).toBe(1.01);
      expect(cell("A3")).toBe(8.33);
      expect(cell("A4")).toBe(0.29);
      expect(cell("A5")).toBe(5.02);
      expect(cell("A6")).toBe(-1.01);
    });

    test("rounds to the left of the decimal point with negative num_digits", () => {
      setContent([
        ["A1", "=ROUND(1234.5678, -2)"],
        ["A2", "=ROUND(1250, -2)"],
        ["A3", "=ROUND(-1250, -2)"],
        ["A4", "=ROUND(49, -2)"],
        ["A5", "=ROUND(50, -2)"],
        ["A6", "=ROUND(987654321, -6)"],
      ]);

      expect(cell("A1")).toBe(1200);
      expect(cell("A2")).toBe(1300);
      expect(cell("A3")).toBe(-1300);
      expect(cell("A4")).toBe(0);
      expect(cell("A5")).toBe(100);
      expect(cell("A6")).toBe(988000000);
    });

    test("truncates num_digits to a whole number", () => {
      setContent([
        ["A1", "=ROUND(2.567, 1.9)"],
        ["A2", "=ROUND(2.567, -0.9)"],
        ["A3", "=ROUND(1234.5, -1.5)"],
      ]);

      expect(cell("A1")).toBe(2.6);
      expect(cell("A2")).toBe(3);
      expect(cell("A3")).toBe(1230);
    });

    test("keeps numbers that need no rounding", () => {
      setContent([
        ["A1", "=ROUND(1/3, 20)"],
        ["A2", "=ROUND(0.1, 5)"],
        ["A3", "=ROUND(123456789012345, 0)"],
        ["A4", "=ROUND(1E-20, 25)"],
        ["A5", "=ROUND(0, 3)"],
      ]);

      expect(cell("A1")).toBe(1 / 3);
      expect(cell("A2")).toBe(0.1);
      expect(cell("A3")).toBe(123456789012345);
      expect(cell("A4")).toBe(1e-20);
      expect(cell("A5")).toBe(0);
    });

    test("rounds numbers with more than 15 significant digits", () => {
      setContent([
        ["A1", "=ROUND(1234567890123.456, 2)"],
        ["A2", "=ROUND(1/3, 15)"],
        ["A3", "=ROUND(0.1+0.2, 15)"],
        ["A4", "=ROUNDUP(0.1+0.2, 16)"],
        ["A5", "=ROUNDDOWN(1234567890123.459, 2)"],
      ]);

      expect(cell("A1")).toBe(1234567890123.46);
      expect(cell("A2")).toBe(0.333333333333333);
      // The binary noise of 0.30000000000000004 is not a digit Excel shows
      expect(cell("A3")).toBe(0.3);
      expect(cell("A4")).toBe(0.3);
      // Excel shows 1234567890123.459 as 1234567890123.46
      expect(cell("A5")).toBe(1234567890123.46);
    });

    test("never returns negative zero", () => {
      setContent([
        ["A1", "=ROUND(-0.4, 0)"],
        ["A2", "=ROUND(-0.0001, 2)"],
        ["A3", "=ROUND(-4, -1)"],
      ]);

      // toBe distinguishes 0 from -0
      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(0);
    });

    test("rounds a referenced cell and recalculates", () => {
      setContent([
        ["A1", 3.14159],
        ["A2", 2],
        ["B1", "=ROUND(A1, A2)"],
      ]);
      expect(cell("B1")).toBe(3.14);

      setCellContent("A1", 2.71828);
      expect(cell("B1")).toBe(2.72);

      setCellContent("A2", 0);
      expect(cell("B1")).toBe(3);
    });
  });

  describe("limits and infinities", () => {
    test("returns infinities unchanged", () => {
      setContent([
        ["A1", "=ROUND(INFINITY, 2)"],
        ["A2", "=ROUND(-INFINITY, -2)"],
        ["A3", "=ROUND(1/0, 0)"],
      ]);

      expect(cell("A1")).toBe("INFINITY");
      expect(cell("A2")).toBe("-INFINITY");
      expect(cell("A3")).toBe("INFINITY");
    });

    test("handles very large and infinite num_digits", () => {
      setContent([
        ["A1", "=ROUND(2.5, 1000)"],
        ["A2", "=ROUND(2.5, -1000)"],
        ["A3", "=ROUND(2.5, INFINITY)"],
        ["A4", "=ROUND(2.5, -INFINITY)"],
      ]);

      expect(cell("A1")).toBe(2.5);
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(2.5);
      expect(cell("A4")).toBe(0);
    });

    test("returns #NUM! when the rounded value is too large for a number", () => {
      setContent([
        ["A1", "=ROUND(1.7E+308, -308)"],
        ["A2", "=ROUND(1.2E+308, -308)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.NUM);
      expect(cell("A2")).toBe(1e308);
    });
  });

  describe("argument types", () => {
    test("accepts only numbers, like INT and CEILING", () => {
      setContent([
        ["B1", "2.5"],
        ["B2", true],
        ["A1", '=ROUND("2.5", 0)'],
        ["A2", "=ROUND(TRUE, 0)"],
        ["A3", "=ROUND(B1, 0)"],
        ["A4", "=ROUND(B2, 0)"],
        // B3 is empty
        ["A5", "=ROUND(B3, 0)"],
        ["A6", '=ROUND(2.5, "1")'],
        ["A7", "=ROUND(2.5, FALSE)"],
        ["A8", "=ROUND(2.5, B3)"],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8"]) {
        expect(cell(ref)).toBe(FormulaError.VALUE);
      }

      setCellContent("B3", 2.25);
      expect(cell("A5")).toBe(2);
      expect(cell("A8")).toBe(2.5);
    });

    test("propagates errors, the number's first", () => {
      setContent([
        ["B1", "=CEILING(5, 0)"],
        ["A1", "=ROUND(B1, 0)"],
        ["A2", "=ROUND(2.5, B1)"],
        ["A3", "=ROUND(NOSUCHFN(), B1)"],
        ["A4", '=ROUND(1+"x", NOSUCHFN())'],
      ]);

      expect(cell("A1")).toBe(FormulaError.DIV0);
      expect(cell("A2")).toBe(FormulaError.DIV0);
      expect(cell("A3")).toBe(FormulaError.NAME);
      expect(cell("A4")).toBe(FormulaError.VALUE);
    });

    test("requires exactly two arguments", () => {
      setContent([
        ["A1", "=ROUND(2.5)"],
        ["A2", "=ROUND(2.5, 1, 1)"],
        ["A3", "=ROUND()"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("reads an omitted argument as 0, as in Excel", () => {
      setContent([
        ["A1", "=ROUND(2.5,)"],
        ["A2", "=ROUND(,2)"],
        ["A3", "=ROUND(-2.567,)"],
      ]);

      expect(cell("A1")).toBe(3);
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(-3);
    });
  });

  describe("arrays", () => {
    test("rounds a range element by element and spills", () => {
      setContent([
        ["A1", 1.25],
        ["A2", 2.5],
        ["A3", -3.75],
        ["C1", "=ROUND(A1:A3, 1)"],
      ]);

      expect(cell("C1")).toBe(1.3);
      expect(cell("C2")).toBe(2.5);
      expect(cell("C3")).toBe(-3.8);

      setCellContent("A2", 9.99);
      expect(cell("C2")).toBe(10);
    });

    test("broadcasts the number against an array of num_digits", () => {
      setContent([
        ["A1", 3.14159],
        ["C1", "=ROUND(A1, {0,1,2,3})"],
      ]);

      expect(cell("C1")).toBe(3);
      expect(cell("D1")).toBe(3.1);
      expect(cell("E1")).toBe(3.14);
      expect(cell("F1")).toBe(3.142);
    });

    test("broadcasts a column of numbers against a row of num_digits", () => {
      setCellContent("C1", "=ROUND({1.25;2.75}, {0,1})");

      expect(cell("C1")).toBe(1);
      expect(cell("D1")).toBe(1.3);
      expect(cell("C2")).toBe(3);
      expect(cell("D2")).toBe(2.8);
    });

    test("keeps an error in one element in that element", () => {
      setContent([
        ["A1", 1.26],
        ["A2", "text"],
        ["A3", "=NOSUCHFN()"],
        ["A5", 4.44],
        ["C1", "=ROUND(A1:A5, 1)"],
      ]);

      expect(cell("C1")).toBe(1.3);
      expect(cell("C2")).toBe(FormulaError.VALUE);
      expect(cell("C3")).toBe(FormulaError.NAME);
      // A4 is empty, which is not a number
      expect(cell("C4")).toBe(FormulaError.VALUE);
      expect(cell("C5")).toBe(4.4);
    });

    test("is read whole by aggregates", () => {
      setContent([
        ["A1", 1.4],
        ["A2", 2.6],
        ["A3", 3.5],
        ["B1", "=SUM(ROUND(A1:A3, 0))"],
        ["B2", "=SUM(ROUND(SEQUENCE(4)/3, 1))"],
        ["B3", "=MAX(ROUND({1.44,2.45}, 1))"],
      ]);

      expect(cell("B1")).toBe(1 + 3 + 4);
      // 0.3 + 0.7 + 1 + 1.3
      expect(cell("B2")).toBeCloseTo(3.3, 10);
      expect(cell("B3")).toBe(2.5);
    });

    test("skips the empty cells of a range in aggregates", () => {
      setContent([
        ["A1", 1.4],
        ["A3", 3.5],
        ["B1", "=SUM(ROUND(A1:A3, 0))"],
        ["B2", "=COUNT(ROUND(A1:A3, 0))"],
      ]);

      expect(cell("B1")).toBe(5);
      expect(cell("B2")).toBe(2);
    });

    test("returns #SPILL! when the spill area is blocked", () => {
      setContent([
        ["A1", 1.25],
        ["A2", 2.5],
        ["C1", "=ROUND(A1:A2, 0)"],
        ["C2", "block"],
      ]);

      expect(cell("C1")).toBe(FormulaError.SPILL);

      setCellContent("C2", undefined);
      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(3);
    });
  });
});
