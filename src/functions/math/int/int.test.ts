import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("INT function", () => {
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
    test("rounds down toward negative infinity", () => {
      setContent([
        ["A1", "=INT(8.9)"],
        ["A2", "=INT(-8.9)"],
        ["A3", "=INT(0.5)"],
        ["A4", "=INT(-0.5)"],
        ["A5", "=INT(-1.0001)"],
        ["A6", "=INT(2.9999)"],
      ]);

      expect(cell("A1")).toBe(8);
      expect(cell("A2")).toBe(-9);
      expect(cell("A3")).toBe(0);
      expect(cell("A4")).toBe(-1);
      expect(cell("A5")).toBe(-2);
      expect(cell("A6")).toBe(2);
    });

    test("keeps integers unchanged", () => {
      setContent([
        ["A1", "=INT(5)"],
        ["A2", "=INT(-5)"],
        ["A3", "=INT(0)"],
        ["A4", "=INT(123456789012)"],
      ]);

      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe(-5);
      expect(cell("A3")).toBe(0);
      expect(cell("A4")).toBe(123456789012);
    });

    test("never returns negative zero", () => {
      setContent([
        ["A1", "=INT(-0)"],
        ["A2", "=INT(0.25)"],
        ["B1", -0],
        ["A3", "=INT(B1)"],
      ]);

      // toBe distinguishes 0 from -0
      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(0);
      expect(cell("A3")).toBe(0);
    });

    test("rounds a referenced cell and recalculates", () => {
      setContent([
        ["A1", 3.7],
        ["B1", "=INT(A1)"],
        ["B2", "=INT(A1 * 2)"],
      ]);
      expect(cell("B1")).toBe(3);
      expect(cell("B2")).toBe(7);

      setCellContent("A1", -3.2);
      expect(cell("B1")).toBe(-4);
      expect(cell("B2")).toBe(-7);
    });

    test("returns infinities unchanged", () => {
      setContent([
        ["A1", "=INT(INFINITY)"],
        ["A2", "=INT(-INFINITY)"],
        ["A3", "=INT(1/0)"],
      ]);

      expect(cell("A1")).toBe("INFINITY");
      expect(cell("A2")).toBe("-INFINITY");
      expect(cell("A3")).toBe("INFINITY");
    });
  });

  describe("argument types", () => {
    test("rejects text, logical values and empty cells like CEILING does", () => {
      setContent([
        ["B1", "4.5"],
        ["B2", true],
        ["A1", '=INT("abc")'],
        ["A2", '=INT("3.7")'],
        ["A3", "=INT(TRUE)"],
        ["A4", "=INT(B1)"],
        ["A5", "=INT(B2)"],
        // B3 is empty
        ["A6", "=INT(B3)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
      expect(cell("A4")).toBe(FormulaError.VALUE);
      expect(cell("A5")).toBe(FormulaError.VALUE);
      expect(cell("A6")).toBe(FormulaError.VALUE);

      setCellContent("B3", 9.99);
      expect(cell("A6")).toBe(9);
    });

    test("propagates errors", () => {
      setContent([
        ["B1", "=CEILING(5, 0)"],
        ["A1", "=INT(B1)"],
        ["A2", "=INT(CEILING())"],
      ]);

      expect(cell("A1")).toBe(FormulaError.DIV0);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=INT()"],
        ["A2", "=INT(1.5, 2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });
  });

  describe("arrays", () => {
    test("rounds a range element by element and spills", () => {
      setContent([
        ["A1", 1.5],
        ["A2", -1.5],
        ["A3", 7],
        ["C1", "=INT(A1:A3)"],
        ["E1", "=ROWS(INT(A1:A3))"],
        ["E2", "=COLUMNS(INT(SEQUENCE(2, 3)/2))"],
      ]);

      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(-2);
      expect(cell("C3")).toBe(7);
      expect(cell("C4")).toBe("");
      // The result keeps the argument's shape.
      expect(cell("E1")).toBe(3);
      expect(cell("E2")).toBe(3);

      setCellContent("A2", 4.75);
      expect(cell("C2")).toBe(4);
    });

    test("keeps errors local to their element", () => {
      setContent([
        ["A1", 2.5],
        ["A2", "text"],
        ["A3", "=CEILING(5, 0)"],
        ["A4", 3.5],
        ["C1", "=INT(A1:A4)"],
      ]);

      expect(cell("C1")).toBe(2);
      expect(cell("C2")).toBe(FormulaError.VALUE);
      expect(cell("C3")).toBe(FormulaError.DIV0);
      expect(cell("C4")).toBe(3);
    });

    test("rounds spilled results", () => {
      setContent([
        ["A1", "=INT(SEQUENCE(4)/2)"],
        ["C1", "=INT(SEQUENCE(2, 3, -1, 0.75))"],
        ["G1", "=INT({2.5,-2.5})"],
      ]);

      // SEQUENCE(4)/2 is 0.5, 1, 1.5, 2
      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(2);
      expect(cell("A5")).toBe("");

      // -1, -0.25, 0.5 / 1.25, 2, 2.75
      expect(cell("C1")).toBe(-1);
      expect(cell("D1")).toBe(-1);
      expect(cell("E1")).toBe(0);
      expect(cell("C2")).toBe(1);
      expect(cell("D2")).toBe(2);
      expect(cell("E2")).toBe(2);

      expect(cell("G1")).toBe(2);
      expect(cell("H1")).toBe(-3);
    });

    test("rounds the whole spill of a referenced spilling cell", () => {
      setContent([
        ["A1", 3],
        ["B1", "=SEQUENCE(A1, 1, 0.5)"],
        ["D1", "=INT(B1)"],
      ]);

      // 0.5, 1.5, 2.5
      expect(cell("D1")).toBe(0);
      expect(cell("D2")).toBe(1);
      expect(cell("D3")).toBe(2);
      expect(cell("D4")).toBe("");

      setCellContent("A1", 4);
      expect(cell("D4")).toBe(3);
    });

    test("provides its cells to aggregates and array functions", () => {
      setContent([
        ["A1", 1.5],
        ["A2", 2.5],
        ["A3", -0.5],
        ["C1", "=SUM(INT(A1:A3))"],
        ["C2", "=MAX(INT(A1:A3))"],
        ["C3", "=SUM(INT(SEQUENCE(4, 1, 0.5)))"],
        ["C4", "=ROWS(UNIQUE(INT(SEQUENCE(4, 1, 0.5, 0.5))))"],
        ["E1", "=UNIQUE(INT(SEQUENCE(4, 1, 0.5, 0.5)))"],
      ]);

      // 1 + 2 - 1
      expect(cell("C1")).toBe(2);
      expect(cell("C2")).toBe(2);
      // INT of 0.5, 1.5, 2.5, 3.5
      expect(cell("C3")).toBe(6);
      // INT of 0.5, 1, 1.5, 2 is 0, 1, 1, 2
      expect(cell("C4")).toBe(3);
      expect(cell("E1")).toBe(0);
      expect(cell("E2")).toBe(1);
      expect(cell("E3")).toBe(2);
      expect(cell("E4")).toBe("");

      setCellContent("A3", 10.9);
      expect(cell("C1")).toBe(13);
      expect(cell("C2")).toBe(10);
    });

    test("aggregates over an open-ended range", () => {
      setContent([
        ["B1", 1.5],
        ["B5", 2.7],
        ["B100", -0.5],
        ["A1", "=SUM(INT(B:B))"],
      ]);

      // 1 + 2 - 1
      expect(cell("A1")).toBe(2);

      setCellContent("B1000", 4.2);
      expect(cell("A1")).toBe(6);
    });
  });
});
