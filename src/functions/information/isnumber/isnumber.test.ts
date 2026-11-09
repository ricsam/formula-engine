import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ISNUMBER function", () => {
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

  describe("single values", () => {
    test("is TRUE for numbers only", () => {
      setContent([
        ["A1", "=ISNUMBER(42)"],
        ["A2", "=ISNUMBER(-0.5)"],
        ["A3", "=ISNUMBER(0)"],
        ["A4", '=ISNUMBER("42")'],
        ["A5", '=ISNUMBER("")'],
        ["A6", "=ISNUMBER(TRUE)"],
        ["A7", "=ISNUMBER(1+1)"],
      ]);

      expect(cell("A1")).toBe(true);
      expect(cell("A2")).toBe(true);
      expect(cell("A3")).toBe(true);
      expect(cell("A4")).toBe(false);
      expect(cell("A5")).toBe(false);
      expect(cell("A6")).toBe(false);
      expect(cell("A7")).toBe(true);
    });

    test("treats infinities as numbers", () => {
      setContent([
        ["A1", "=ISNUMBER(INFINITY)"],
        ["A2", "=ISNUMBER(-INFINITY)"],
        ["A3", "=ISNUMBER(1/0)"],
      ]);

      expect(cell("A1")).toBe(true);
      expect(cell("A2")).toBe(true);
      expect(cell("A3")).toBe(true);
    });

    test("tests referenced cells and recalculates", () => {
      setContent([
        ["B1", 10],
        ["B2", "10"],
        ["B3", false],
        // B4 is empty
        ["B5", "=B1*2"],
        ["A1", "=ISNUMBER(B1)"],
        ["A2", "=ISNUMBER(B2)"],
        ["A3", "=ISNUMBER(B3)"],
        ["A4", "=ISNUMBER(B4)"],
        ["A5", "=ISNUMBER(B5)"],
      ]);

      expect(cell("A1")).toBe(true);
      expect(cell("A2")).toBe(false);
      expect(cell("A3")).toBe(false);
      expect(cell("A4")).toBe(false);
      expect(cell("A5")).toBe(true);

      setCellContent("B4", 3);
      setCellContent("B1", "text");
      expect(cell("A4")).toBe(true);
      expect(cell("A1")).toBe(false);
      // B5 is now #VALUE!
      expect(cell("A5")).toBe(false);
    });

    test("returns FALSE for errors instead of propagating them", () => {
      setContent([
        ["B1", '=1+"x"'],
        ["A1", "=ISNUMBER(B1)"],
        ["A2", "=ISNUMBER(NOSUCHFN())"],
        ["A3", "=ISNUMBER(UnknownName)"],
        ["A4", '=ISNUMBER(MATCH("z",{"x","y"},0))'],
        ["A5", '=ISNUMBER(MATCH("y",{"x","y"},0))'],
      ]);

      expect(cell("A1")).toBe(false);
      expect(cell("A2")).toBe(false);
      expect(cell("A3")).toBe(false);
      expect(cell("A4")).toBe(false);
      expect(cell("A5")).toBe(true);
    });

    test("is used to guard calculations", () => {
      setContent([
        ["B1", 4],
        ["B2", "n/a"],
        ["A1", '=IF(ISNUMBER(B1), B1*2, "")'],
        ["A2", '=IF(ISNUMBER(B2), B2*2, "")'],
      ]);

      expect(cell("A1")).toBe(8);
      expect(cell("A2")).toBe("");
    });

    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=ISNUMBER()"],
        ["A2", "=ISNUMBER(1, 2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });
  });

  describe("arrays", () => {
    test("tests a range element by element and spills", () => {
      setContent([
        ["A1", 1],
        ["A2", "two"],
        // A3 is empty
        ["A4", "=1/0"],
        ["A5", "=NOSUCHFN()"],
        ["C1", "=ISNUMBER(A1:A5)"],
      ]);

      expect(cell("C1")).toBe(true);
      expect(cell("C2")).toBe(false);
      expect(cell("C3")).toBe(false);
      expect(cell("C4")).toBe(true);
      expect(cell("C5")).toBe(false);

      setCellContent("A3", 3);
      expect(cell("C3")).toBe(true);
    });

    test("tests computed arrays and array literals", () => {
      setContent([
        ["A1", '=ISNUMBER({1,"a";TRUE,2})'],
        ["A3", "=ISNUMBER(SEQUENCE(2)/{1;0})"],
      ]);

      expect(cell("A1")).toBe(true);
      expect(cell("B1")).toBe(false);
      expect(cell("A2")).toBe(false);
      expect(cell("B2")).toBe(true);
      expect(cell("A3")).toBe(true);
      expect(cell("A4")).toBe(true);
    });

    test("counts the numbers of a range in an aggregate", () => {
      setContent([
        ["A1", 1],
        ["A2", "two"],
        ["A4", 4],
        ["A5", "=1+\"x\""],
        ["B1", "=SUM(IF(ISNUMBER(A1:A6), 1, 0))"],
        ["B2", "=COUNTA(ISNUMBER(A1:A6))"],
      ]);

      expect(cell("B1")).toBe(2);
      // every cell of the bounded result, the empty ones included
      expect(cell("B2")).toBe(6);
    });

    test("reads an open-ended range by its populated cells", () => {
      setContent([
        ["A1", 1],
        ["A3", "x"],
        ["B1", "=SUM(IF(ISNUMBER(A:A), 1, 0))"],
      ]);

      expect(cell("B1")).toBe(1);
    });
  });
});
