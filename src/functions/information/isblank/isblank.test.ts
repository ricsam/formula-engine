import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ISBLANK function", () => {
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
    test("is TRUE for an empty cell and FALSE for a populated one", () => {
      setContent([
        ["B2", 0],
        ["B3", "text"],
        ["B4", false],
        ["B5", " "],
        ["A1", "=ISBLANK(B1)"],
        ["A2", "=ISBLANK(B2)"],
        ["A3", "=ISBLANK(B3)"],
        ["A4", "=ISBLANK(B4)"],
        ["A5", "=ISBLANK(B5)"],
      ]);

      expect(cell("A1")).toBe(true);
      expect(cell("A2")).toBe(false);
      expect(cell("A3")).toBe(false);
      expect(cell("A4")).toBe(false);
      expect(cell("A5")).toBe(false);
    });

    test("recalculates when the cell is filled or cleared", () => {
      setContent([["A1", "=ISBLANK(B1)"]]);
      expect(cell("A1")).toBe(true);

      setCellContent("B1", 5);
      expect(cell("A1")).toBe(false);

      setCellContent("B1", undefined);
      expect(cell("A1")).toBe(true);
    });

    test("treats every value that reads as empty text as blank", () => {
      // The engine reads an empty cell as "", so it cannot tell them apart.
      // Excel returns FALSE for A2 and A3.
      setContent([
        ["B2", '=""'],
        ["A1", "=ISBLANK(B1)"],
        ["A2", "=ISBLANK(B2)"],
        ["A3", '=ISBLANK("")'],
        ["A4", '=ISBLANK(IF(TRUE, B1))'],
        ["A5", "=ISBLANK(INDEX(B1:B3, 3))"],
      ]);

      expect(cell("A1")).toBe(true);
      expect(cell("A2")).toBe(true);
      expect(cell("A3")).toBe(true);
      expect(cell("A4")).toBe(true);
      expect(cell("A5")).toBe(true);
    });

    test("returns FALSE for errors instead of propagating them", () => {
      setContent([
        ["B1", '=1+"x"'],
        ["A1", "=ISBLANK(B1)"],
        ["A2", "=ISBLANK(NOSUCHFN())"],
      ]);

      expect(cell("A1")).toBe(false);
      expect(cell("A2")).toBe(false);
    });

    test("guards a calculation on an empty input", () => {
      setContent([
        ["B1", 4],
        ["A1", '=IF(ISBLANK(B1), "", B1*2)'],
        ["A2", '=IF(ISBLANK(B2), "", B2*2)'],
      ]);

      expect(cell("A1")).toBe(8);
      expect(cell("A2")).toBe("");
    });

    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=ISBLANK()"],
        ["A2", "=ISBLANK(B1, B2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });
  });

  describe("arrays", () => {
    test("tests a range element by element and spills", () => {
      setContent([
        ["A1", 1],
        // A2 is empty
        ["A3", "x"],
        ["C1", "=ISBLANK(A1:A4)"],
      ]);

      expect(cell("C1")).toBe(false);
      expect(cell("C2")).toBe(true);
      expect(cell("C3")).toBe(false);
      expect(cell("C4")).toBe(true);

      setCellContent("A2", "filled");
      expect(cell("C2")).toBe(false);
    });

    test("counts the empty cells of a range in an aggregate", () => {
      setContent([
        ["A1", 1],
        ["A3", "x"],
        ["A5", "=1/\"x\""],
        ["B2", 2],
        ["D1", "=SUM(IF(ISBLANK(A1:A6), 1, 0))"],
        ["D2", "=SUM(IF(ISBLANK(A1:B6), 1, 0))"],
      ]);

      // A2, A4 and A6
      expect(cell("D1")).toBe(3);
      // and B1 and B3:B6
      expect(cell("D2")).toBe(8);

      setCellContent("A4", 4);
      expect(cell("D1")).toBe(2);
      expect(cell("D2")).toBe(7);
    });

    test("counts the blank cells of a computed array", () => {
      setContent([
        ["A1", 1],
        ["A3", 3],
        // A1:A3&"" turns the empty A2 into "", the others into text
        ["B1", '=SUM(IF(ISBLANK(A1:A3&""), 1, 0))'],
        ["B2", '=SUM(IF(ISBLANK({"",1,""}), 1, 0))'],
      ]);

      expect(cell("B1")).toBe(1);
      expect(cell("B2")).toBe(2);
    });

    test("drives an array IF cell by cell", () => {
      setContent([
        ["A1", "a"],
        ["A3", "c"],
        ["C1", '=IF(ISBLANK(A1:A3), "missing", A1:A3)'],
      ]);

      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("missing");
      expect(cell("C3")).toBe("c");
    });
  });
});
