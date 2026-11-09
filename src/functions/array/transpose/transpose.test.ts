import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("TRANSPOSE function", () => {
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

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("ranges", () => {
    test("swaps rows and columns", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["C1", 3],
        ["A2", 4],
        ["B2", 5],
        ["C2", 6],
        ["E1", "=TRANSPOSE(A1:C2)"],
      ]);

      expect(block("E1", 3, 2)).toEqual([
        [1, 4],
        [2, 5],
        [3, 6],
      ]);
    });

    test("turns a row into a column and back", () => {
      setContent([
        ["A1", "a"],
        ["B1", "b"],
        ["C1", "c"],
        ["A3", "=TRANSPOSE(A1:C1)"],
        ["C3", "=TRANSPOSE(A3:A5)"],
      ]);

      expect(block("A3", 3, 1)).toEqual([["a"], ["b"], ["c"]]);
      expect(block("C3", 1, 3)).toEqual([["a", "b", "c"]]);
    });

    test("follows edits of the source cells", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["D1", "=TRANSPOSE(A1:B1)"],
      ]);
      expect(cell("D2")).toBe(2);

      setCellContent("B1", "changed");
      expect(cell("D2")).toBe("changed");
    });

    test("reads blank cells as empty text and keeps errors in their cells", () => {
      setContent([
        ["A1", 1],
        // B1 is empty
        ["C1", "=NOSUCHFN()"],
        ["E1", "=TRANSPOSE(A1:C1)"],
      ]);

      expect(cell("E1")).toBe(1);
      expect(cell("E2")).toBe("");
      expect(cell("E3")).toBe(FormulaError.NAME);
    });

    test("transposes an open-ended column into a row", () => {
      setContent([
        ["A1", 10],
        ["A2", 20],
        ["A4", 40],
        ["C1", "=INDEX(TRANSPOSE(A:A), 1, 4)"],
        ["C2", "=SUM(TRANSPOSE(A:A))"],
        ["C3", "=COLUMNS(TRANSPOSE(A1:A4))"],
      ]);

      expect(cell("C1")).toBe(40);
      expect(cell("C2")).toBe(70);
      expect(cell("C3")).toBe(4);
    });
  });

  describe("other arguments", () => {
    test("transposes computed arrays and array literals", () => {
      setContent([
        ["A1", "=TRANSPOSE(SEQUENCE(2, 3))"],
        ["D1", '=TRANSPOSE({"x","y"})'],
      ]);

      expect(block("A1", 3, 2)).toEqual([
        [1, 4],
        [2, 5],
        [3, 6],
      ]);
      expect(block("D1", 2, 1)).toEqual([["x"], ["y"]]);
    });

    test("returns a single value as is", () => {
      setContent([
        ["B1", "single"],
        ["A1", "=TRANSPOSE(5)"],
        ["A2", "=TRANSPOSE(B1)"],
        ["A3", "=TRANSPOSE({7})"],
      ]);

      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe("single");
      expect(cell("A3")).toBe(7);
    });

    test("returns an error argument", () => {
      setContent([
        ["A1", "=TRANSPOSE(NOSUCHFN())"],
        ["A2", '=TRANSPOSE(1+"x")'],
      ]);

      expect(cell("A1")).toBe(FormulaError.NAME);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=TRANSPOSE()"],
        ["A2", "=TRANSPOSE(B1:B2, 1)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });
  });

  describe("consumers", () => {
    test("is read whole by aggregates", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["E1", "=SUM(TRANSPOSE(A1:B2))"],
        ["E2", "=COUNTA(TRANSPOSE(A1:B2))"],
        ["E3", "=SUM(TRANSPOSE(SEQUENCE(3, 4)))"],
      ]);

      expect(cell("E1")).toBe(6);
      // the populated cells of A1:B2
      expect(cell("E2")).toBe(3);
      expect(cell("E3")).toBe(78);
    });

    test("is read in its own order by TOCOL", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=TOCOL(TRANSPOSE(A1:B2))"],
      ]);

      expect(block("D1", 4, 1)).toEqual([[1], [3], [2], [4]]);
    });

    test("broadcasts against other arrays", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        // a column times its transpose: a multiplication table
        ["C1", "=A1:A3*TRANSPOSE(A1:A3)"],
      ]);

      expect(block("C1", 3, 3)).toEqual([
        [1, 2, 3],
        [2, 4, 6],
        [3, 6, 9],
      ]);
    });

    test("undoes itself when applied twice", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=TRANSPOSE(TRANSPOSE(A1:B2))"],
      ]);

      expect(block("D1", 2, 2)).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    test("returns #SPILL! when the spill area is blocked", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["D1", "=TRANSPOSE(A1:B1)"],
        ["D2", "block"],
      ]);

      expect(cell("D1")).toBe(FormulaError.SPILL);

      setCellContent("D2", undefined);
      expect(cell("D1")).toBe(1);
      expect(cell("D2")).toBe(2);
    });
  });
});
