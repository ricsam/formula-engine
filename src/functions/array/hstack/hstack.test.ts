import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { indexToColumn, parseCellReference } from "../../../core/utils";

describe("HSTACK function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({
      sheetName,
      workbookName,
      ...parseCellReference(ref),
    });

  /** The values of the rectangle `from`:`to`, row by row. */
  const grid = (from: string, to: string) => {
    const start = parseCellReference(from);
    const end = parseCellReference(to);
    const rows: SerializedCellValue[][] = [];
    for (let row = start.rowIndex; row <= end.rowIndex; row++) {
      const values: SerializedCellValue[] = [];
      for (let col = start.colIndex; col <= end.colIndex; col++) {
        values.push(cell(`${indexToColumn(col)}${row + 1}`));
      }
      rows.push(values);
    }
    return rows;
  };

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

  const NA = FormulaError.NA;

  describe("stacking", () => {
    test("places arrays of equal height side by side", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", "a"],
        ["C1", "b"],
        ["B2", "c"],
        ["C2", "d"],
        ["E1", "=HSTACK(A1:A2,B1:C2)"],
      ]);

      expect(grid("E1", "H3")).toEqual([
        [1, "a", "b", ""],
        [2, "c", "d", ""],
        ["", "", "", ""],
      ]);
    });

    test("pads shorter arguments with #N/A without repeating a row or value", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", 10],
        ["B2", 20],
        ["C1", "x"],
        ["D1", "y"],
        ["F1", "=HSTACK(A1:A3,B1:B2,99,C1:D1)"],
      ]);

      expect(grid("F1", "J4")).toEqual([
        [1, 10, 99, "x", "y"],
        [2, 20, NA, NA, NA],
        [3, NA, NA, NA, NA],
        ["", "", "", "", ""],
      ]);
    });

    test("stacks single values into one row", () => {
      setContent([
        ["A1", '=HSTACK(1,"a",TRUE)'],
        ["A3", "=HSTACK(SEQUENCE(1,2),3)"],
      ]);

      expect(grid("A1", "D1")).toEqual([[1, "a", true, ""]]);
      expect(grid("A3", "D3")).toEqual([[1, 2, 3, ""]]);
    });

    test("returns a single array unchanged", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["D1", "=HSTACK(A1:B2)"],
      ]);

      expect(grid("D1", "E2")).toEqual([
        [1, 2],
        [3, 4],
      ]);
    });

    test("collapses a one-cell result to a single value", () => {
      setContent([
        ["A1", "v"],
        ["C1", "=HSTACK(5)"],
        ["C2", "=HSTACK(A1)"],
      ]);

      expect(cell("C1")).toBe(5);
      expect(cell("D1")).toBe("");
      expect(cell("C2")).toBe("v");
    });

    test("keeps an argument that evaluates to an error as a 1x1 error cell", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=HSTACK(NOSUCHFN(),A1:A2)"],
      ]);

      expect(grid("C1", "D2")).toEqual([
        [FormulaError.NAME, 1],
        [NA, 2],
      ]);
    });

    test("keeps error cells of an argument", () => {
      setContent([
        ["A1", "=NOSUCHFN()"],
        ["A2", 2],
        ["C1", "=HSTACK(A1:A2,A1:A2)"],
      ]);

      expect(grid("C1", "D2")).toEqual([
        [FormulaError.NAME, FormulaError.NAME],
        [2, 2],
      ]);
    });
  });

  describe("spilled inputs", () => {
    test("stacks TOCOL results of different lengths", () => {
      setContent([
        ["A1", "a"],
        ["A3", "b"],
        ["B1", 1],
        ["B2", 2],
        ["B4", 3],
        ["D1", "=HSTACK(TOCOL(A1:A4,1),TOCOL(B1:B4,1))"],
      ]);

      expect(grid("D1", "E4")).toEqual([
        ["a", 1],
        ["b", 2],
        [NA, 3],
        ["", ""],
      ]);
    });

    test("stacks SEQUENCE and UNIQUE results", () => {
      setContent([
        ["A1", "x"],
        ["A2", "y"],
        ["A3", "x"],
        ["C1", "=HSTACK(SEQUENCE(3),UNIQUE(A1:A3),SEQUENCE(2,2,10))"],
      ]);

      expect(grid("C1", "F3")).toEqual([
        [1, "x", 10, 11],
        [2, "y", 12, 13],
        [3, NA, NA, NA],
      ]);
    });

    test("stacks references to spilling cells", () => {
      setContent([
        ["A1", "=SEQUENCE(2)"],
        ["B1", "=SEQUENCE(2,1,5)"],
        ["D1", "=HSTACK(A1,B1)"],
      ]);

      expect(grid("D1", "E2")).toEqual([
        [1, 5],
        [2, 6],
      ]);
    });
  });

  describe("consumers", () => {
    test("feeds aggregates without applying an outer operation twice", () => {
      setContent([
        ["A1", "=SUM(HSTACK(SEQUENCE(2),SEQUENCE(2)*10))"],
        ["A2", "=SUM(HSTACK(SEQUENCE(2),SEQUENCE(2))*10)"],
        ["A3", "=COUNT(HSTACK(SEQUENCE(2,2),SEQUENCE(2)))"],
      ]);

      expect(cell("A1")).toBe(33);
      expect(cell("A2")).toBe(60);
      expect(cell("A3")).toBe(6);
    });

    test("passes #N/A padding on to consumers", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", 10],
        ["D1", "=COUNT(HSTACK(A1:A2,B1))"],
        ["D2", "=COUNT(HSTACK(A1:A2,A1:A2))"],
      ]);

      expect(cell("D1")).toBe(NA);
      expect(cell("D2")).toBe(4);
    });

    test("reads raw ranges sparsely", () => {
      setContent([
        ["A1", 1],
        ["A5", 5],
        ["B3", 30],
        ["D1", "=SUM(HSTACK(A1:A5,B1:B5))"],
        ["D2", "=SUM(HSTACK(A:A,B:B))"],
      ]);

      expect(cell("D1")).toBe(36);
      expect(cell("D2")).toBe(36);
    });

    test("feeds UNIQUE, which compares whole rows", () => {
      setContent([
        ["A1", "a"],
        ["A2", "a"],
        ["A3", "b"],
        ["B1", 1],
        ["B2", 1],
        ["B3", 1],
        ["D1", "=UNIQUE(HSTACK(A1:A3,B1:B3))"],
      ]);

      expect(grid("D1", "E3")).toEqual([
        ["a", 1],
        ["b", 1],
        ["", ""],
      ]);
    });

    test("supports INDEX", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", 3],
        ["D1", "=INDEX(HSTACK(A1:A2,B1),2,1)"],
        ["D2", "=INDEX(HSTACK(A1:A2,B1),1,2)"],
        ["D3", "=INDEX(HSTACK(A1:A2,B1),2,2)"],
      ]);

      expect(cell("D1")).toBe(2);
      expect(cell("D2")).toBe(3);
      expect(cell("D3")).toBe(NA);
    });

    test("is read through ranges that cover its spill", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", 10],
        ["B2", 20],
        ["D1", "=HSTACK(A1:A2,B1:B2)"],
        ["G1", "=SUM(E:E)"],
        ["G2", "=SUM(D2:E2)"],
      ]);

      expect(cell("G1")).toBe(30);
      expect(cell("G2")).toBe(22);
    });
  });

  describe("unbounded arguments", () => {
    test("stays lazy with an unbounded number of rows", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", "x"],
        ["D1", "=HSTACK(A:A,B1)"],
      ]);

      expect(grid("D1", "E3")).toEqual([
        [1, "x"],
        [2, NA],
        ["", NA],
      ]);
      expect(cell("D500")).toBe("");
      expect(cell("E500")).toBe(NA);

      setCellContent("A500", "far");
      expect(cell("D500")).toBe("far");
    });

    test("allows an unbounded number of columns in the last argument only", () => {
      setContent([
        ["A2", "a"],
        ["B2", "b"],
        ["Z9", "first"],
        ["A5", "=HSTACK(Z9,2:2)"],
        ["A6", "=HSTACK(2:2,Z9)"],
      ]);

      expect(grid("A5", "D5")).toEqual([["first", "a", "b", ""]]);
      expect(cell("A6")).toBe(FormulaError.VALUE);
    });

    test("cannot list infinitely many padding cells", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["D1", "=COUNT(HSTACK(A:A,B1))"],
      ]);

      expect(cell("D1")).toBe(FormulaError.REF);
    });
  });

  describe("recalculation", () => {
    test("reflects an edited source cell", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["D1", "=HSTACK(A1,B1)"],
      ]);

      expect(grid("D1", "E1")).toEqual([[1, 2]]);

      setCellContent("B1", 20);
      expect(grid("D1", "E1")).toEqual([[1, 20]]);
    });

    test("follows an argument that changes size", () => {
      setContent([
        ["A1", "a"],
        ["B1", 1],
        ["B2", 2],
        ["D1", "=HSTACK(TOCOL(A1:A3,1),B1:B2)"],
      ]);

      expect(grid("D1", "E2")).toEqual([
        ["a", 1],
        [NA, 2],
      ]);

      setCellContent("A3", "c");
      expect(grid("D1", "E2")).toEqual([
        ["a", 1],
        ["c", 2],
      ]);

      setCellContent("A2", "b");
      expect(grid("D1", "E4")).toEqual([
        ["a", 1],
        ["b", 2],
        ["c", NA],
        ["", ""],
      ]);
    });
  });

  describe("errors", () => {
    test("returns #SPILL! when the result area is blocked", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=HSTACK(A1:A2,A1:A2)"],
        ["D2", "blocker"],
      ]);

      expect(cell("C1")).toBe(FormulaError.SPILL);
    });

    test("requires at least one argument", () => {
      setContent([["A1", "=HSTACK()"]]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("returns #VALUE! for an omitted argument, as Excel does", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", 3],
        ["D1", "=HSTACK(A1:A2,,B1)"],
        ["D2", "=HSTACK(A1:A2,)"],
        ["D3", "=HSTACK(,)"],
        ["D5", '=HSTACK(A1:A2,"",B1)'],
      ]);

      expect(cell("D1")).toBe(FormulaError.VALUE);
      expect(cell("D2")).toBe(FormulaError.VALUE);
      expect(cell("D3")).toBe(FormulaError.VALUE);
      // An empty string stands in for a blank cell.
      expect(grid("D5", "F6")).toEqual([
        [1, "", 3],
        [2, NA, NA],
      ]);
    });
  });

  describe("performance", () => {
    test("reads the last row of a tall stack without reading the rest", () => {
      const start = performance.now();
      setContent([
        ["A1", "=C90000"],
        ["B1", "=HSTACK(TOCOL(SEQUENCE(300,300)),TOCOL(SEQUENCE(300,300)*2))"],
      ]);

      expect(cell("A1")).toBe(180000);
      expect(performance.now() - start).toBeLessThan(200);
    });
  });
});
