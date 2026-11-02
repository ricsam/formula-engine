import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { indexToColumn, parseCellReference } from "../../../core/utils";

describe("TOROW function", () => {
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

  /** Values of row `rowNumber` from column `from` to column `to` (letters). */
  const row = (rowNumber: number, from: string, to: string) => {
    const values: SerializedCellValue[] = [];
    const first = parseCellReference(`${from}1`).colIndex;
    const last = parseCellReference(`${to}1`).colIndex;
    for (let col = first; col <= last; col++) {
      values.push(cell(`${indexToColumn(col)}${rowNumber}`));
    }
    return values;
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

  /**
   * A1:B3 holds    1 | #NAME?
   *            blank | 4
   *              "c" | #VALUE!
   */
  const gridWithBlanksAndErrors: Array<[string, SerializedCellValue]> = [
    ["A1", 1],
    ["B1", "=NOSUCHFN()"],
    ["B2", 4],
    ["A3", "c"],
    ["B3", '="a"+1'],
  ];

  describe("layout", () => {
    test("lays a 2D range out in one row, row by row by default", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["A5", "=TOROW(A1:B2)"],
        ["A6", "=TOROW(A1:B2,0,TRUE)"],
      ]);

      expect(row(5, "A", "E")).toEqual([1, 2, 3, 4, ""]);
      expect(row(6, "A", "E")).toEqual([1, 3, 2, 4, ""]);
      expect(cell("A7")).toBe("");
    });

    test("turns a column into a row", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "c"],
        ["C1", "=TOROW(A1:A3)"],
      ]);

      expect(row(1, "C", "F")).toEqual(["a", "b", "c", ""]);
      expect(cell("C2")).toBe("");
    });
  });

  describe("ignore", () => {
    test("supports every mode", () => {
      setContent([
        ...gridWithBlanksAndErrors,
        ["A5", "=TOROW(A1:B3)"],
        ["A6", "=TOROW(A1:B3,1)"],
        ["A7", "=TOROW(A1:B3,2)"],
        ["A8", "=TOROW(A1:B3,3)"],
        ["A9", "=TOROW(A1:B3,3,TRUE)"],
      ]);

      expect(row(5, "A", "G")).toEqual([
        1,
        FormulaError.NAME,
        "",
        4,
        "c",
        FormulaError.VALUE,
        "",
      ]);
      expect(row(6, "A", "F")).toEqual([
        1,
        FormulaError.NAME,
        4,
        "c",
        FormulaError.VALUE,
        "",
      ]);
      expect(row(7, "A", "E")).toEqual([1, "", 4, "c", ""]);
      expect(row(8, "A", "D")).toEqual([1, 4, "c", ""]);
      expect(row(9, "A", "D")).toEqual([1, "c", 4, ""]);
    });

    test("rejects an invalid ignore value and returns #CALC! when nothing is left", () => {
      setContent([
        ["A1", 1],
        ["C1", "=TOROW(A1:B1,7)"],
        ["C2", "=TOROW(E5:G5,1)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.VALUE);
      expect(cell("C2")).toBe(FormulaError.CALC);
    });

    test("keeps the defaults of omitted options", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["A2", 3],
        ["B2", 4],
        ["A5", "=TOROW(A1:B2,,TRUE)"],
        ["A6", "=TOROW(A1:B3,1,)"],
        ["A7", "=TOROW(,1)"],
      ]);

      expect(row(5, "A", "E")).toEqual([1, 3, 2, 4, ""]);
      expect(row(6, "A", "E")).toEqual([1, 2, 3, 4, ""]);
      expect(cell("A7")).toBe(FormulaError.VALUE);
    });
  });

  describe("single values", () => {
    test("returns a single value as it is", () => {
      setContent([
        ["A1", "x"],
        ["C1", "=TOROW(A1)"],
        ["C2", "=TOROW(42)"],
      ]);

      expect(cell("C1")).toBe("x");
      expect(cell("D1")).toBe("");
      expect(cell("C2")).toBe(42);
    });
  });

  describe("open-ended ranges", () => {
    test("skips the unbounded blank tail of a whole column", () => {
      setContent([
        ["A1", "a"],
        ["A3", "b"],
        ["A6", "c"],
        ["C1", "=TOROW(A:A,1)"],
      ]);

      expect(row(1, "C", "F")).toEqual(["a", "b", "c", ""]);
    });

    test("maps a whole column lazily into an unbounded row when blanks are kept", () => {
      setContent([
        ["A1", 1],
        ["A3", 3],
        ["C1", "=TOROW(A:A)"],
      ]);

      expect(row(1, "C", "F")).toEqual([1, "", 3, ""]);

      setCellContent("A30", "far");
      // A30 is the 30th cell of the column, so it lands 29 columns right of C1.
      expect(cell(`${indexToColumn(2 + 29)}1`)).toBe("far");
    });
  });

  describe("spilled inputs and consumers", () => {
    test("flattens SEQUENCE and UNIQUE results", () => {
      setContent([
        ["A1", "x"],
        ["A2", "y"],
        ["A3", "x"],
        ["C1", "=TOROW(SEQUENCE(2,2))"],
        ["C2", "=TOROW(UNIQUE(A1:A3))"],
      ]);

      expect(row(1, "C", "G")).toEqual([1, 2, 3, 4, ""]);
      expect(row(2, "C", "E")).toEqual(["x", "y", ""]);
    });

    test("feeds UNIQUE and SUM", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A4", "a"],
        ["C1", "=UNIQUE(TOROW(A1:A5,1),TRUE)"],
        ["C2", "=SUM(TOROW(SEQUENCE(3,3)))"],
        ["C3", "=SUM(-TOROW(SEQUENCE(2,2)))"],
      ]);

      expect(row(1, "C", "E")).toEqual(["a", "b", ""]);
      expect(cell("C2")).toBe(45);
      expect(cell("C3")).toBe(-10);
    });

    test("is read through ranges that cover its spill", () => {
      setContent([
        ["A3", 1],
        ["A4", 2],
        ["A5", 4],
        ["C1", "=TOROW(A3:A5)"],
        ["C2", "=SUM(1:1)"],
        ["C3", "=SUM(D1:E1)"],
      ]);

      expect(cell("C2")).toBe(7);
      expect(cell("C3")).toBe(6);
    });
  });

  describe("recalculation", () => {
    test("reflects edits, including a filled blank", () => {
      setContent([
        ["A1", "a"],
        ["A3", "c"],
        ["C1", "=TOROW(A1:A3)"],
        ["C2", "=TOROW(A1:A3,1)"],
      ]);

      expect(row(1, "C", "E")).toEqual(["a", "", "c"]);
      expect(row(2, "C", "E")).toEqual(["a", "c", ""]);

      setCellContent("A2", "b");

      expect(row(1, "C", "E")).toEqual(["a", "b", "c"]);
      expect(row(2, "C", "F")).toEqual(["a", "b", "c", ""]);
    });
  });

  describe("spilling", () => {
    test("returns #SPILL! when the result area is blocked", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=TOROW(A1:A2)"],
        ["D1", "blocker"],
      ]);

      expect(cell("C1")).toBe(FormulaError.SPILL);
    });
  });

  describe("performance", () => {
    test("reads the last cell of a long row without reading the rest", () => {
      const start = performance.now();
      setContent([
        ["A1", `=${indexToColumn(9999)}2`],
        ["A2", "=TOROW(SEQUENCE(100,100))"],
      ]);

      expect(cell("A1")).toBe(10000);
      expect(performance.now() - start).toBeLessThan(200);
    });
  });
});
