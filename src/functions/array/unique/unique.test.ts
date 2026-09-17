import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("UNIQUE function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

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
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("single column input", () => {
    test("keeps the first occurrence of each value in order", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["A4", "c"],
        ["A5", "b"],
        ["C1", "=UNIQUE(A1:A5)"],
      ]);

      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("b");
      expect(cell("C3")).toBe("c");
      // The result is shorter than the input, so C4 stays empty.
      expect(cell("C4")).toBe("");
    });

    test("collapses to a single value when every entry is identical", () => {
      setContent([
        ["A1", 7],
        ["A2", 7],
        ["A3", 7],
        ["C1", "=UNIQUE(A1:A3)"],
      ]);

      expect(cell("C1")).toBe(7);
      expect(cell("C2")).toBe("");
    });

    test("treats a single cell as already distinct", () => {
      setContent([
        ["A1", 5],
        ["C1", "=UNIQUE(A1)"],
      ]);

      expect(cell("C1")).toBe(5);
    });

    test("deduplicates numbers, booleans and infinity by value", () => {
      setContent([
        ["A1", 1],
        ["A2", true],
        ["A3", 1],
        ["A4", false],
        ["A5", "=1/0"],
        ["A6", "=2/0"],
        ["A7", true],
        ["C1", "=UNIQUE(A1:A7)"],
      ]);

      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(true);
      expect(cell("C3")).toBe(false);
      expect(cell("C4")).toBe("INFINITY");
      expect(cell("C5")).toBe("");
    });

    test("does not conflate values of different types", () => {
      setContent([
        ["A1", 1],
        ["A2", "1"],
        ["A3", true],
        ["C1", "=UNIQUE(A1:A3)"],
      ]);

      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe("1");
      expect(cell("C3")).toBe(true);
    });

    test("compares text case-insensitively and keeps the first casing", () => {
      setContent([
        ["A1", "Ab"],
        ["A2", "aB"],
        ["A3", "cd"],
        ["C1", "=UNIQUE(A1:A3)"],
      ]);

      expect(cell("C1")).toBe("Ab");
      expect(cell("C2")).toBe("cd");
      expect(cell("C3")).toBe("");
    });

    test("treats empty cells as a single empty value", () => {
      setContent([
        ["A1", "a"],
        ["A3", "b"],
        ["A4", "a"],
        ["C1", "=UNIQUE(A1:A4)"],
      ]);

      // A2 is empty and reads as "", which is a distinct value of its own.
      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("");
      expect(cell("C3")).toBe("b");
      expect(cell("C4")).toBe("");
    });

    test("spills away from the anchor cell", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["E5", "=UNIQUE(A1:A3)"],
      ]);

      expect(cell("E5")).toBe("a");
      expect(cell("E6")).toBe("b");
      expect(cell("E7")).toBe("");
    });
  });

  describe("single row input", () => {
    test("deduplicates down a column by default", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["C1", 1],
        ["D1", 3],
        ["A3", "=UNIQUE(A1:D1)"],
      ]);

      // by_col defaults to FALSE, so the single row is a single entry and the
      // result keeps the row shape.
      expect(cell("A3")).toBe(1);
      expect(cell("B3")).toBe(2);
      expect(cell("C3")).toBe(1);
      expect(cell("D3")).toBe(3);
      expect(cell("A4")).toBe("");
    });

    test("deduplicates across columns when by_col is TRUE", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["C1", 1],
        ["D1", 3],
        ["A3", "=UNIQUE(A1:D1,TRUE)"],
      ]);

      expect(cell("A3")).toBe(1);
      expect(cell("B3")).toBe(2);
      expect(cell("C3")).toBe(3);
      expect(cell("D3")).toBe("");
    });

    test("accepts 1 and 0 in place of TRUE and FALSE for by_col", () => {
      setContent([
        ["A1", 1],
        ["B1", 2],
        ["C1", 1],
        ["A3", "=UNIQUE(A1:C1,1)"],
        ["A5", "=UNIQUE(A1:C1,0)"],
      ]);

      expect(cell("A3")).toBe(1);
      expect(cell("B3")).toBe(2);
      expect(cell("C3")).toBe("");

      expect(cell("A5")).toBe(1);
      expect(cell("B5")).toBe(2);
      expect(cell("C5")).toBe(1);
    });
  });

  describe("two dimensional input", () => {
    test("compares whole rows by default", () => {
      setContent([
        ["A1", "x"],
        ["B1", 1],
        ["A2", "y"],
        ["B2", 2],
        ["A3", "x"],
        ["B3", 1],
        ["D1", "=UNIQUE(A1:B3)"],
      ]);

      expect(cell("D1")).toBe("x");
      expect(cell("E1")).toBe(1);
      expect(cell("D2")).toBe("y");
      expect(cell("E2")).toBe(2);
      expect(cell("D3")).toBe("");
      expect(cell("E3")).toBe("");
    });

    test("keeps rows that differ in only one column", () => {
      setContent([
        ["A1", "x"],
        ["B1", 1],
        ["A2", "x"],
        ["B2", 2],
        ["D1", "=UNIQUE(A1:B2)"],
      ]);

      expect(cell("D1")).toBe("x");
      expect(cell("E1")).toBe(1);
      expect(cell("D2")).toBe("x");
      expect(cell("E2")).toBe(2);
    });

    test("compares whole columns when by_col is TRUE", () => {
      setContent([
        ["A1", "x"],
        ["B1", "y"],
        ["C1", "x"],
        ["A2", 1],
        ["B2", 2],
        ["C2", 1],
        ["A4", "=UNIQUE(A1:C2,TRUE)"],
      ]);

      expect(cell("A4")).toBe("x");
      expect(cell("B4")).toBe("y");
      expect(cell("C4")).toBe("");
      expect(cell("A5")).toBe(1);
      expect(cell("B5")).toBe(2);
      expect(cell("C5")).toBe("");
    });
  });

  describe("exactly_once", () => {
    test("returns only values that occur once", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["A4", "c"],
        ["C1", "=UNIQUE(A1:A4,FALSE,TRUE)"],
      ]);

      expect(cell("C1")).toBe("b");
      expect(cell("C2")).toBe("c");
      expect(cell("C3")).toBe("");
    });

    test("counts a value appearing three times as a duplicate", () => {
      setContent([
        ["A1", "a"],
        ["A2", "a"],
        ["A3", "a"],
        ["A4", "b"],
        ["C1", "=UNIQUE(A1:A4,FALSE,TRUE)"],
      ]);

      expect(cell("C1")).toBe("b");
      expect(cell("C2")).toBe("");
    });

    test("applies to whole rows", () => {
      setContent([
        ["A1", "x"],
        ["B1", 1],
        ["A2", "y"],
        ["B2", 2],
        ["A3", "x"],
        ["B3", 1],
        ["A4", "z"],
        ["B4", 3],
        ["D1", "=UNIQUE(A1:B4,FALSE,TRUE)"],
      ]);

      expect(cell("D1")).toBe("y");
      expect(cell("E1")).toBe(2);
      expect(cell("D2")).toBe("z");
      expect(cell("E2")).toBe(3);
      expect(cell("D3")).toBe("");
    });

    test("returns #CALC! when every value is duplicated", () => {
      setContent([
        ["A1", "a"],
        ["A2", "a"],
        ["A3", "b"],
        ["A4", "b"],
        ["C1", "=UNIQUE(A1:A4,FALSE,TRUE)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.CALC);
    });
  });

  describe("unbounded ranges", () => {
    test("deduplicates a whole column", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["C1", "=UNIQUE(A:A)"],
      ]);

      // The unbounded tail of empty rows collapses into the single "" entry.
      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("b");
      expect(cell("C3")).toBe("");
    });

    test("deduplicates a whole row by column", () => {
      setContent([
        ["A1", "x"],
        ["B1", "y"],
        ["C1", "x"],
        ["A3", "=UNIQUE(1:1,TRUE)"],
      ]);

      expect(cell("A3")).toBe("x");
      expect(cell("B3")).toBe("y");
      expect(cell("C3")).toBe("");
    });

    test("never reports an unbounded empty tail as occurring exactly once", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["C1", "=UNIQUE(A:A,FALSE,TRUE)"],
      ]);

      expect(cell("C1")).toBe("b");
      expect(cell("C2")).toBe("");
    });

    test("rejects an unbounded number of columns when comparing rows", () => {
      setContent([
        ["A1", 1],
        ["A5", "=UNIQUE(1:2)"],
      ]);

      expect(cell("A5")).toBe(FormulaError.VALUE);
    });

    test("rejects an unbounded number of rows when comparing columns", () => {
      setContent([
        ["A1", 1],
        ["D1", "=UNIQUE(A:B,TRUE)"],
      ]);

      expect(cell("D1")).toBe(FormulaError.VALUE);
    });
  });

  describe("composition with other formulas", () => {
    test("feeds aggregate functions", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 1],
        ["A4", 3],
        ["C1", "=SUM(UNIQUE(A1:A4))"],
        ["C2", "=COUNT(UNIQUE(A1:A4))"],
        ["C3", "=MAX(UNIQUE(A1:A4))"],
      ]);

      expect(cell("C1")).toBe(6);
      expect(cell("C2")).toBe(3);
      expect(cell("C3")).toBe(3);
    });

    test("supports INDEX and MATCH lookups", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["C1", "=INDEX(UNIQUE(A1:A3),2)"],
        ["C2", '=MATCH("b",UNIQUE(A1:A3),0)'],
      ]);

      expect(cell("C1")).toBe("b");
      expect(cell("C2")).toBe(2);
    });

    test("broadcasts through scalar arithmetic", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 1],
        ["C1", "=UNIQUE(A1:A3)*10"],
      ]);

      expect(cell("C1")).toBe(10);
      expect(cell("C2")).toBe(20);
      expect(cell("C3")).toBe("");
    });

    test("accepts a spilled array as input", () => {
      setContent([
        ["A1", "=SEQUENCE(4,1,1,0)"],
        ["C1", "=UNIQUE(A1:A4)"],
      ]);

      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe("");
    });

    test("accepts a function result directly", () => {
      setContent([["A1", "=UNIQUE(SEQUENCE(3,2,1,0))"]]);

      // Every row of the SEQUENCE is {1, 1}, so one row survives.
      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(1);
      expect(cell("A2")).toBe("");
    });

    test("nests inside itself", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["C1", "=UNIQUE(UNIQUE(A1:A3))"],
      ]);

      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("b");
      expect(cell("C3")).toBe("");
    });
  });

  describe("errors", () => {
    test("propagates an error found in the input", () => {
      setContent([
        ["A1", 1],
        ["A2", "=NOSUCHFN(1)"],
        ["C1", "=UNIQUE(A1:A2)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.NAME);
    });

    test("rejects the wrong number of arguments", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=UNIQUE()"],
        ["C2", "=UNIQUE(A1:A2,FALSE,FALSE,1)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.VALUE);
      expect(cell("C2")).toBe(FormulaError.VALUE);
    });

    test("rejects non-logical by_col and exactly_once", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", '=UNIQUE(A1:A2,"yes")'],
        ["C3", '=UNIQUE(A1:A2,FALSE,"no")'],
      ]);

      expect(cell("C1")).toBe(FormulaError.VALUE);
      expect(cell("C3")).toBe(FormulaError.VALUE);
    });

    test("rejects an array as by_col", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["C1", "=UNIQUE(A1:A2,A1:A2)"],
      ]);

      expect(cell("C1")).toBe(FormulaError.VALUE);
    });

    test("returns #SPILL! when the result area is blocked", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["C1", "=UNIQUE(A1:A2)"],
        ["C2", "blocker"],
      ]);

      expect(cell("C1")).toBe(FormulaError.SPILL);
    });
  });

  describe("recalculation", () => {
    test("reflects an input edit", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["C1", "=UNIQUE(A1:A3)"],
      ]);

      setCellContent("A3", "c");

      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("b");
      expect(cell("C3")).toBe("c");
    });

    test("shrinks when the input gains a duplicate", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "c"],
        ["C1", "=UNIQUE(A1:A3)"],
      ]);

      setCellContent("A3", "a");

      expect(cell("C1")).toBe("a");
      expect(cell("C2")).toBe("b");
      expect(cell("C3")).toBe("");
    });
  });
});
