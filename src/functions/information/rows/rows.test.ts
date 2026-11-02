import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ROWS function", () => {
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

  describe("ranges", () => {
    test("counts the rows of a range", () => {
      setContent([
        ["A1", "=ROWS(B1:B10)"],
        ["A2", "=ROWS(B3:D4)"],
        ["A3", "=ROWS(C7:C7)"],
        ["A4", "=ROWS($B$2:B5)"],
      ]);

      expect(cell("A1")).toBe(10);
      expect(cell("A2")).toBe(2);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(4);
    });

    test("does not read the cells of the range", () => {
      setContent([
        ["B1", "=CEILING()"],
        ["B2", "text"],
        // A formula inside its own range is not circular: no cell is read.
        ["B3", "=ROWS(B1:B5)"],
      ]);

      expect(cell("B1")).toBe(FormulaError.VALUE);
      expect(cell("B3")).toBe(5);
    });

    test("does not change when cells inside the range are edited", () => {
      setContent([
        ["A1", "=ROWS(B1:B3)"],
        ["B1", 1],
      ]);
      expect(cell("A1")).toBe(3);

      setCellContent("B2", "=SEQUENCE(10)");
      setCellContent("B1", "");
      expect(cell("A1")).toBe(3);
    });

    test("counts unbounded rows as INFINITY", () => {
      setContent([
        ["A1", "=ROWS(B:B)"],
        ["A2", "=ROWS(B2:B)"],
        ["A3", "=ROWS(B2:INFINITY)"],
        ["A4", "=ROWS(1:1)"],
        ["A5", "=ROWS(2:4)"],
      ]);

      expect(cell("A1")).toBe("INFINITY");
      expect(cell("A2")).toBe("INFINITY");
      expect(cell("A3")).toBe("INFINITY");
      expect(cell("A4")).toBe(1);
      expect(cell("A5")).toBe(3);
    });

    test("counts the rows of a range on another sheet", () => {
      engine.addSheet({ workbookName, sheetName: "Data" });
      setCellContent("A1", "=ROWS(Data!A1:C4)");

      expect(cell("A1")).toBe(4);
    });
  });

  describe("arrays and spilled results", () => {
    test("counts the rows of SEQUENCE results", () => {
      setContent([
        ["A1", "=ROWS(SEQUENCE(5))"],
        ["A2", "=ROWS(SEQUENCE(3, 4))"],
        ["A3", "=ROWS(SEQUENCE(1, 4))"],
        ["A4", "=ROWS(SEQUENCE(1))"],
        ["A5", "=ROWS(SEQUENCE(INFINITY))"],
      ]);

      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe(3);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(1);
      expect(cell("A5")).toBe("INFINITY");
    });

    test("counts the rows of array literals", () => {
      setContent([
        ["A1", "=ROWS({1,2,3;4,5,6})"],
        ["A2", "=ROWS({1,2,3})"],
        ["A3", "=ROWS({1;2;3})"],
      ]);

      expect(cell("A1")).toBe(2);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(3);
    });

    test("counts the rows of a UNIQUE result and recalculates", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["A3", "a"],
        ["A4", "c"],
        ["A5", "b"],
        ["C1", "=ROWS(UNIQUE(A1:A5))"],
      ]);
      expect(cell("C1")).toBe(3);

      setCellContent("A3", "d");
      expect(cell("C1")).toBe(4);

      setCellContent("A2", "a");
      setCellContent("A4", "a");
      setCellContent("A5", "a");
      // a, a, d, a, a
      expect(cell("C1")).toBe(2);
    });

    test("drives SEQUENCE(ROWS(...)) over a UNIQUE result", () => {
      setContent([
        ["A1", "x"],
        ["A2", "y"],
        ["A3", "x"],
        ["A4", "z"],
        ["C1", "=SEQUENCE(ROWS(UNIQUE(A1:A4)))"],
        ["E1", "=SEQUENCE(1, ROWS(UNIQUE(A1:A4)))"],
      ]);

      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(2);
      expect(cell("C3")).toBe(3);
      expect(cell("C4")).toBe("");
      expect(cell("E1")).toBe(1);
      expect(cell("F1")).toBe(2);
      expect(cell("G1")).toBe(3);
      expect(cell("H1")).toBe("");

      setCellContent("A5", "w");
      // A5 is outside the UNIQUE range, so nothing changes.
      expect(cell("C4")).toBe("");

      setCellContent("A3", "w");
      expect(cell("C4")).toBe(4);
      expect(cell("C5")).toBe("");
      expect(cell("H1")).toBe(4);
    });

    test("measures the whole spill of a referenced spilling cell", () => {
      setContent([
        ["A1", 5],
        ["C1", "=SEQUENCE(A1)"],
        ["E1", "=ROWS(C1)"],
        ["E2", "=ROWS(C2)"],
      ]);

      expect(cell("E1")).toBe(5);
      // C2 is one cell of the spill, not its origin.
      expect(cell("E2")).toBe(1);

      setCellContent("A1", 2);
      expect(cell("E1")).toBe(2);

      setCellContent("C1", 7);
      expect(cell("E1")).toBe(1);

      setCellContent("C1", "=SEQUENCE(4, 2)");
      expect(cell("E1")).toBe(4);
    });

    test("propagates #SPILL! from a referenced spill that is blocked", () => {
      setContent([
        ["C1", "=SEQUENCE(3)"],
        ["C3", "blocker"],
        ["E1", "=ROWS(C1)"],
      ]);
      expect(cell("E1")).toBe(FormulaError.SPILL);

      setCellContent("C3", "");
      expect(cell("E1")).toBe(3);
    });

    test("measures a referenced UNIQUE result", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 2],
        ["C1", "=UNIQUE(A1:A3)"],
        ["E1", "=ROWS(C1)"],
      ]);
      expect(cell("E1")).toBe(2);

      setCellContent("A3", 3);
      expect(cell("E1")).toBe(3);
    });
  });

  describe("single values", () => {
    test("counts a single value as one row", () => {
      setContent([
        ["B1", 42],
        ["B2", "text"],
        ["A1", "=ROWS(5)"],
        ["A2", '=ROWS("text")'],
        ["A3", "=ROWS(TRUE)"],
        ["A4", "=ROWS(B1)"],
        ["A5", "=ROWS(B2)"],
        // An empty cell is still one cell.
        ["A6", "=ROWS(B3)"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(1);
      expect(cell("A5")).toBe(1);
      expect(cell("A6")).toBe(1);
    });
  });

  describe("errors", () => {
    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=ROWS()"],
        ["A2", "=ROWS(B1:B2, C1:C2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("propagates an error argument", () => {
      setContent([
        ["B1", "x"],
        ["B2", "x"],
        ["A1", "=ROWS(CEILING())"],
        // No value occurs exactly once, so UNIQUE returns #CALC!.
        ["A2", "=ROWS(UNIQUE(B1:B2, FALSE, TRUE))"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.CALC);

      setCellContent("B2", "y");
      expect(cell("A2")).toBe(2);
    });

    test("propagates the error of a referenced cell", () => {
      // Single-cell references are evaluated to detect spills, so the
      // referenced error propagates (Excel would return 1).
      setContent([
        ["B1", "=CEILING(5, 0)"],
        ["A1", "=ROWS(B1)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.DIV0);

      setCellContent("B1", 3);
      expect(cell("A1")).toBe(1);
    });
  });
});
