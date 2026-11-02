import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("COLUMNS function", () => {
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
    test("counts the columns of a range", () => {
      setContent([
        ["A1", "=COLUMNS(B1:D10)"],
        ["A2", "=COLUMNS(C3:C4)"],
        ["A3", "=COLUMNS(E7:E7)"],
        ["A4", "=COLUMNS($B$2:F2)"],
      ]);

      expect(cell("A1")).toBe(3);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(5);
    });

    test("is also callable as COLS", () => {
      setContent([
        ["A1", "=COLS(B1:E1)"],
        ["A2", "=cols(SEQUENCE(2, 3))"],
      ]);

      expect(cell("A1")).toBe(4);
      expect(cell("A2")).toBe(3);
    });

    test("does not read the cells of the range", () => {
      setContent([
        ["A1", "=CEILING()"],
        ["B1", "text"],
        // A formula inside its own range is not circular: no cell is read.
        ["C1", "=COLUMNS(A1:D1)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("C1")).toBe(4);

      setCellContent("B1", "=SEQUENCE(1, 3)");
      expect(cell("C1")).toBe(4);
    });

    test("counts unbounded columns as INFINITY", () => {
      setContent([
        ["A1", "=COLUMNS(1:1)"],
        ["A2", "=COLUMNS(B2:2)"],
        ["A3", "=COLUMNS(B2:INFINITY)"],
        ["A4", "=COLUMNS(B:B)"],
        ["A5", "=COLUMNS(B:D)"],
      ]);

      expect(cell("A1")).toBe("INFINITY");
      expect(cell("A2")).toBe("INFINITY");
      expect(cell("A3")).toBe("INFINITY");
      expect(cell("A4")).toBe(1);
      expect(cell("A5")).toBe(3);
    });

    test("counts the columns of a range on another sheet", () => {
      engine.addSheet({ workbookName, sheetName: "Data" });
      setCellContent("A1", "=COLUMNS(Data!A1:C4)");

      expect(cell("A1")).toBe(3);
    });
  });

  describe("arrays and spilled results", () => {
    test("counts the columns of SEQUENCE results", () => {
      setContent([
        ["A1", "=COLUMNS(SEQUENCE(5))"],
        ["A2", "=COLUMNS(SEQUENCE(3, 4))"],
        ["A3", "=COLUMNS(SEQUENCE(1, 6))"],
        ["A4", "=COLUMNS(SEQUENCE(1, INFINITY))"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(4);
      expect(cell("A3")).toBe(6);
      expect(cell("A4")).toBe("INFINITY");
    });

    test("counts the columns of array literals", () => {
      setContent([
        ["A1", "=COLUMNS({1,2,3;4,5,6})"],
        ["A2", "=COLUMNS({1,2,3})"],
        ["A3", "=COLUMNS({1;2;3})"],
      ]);

      expect(cell("A1")).toBe(3);
      expect(cell("A2")).toBe(3);
      expect(cell("A3")).toBe(1);
    });

    test("counts the columns of a UNIQUE result and recalculates", () => {
      setContent([
        ["A1", "a"],
        ["B1", "b"],
        ["C1", "a"],
        ["D1", "c"],
        ["A3", "=COLUMNS(UNIQUE(A1:D1, TRUE))"],
        ["A4", "=ROWS(UNIQUE(A1:D1, TRUE))"],
      ]);
      expect(cell("A3")).toBe(3);
      expect(cell("A4")).toBe(1);

      setCellContent("C1", "d");
      expect(cell("A3")).toBe(4);

      setCellContent("B1", "a");
      setCellContent("C1", "a");
      setCellContent("D1", "a");
      // a, a, a, a collapses to a single value
      expect(cell("A3")).toBe(1);
    });

    test("measures the whole spill of a referenced spilling cell", () => {
      setContent([
        ["A1", 5],
        ["C1", "=SEQUENCE(1, A1)"],
        ["C3", "=COLUMNS(C1)"],
        ["C4", "=COLUMNS(D1)"],
      ]);

      expect(cell("C3")).toBe(5);
      // D1 is one cell of the spill, not its origin.
      expect(cell("C4")).toBe(1);

      setCellContent("A1", 2);
      expect(cell("C3")).toBe(2);

      setCellContent("C1", "seven");
      expect(cell("C3")).toBe(1);
    });
  });

  describe("single values", () => {
    test("counts a single value as one column", () => {
      setContent([
        ["B1", 42],
        ["A1", "=COLUMNS(5)"],
        ["A2", '=COLUMNS("text")'],
        ["A3", "=COLUMNS(B1)"],
        // An empty cell is still one cell.
        ["A4", "=COLUMNS(B2)"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(1);
      expect(cell("A4")).toBe(1);
    });
  });

  describe("errors", () => {
    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=COLUMNS()"],
        ["A2", "=COLUMNS(B1:B2, C1:C2)"],
        ["A3", "=COLS()"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("propagates an error argument", () => {
      setContent([
        ["B1", "x"],
        ["C1", "x"],
        ["A1", "=COLUMNS(CEILING())"],
        // No column occurs exactly once, so UNIQUE returns #CALC!.
        ["A2", "=COLUMNS(UNIQUE(B1:C1, TRUE, TRUE))"],
        ["D2", "=CEILING(5, 0)"],
        ["A3", "=COLUMNS(D2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.CALC);
      expect(cell("A3")).toBe(FormulaError.DIV0);

      setCellContent("C1", "y");
      expect(cell("A2")).toBe(2);
    });
  });
});
