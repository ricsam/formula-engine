import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("OFFSET function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string | number) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("error handling", () => {
    test("OFFSET with too few arguments returns error", () => {
      setCellContent("B1", "=OFFSET(A1, 1)");
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with too many arguments returns error", () => {
      setCellContent("B1", "=OFFSET(A1, 1, 1, 1, 1, 1)");
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with non-numeric rows returns error", () => {
      setCellContent("B1", '=OFFSET(A1, "text", 1)');
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with non-numeric cols returns error", () => {
      setCellContent("B1", '=OFFSET(A1, 1, "text")');
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with zero height returns error", () => {
      setCellContent("B1", "=OFFSET(A1, 0, 0, 0)");
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with negative height returns error", () => {
      setCellContent("B1", "=OFFSET(A1, 0, 0, -1)");
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with zero width returns error", () => {
      setCellContent("B1", "=OFFSET(A1, 0, 0, 1, 0)");
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("OFFSET with negative width returns error", () => {
      setCellContent("B1", "=OFFSET(A1, 0, 0, 1, -1)");
      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  // These tests are now implemented
  describe("basic functionality", () => {
    test("OFFSET with positive row offset", () => {
      setCellContent("A1", 10);
      setCellContent("A3", 30);
      setCellContent("B1", "=OFFSET(A1, 2, 0)");
      expect(cell("B1")).toBe(30);
    });

    test("OFFSET with positive column offset", () => {
      setCellContent("A1", 10);
      setCellContent("D1", 40);
      setCellContent("B1", "=OFFSET(A1, 0, 3)");
      expect(cell("B1")).toBe(40);
    });

    test("OFFSET with negative offsets", () => {
      setCellContent("D5", 99);
      setCellContent("C3", 88);
      setCellContent("E6", "=OFFSET(D5, -2, -1)");
      expect(cell("E6")).toBe(88);
    });

    test("OFFSET with height and width creates range", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 2],
          ["A3", 3],
          ["C1", "=OFFSET(A1, 0, 0, 3, 1)"],
        ])
      );
      // Should spill values 1, 2, 3 down from C1
      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(2);
      expect(cell("C3")).toBe(3);
    });

    test("OFFSET with range as reference", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 20],
          ["A2", 30],
          ["B2", 40],
          ["C3", 50],
          ["D1", "=OFFSET(A1:B2, 2, 2)"],
        ])
      );
      // A1:B2 offset by (2, 2) = C3:D4, first cell is C3
      expect(cell("D1")).toBe(50);
    });

    test("OFFSET returning range that spills", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 2],
          ["B1", 3],
          ["B2", 4],
          ["D1", "=OFFSET(A1, 0, 0, 2, 2)"],
        ])
      );
      // Should spill a 2x2 array starting at D1
      expect(cell("D1")).toBe(1);
      expect(cell("E1")).toBe(3);
      expect(cell("D2")).toBe(2);
      expect(cell("E2")).toBe(4);
    });
  });

  describe("edge cases", () => {
    test("OFFSET going out of bounds returns #REF!", () => {
      setCellContent("B1", "=OFFSET(A1, -1, 0)");
      expect(cell("B1")).toBe(FormulaError.REF);
    });

    test("OFFSET with decimal offsets rounds down", () => {
      setCellContent("A1", 10);
      setCellContent("A3", 30);
      setCellContent("B1", "=OFFSET(A1, 2.9, 0)");
      expect(cell("B1")).toBe(30);
    });
  });
});
