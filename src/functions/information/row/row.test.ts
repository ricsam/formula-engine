import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("ROW function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("ROW() returns current cell row number", () => {
      setCellContent("B5", "=ROW()");
      expect(cell("B5")).toBe(5);
    });

    test("ROW() in different rows returns correct row numbers", () => {
      setCellContent("A1", "=ROW()");
      setCellContent("A10", "=ROW()");
      setCellContent("A100", "=ROW()");
      
      expect(cell("A1")).toBe(1);
      expect(cell("A10")).toBe(10);
      expect(cell("A100")).toBe(100);
    });

    test("ROW() works regardless of column", () => {
      setCellContent("A5", "=ROW()");
      setCellContent("Z5", "=ROW()");
      
      expect(cell("A5")).toBe(5);
      expect(cell("Z5")).toBe(5);
    });
  });

  describe("error handling", () => {
    test("ROW with too many arguments returns error", () => {
      setCellContent("A1", "=ROW(A1, A2)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  // These tests are now implemented
  describe("with reference argument", () => {
    test("ROW(A5) should return 5", () => {
      setCellContent("A1", "=ROW(A5)");
      expect(cell("A1")).toBe(5);
    });

    test("ROW(B10) should return 10", () => {
      setCellContent("A1", "=ROW(B10)");
      expect(cell("A1")).toBe(10);
    });
  });

  describe("with range argument", () => {
    test("ROW(A5:A10) should return array of row numbers", () => {
      setCellContent("A1", "=ROW(A5:A10)");
      // Should spill values 5, 6, 7, 8, 9, 10 down from A1
      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe(6);
      expect(cell("A3")).toBe(7);
      expect(cell("A4")).toBe(8);
      expect(cell("A5")).toBe(9);
      expect(cell("A6")).toBe(10);
    });

    test("ROW(B2:D4) should return array of row numbers", () => {
      setCellContent("B1", "=ROW(B2:D4)");
      // Should spill values 2, 3, 4 down from B1
      expect(cell("B1")).toBe(2);
      expect(cell("B2")).toBe(3);
      expect(cell("B3")).toBe(4);
    });
  });
});
