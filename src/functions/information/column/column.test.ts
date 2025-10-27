import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("COLUMN function", () => {
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
    test("COLUMN() returns current cell column number", () => {
      setCellContent("C5", "=COLUMN()");
      expect(cell("C5")).toBe(3);
    });

    test("COLUMN() in different columns returns correct column numbers", () => {
      setCellContent("A1", "=COLUMN()");
      setCellContent("B1", "=COLUMN()");
      setCellContent("Z1", "=COLUMN()");
      
      expect(cell("A1")).toBe(1);
      expect(cell("B1")).toBe(2);
      expect(cell("Z1")).toBe(26);
    });

    test("COLUMN() works regardless of row", () => {
      setCellContent("E1", "=COLUMN()");
      setCellContent("E100", "=COLUMN()");
      
      expect(cell("E1")).toBe(5);
      expect(cell("E100")).toBe(5);
    });
  });

  describe("error handling", () => {
    test("COLUMN with too many arguments returns error", () => {
      setCellContent("A1", "=COLUMN(A1, A2)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  // These tests are now implemented
  describe("with reference argument", () => {
    test("COLUMN(C5) should return 3", () => {
      setCellContent("A1", "=COLUMN(C5)");
      expect(cell("A1")).toBe(3);
    });

    test("COLUMN(Z10) should return 26", () => {
      setCellContent("A1", "=COLUMN(Z10)");
      expect(cell("A1")).toBe(26);
    });
  });

  describe("with range argument", () => {
    test("COLUMN(A1:E1) should return array of column numbers", () => {
      setCellContent("A2", "=COLUMN(A1:E1)");
      // Should spill values 1, 2, 3, 4, 5 across from A2
      expect(cell("A2")).toBe(1);
      expect(cell("B2")).toBe(2);
      expect(cell("C2")).toBe(3);
      expect(cell("D2")).toBe(4);
      expect(cell("E2")).toBe(5);
    });

    test("COLUMN(B2:D4) should return array of column numbers", () => {
      setCellContent("A1", "=COLUMN(B2:D4)");
      // Should spill values 2, 3, 4 across from A1
      expect(cell("A1")).toBe(2);
      expect(cell("B1")).toBe(3);
      expect(cell("C1")).toBe(4);
    });
  });
});
