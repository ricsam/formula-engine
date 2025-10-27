import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("SUMIFS function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("single criteria", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["B1", "Apple"],
          ["B2", "Banana"],
          ["B3", "Apple"],
          ["C1", '=SUMIFS(A1:A3, B1:B3, "Apple")'], // Sum A1:A3 where B1:B3 = "Apple"
        ])
      );

      expect(cell("C1")).toBe(40); // 10 + 30
    });

    test("multiple criteria", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10], // values to sum
          ["A2", 20],
          ["A3", 30],
          ["A4", 40],
          ["B1", "Apple"], // first criteria range
          ["B2", "Banana"],
          ["B3", "Apple"],
          ["B4", "Apple"],
          ["C1", 5], // second criteria range
          ["C2", 15],
          ["C3", 25],
          ["C4", 35],
          ["D1", '=SUMIFS(A1:A4, B1:B4, "Apple", C1:C4, ">10")'], // Apple AND > 10
        ])
      );

      expect(cell("D1")).toBe(70); // Only A3 and A4 match both criteria: 30 + 40
    });

    test("comparison operators", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["A4", 40],
          ["B1", 5],
          ["B2", 15],
          ["B3", 25],
          ["B4", 35],
          ["C1", '=SUMIFS(A1:A4, B1:B4, ">=20")'], // >= 20
          ["C2", '=SUMIFS(A1:A4, B1:B4, "<>15")'], // != 15
        ])
      );

      expect(cell("C1")).toBe(70); // A3 and A4: 30 + 40
      expect(cell("C2")).toBe(80); // A1, A3, A4: 10 + 30 + 40
    });

    test("no matches found", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["B1", "Apple"],
          ["B2", "Banana"],
          ["C1", '=SUMIFS(A1:A2, B1:B2, "Orange")'],
        ])
      );

      expect(cell("C1")).toBe(0); // SUMIFS returns 0 when no matches
    });

    test("wildcard criteria", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 100],
          ["A2", 200],
          ["A3", 300],
          ["B1", "Apple"],
          ["B2", "Apricot"],
          ["B3", "Banana"],
          ["C1", '=SUMIFS(A1:A3, B1:B3, "Ap*")'], // Starts with "Ap"
        ])
      );

      expect(cell("C1")).toBe(300); // A1 + A2 = 100 + 200
    });
  });

  describe("empty cell criteria", () => {
    test("should handle empty cell criteria", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", ""],      // Empty
          ["A2", "Apple"], // Non-empty
          ["A3", ""],      // Empty
          ["A4", "Banana"], // Non-empty
          ["B1", 100],
          ["B2", 200],
          ["B3", 150],
          ["B4", 300],
          ["C1", '=SUMIFS(B1:B4, A1:A4, "=")'],  // Sum where A is empty
          ["C2", '=SUMIFS(B1:B4, A1:A4, "")'],   // Sum where A is empty (equivalent)
          ["C3", '=SUMIFS(B1:B4, A1:A4, "<>")'], // Sum where A is not empty
        ])
      );

      expect(cell("C1")).toBe(250); // Sum of B1, B3 = 250 (where A1, A3 are empty)
      expect(cell("C2")).toBe(250); // Same as C1
      expect(cell("C3")).toBe(500); // Sum of B2, B4 = 500 (where A2, A4 are non-empty)
    });
  });

  describe("error handling", () => {
    test("wrong number of arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUMIFS(A1:A5, B1:B5)"], // Missing criteria
          ["A2", "=SUMIFS(A1:A5, B1:B5, 1, C1:C5)"], // Even number of arguments
        ])
      );

      expect(cell("A1")).toBe("#VALUE!");
      expect(cell("A2")).toBe("#VALUE!");
    });
  });
});