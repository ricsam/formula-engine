import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("COUNTIFS function", () => {
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
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["A3", "Apple"],
          ["A4", "Cherry"],
          ["B1", '=COUNTIFS(A1:A4, "Apple")'], // Count cells equal to "Apple"
        ])
      );

      expect(cell("B1")).toBe(2); // A1 and A3
    });

    test("multiple criteria", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["A3", "Apple"],
          ["A4", "Apple"],
          ["B1", 10],
          ["B2", 20],
          ["B3", 30],
          ["B4", 5],
          ["C1", '=COUNTIFS(A1:A4, "Apple", B1:B4, ">10")'], // Apple AND > 10
        ])
      );

      expect(cell("C1")).toBe(1); // Only A3 matches both criteria
    });

    test("comparison operators", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["A4", 40],
          ["B1", '=COUNTIFS(A1:A4, ">15")'], // Greater than 15
          ["B2", '=COUNTIFS(A1:A4, "<=25")'], // Less than or equal to 25
        ])
      );

      expect(cell("B1")).toBe(3); // A2, A3, A4
      expect(cell("B2")).toBe(2); // A1, A2 (A3=30 is not <=25)
    });

    test("no matches found", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["B1", '=COUNTIFS(A1:A2, "Orange")'],
        ])
      );

      expect(cell("B1")).toBe(0); // COUNTIFS returns 0 when no matches
    });

    test("three criteria pairs", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Apple"],
          ["A3", "Apple"],
          ["A4", "Banana"],
          ["B1", 10],
          ["B2", 20],
          ["B3", 30],
          ["B4", 40],
          ["C1", "Red"],
          ["C2", "Blue"],
          ["C3", "Red"],
          ["C4", "Red"],
          ["D1", '=COUNTIFS(A1:A4, "Apple", B1:B4, ">15", C1:C4, "Red")'], // All three criteria
        ])
      );

      expect(cell("D1")).toBe(1); // Only A3 matches all criteria
    });
  });

  describe("empty cell criteria", () => {
    test("should handle empty cell criteria", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", ""], // Empty
          ["A2", "Apple"],
          ["A3", ""], // Empty
          ["A4", "Banana"],
          ["B1", ""], // Empty
          ["B2", "Orange"],
          ["B3", "Grape"],
          ["B4", ""], // Empty
          ["C1", '=COUNTIFS(A1:A4, "=", B1:B4, "=")'], // Count where both A and B are empty
          ["C2", '=COUNTIFS(A1:A4, "<>", B1:B4, "<>")'], // Count where both A and B are non-empty
          ["C3", '=COUNTIFS(A1:A4, "=")'], // Count where A is empty
          ["C4", '=COUNTIFS(A1:A4, "<>")'], // Count where A is non-empty
        ])
      );

      expect(cell("C1")).toBe(1); // Only A1,B1 are both empty
      expect(cell("C2")).toBe(1); // Only A2,B2 are both non-empty
      expect(cell("C3")).toBe(2); // A1 and A3 are empty
      expect(cell("C4")).toBe(2); // A2 and A4 are non-empty
    });
  });

  describe("error handling", () => {
    test("wrong number of arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=COUNTIFS(A1:A5)"], // Too few arguments
          ["A2", "=COUNTIFS(A1:A5, 1, B1:B5)"], // Even number of arguments
        ])
      );

      expect(cell("A1")).toBe("#VALUE!");
      expect(cell("A2")).toBe("#VALUE!");
    });
  });
});
