import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("SUMIF function", () => {
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
    test("exact match criteria - 2 arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 10],
          ["A4", 30],
          ["B1", "=SUMIF(A1:A4, 10)"], // Sum cells equal to 10
        ])
      );

      expect(cell("B1")).toBe(20); // 10 + 10
    });

    test("exact match criteria - 3 arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["B1", "Apple"],
          ["B2", "Banana"],
          ["B3", "Apple"],
          ["C1", '=SUMIF(B1:B3, "Apple", A1:A3)'], // Sum A1:A3 where B1:B3 = "Apple"
        ])
      );

      expect(cell("C1")).toBe(40); // 10 + 30
    });

    test("comparison operators", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["A4", 40],
          ["B1", '=SUMIF(A1:A4, ">15")'], // Greater than 15
          ["B2", '=SUMIF(A1:A4, "<=20")'], // Less than or equal to 20
          ["B3", '=SUMIF(A1:A4, "<>20")'], // Not equal to 20
        ])
      );

      expect(cell("B1")).toBe(90); // 20 + 30 + 40
      expect(cell("B2")).toBe(30); // 10 + 20
      expect(cell("B3")).toBe(80); // 10 + 30 + 40
    });

    test("wildcard patterns", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 100],
          ["A2", 200],
          ["A3", 300],
          ["B1", "Apple"],
          ["B2", "Apricot"],
          ["B3", "Banana"],
          ["C1", '=SUMIF(B1:B3, "Ap*", A1:A3)'], // Starts with "Ap"
          ["C2", '=SUMIF(B1:B3, "?pple", A1:A3)'], // Single char + "pple"
        ])
      );

      expect(cell("C1")).toBe(300); // A1 + A2 = 100 + 200
      expect(cell("C2")).toBe(100); // Only A1 matches "?pple"
    });

    test("no matches found", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["B1", "Apple"],
          ["B2", "Banana"],
          ["C1", '=SUMIF(B1:B2, "Orange", A1:A2)'],
        ])
      );

      expect(cell("C1")).toBe(0); // SUMIF returns 0 when no matches
    });

    test("mixed data types - only numbers summed", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", "text"],
          ["A3", 30],
          ["B1", "Yes"],
          ["B2", "Yes"],
          ["B3", "Yes"],
          ["C1", '=SUMIF(B1:B3, "Yes", A1:A3)'],
        ])
      );

      expect(cell("C1")).toBe(40); // Only A1 and A3 are numeric: 10 + 30
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
          ["B1", 10],
          ["B2", 20],
          ["B3", 30],
          ["B4", 40],
          ["C1", '=SUMIF(A1:A4, "=", B1:B4)'], // Match empty cells
          ["C2", '=SUMIF(A1:A4, "<>", B1:B4)'], // Match non-empty cells
        ])
      );

      expect(cell("C1")).toBe(40); // 10 + 30 = 40 (empty cells in A1, A3)
      expect(cell("C2")).toBe(60); // 20 + 40 = 60 (non-empty cells in A2, A4)
    });
  });

  describe("error handling", () => {
    test("wrong number of arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUMIF(A1:A5)"], // Too few arguments
          ["A2", "=SUMIF(A1:A5, 1, B1:B5, 2)"], // Too many arguments
        ])
      );

      expect(cell("A1")).toBe("#VALUE!");
      expect(cell("A2")).toBe("#VALUE!");
    });

    test("errors in value range are skipped", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", "=CEILING()"], // #VALUE! error (missing arguments)
          ["A3", 30],
          ["B1", "Yes"],
          ["B2", "Yes"],
          ["B3", "Yes"],
          ["C1", '=SUMIF(B1:B3, "Yes", A1:A3)'], // Should skip A2 error and sum A1+A3=40
        ])
      );

      expect(cell("A2")).toBe(FormulaError.VALUE); // Verify A2 is an error
      expect(cell("C1")).toBe(40);
    });

    test("criteria can match literal error strings", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "#VALUE!"], // Literal string, not an error
          ["A2", "text"],
          ["A3", "#VALUE!"], // Another literal string
          ["B1", '=COUNTIF(A1:A3, "#VALUE!")'], // Count literal "#VALUE!" strings
        ])
      );

      expect(cell("B1")).toBe(2); // Should count A1 and A3 (literal strings)
    });
  });
});