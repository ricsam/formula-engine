import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("MAXIF function", () => {
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
          ["B1", "=MAXIF(A1:A4, 10)"], // Max of cells equal to 10
        ])
      );

      expect(cell("B1")).toBe(10); // max(10, 10) = 10
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
          ["C1", '=MAXIF(B1:B3, "Apple", A1:A3)'], // Max of A1:A3 where B1:B3 = "Apple"
        ])
      );

      expect(cell("C1")).toBe(30); // max(10, 30) = 30
    });

    test("comparison operators", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["A4", 40],
          ["B1", '=MAXIF(A1:A4, ">15")'], // Greater than 15
          ["B2", '=MAXIF(A1:A4, "<=25")'], // Less than or equal to 25
        ])
      );

      expect(cell("B1")).toBe(40); // max(20, 30, 40) = 40
      expect(cell("B2")).toBe(20); // max(10, 20) = 20
    });

    test("no matches found", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["B1", "Apple"],
          ["B2", "Banana"],
          ["C1", '=MAXIF(B1:B2, "Orange", A1:A2)'],
        ])
      );

      expect(cell("C1")).toBe("#VALUE!"); // MAXIF returns error when no matches
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
          ["C1", '=MAXIF(A1:A4, "=", B1:B4)'], // Max of values where criteria is empty
          ["C2", '=MAXIF(A1:A4, "<>", B1:B4)'], // Max of values where criteria is non-empty
        ])
      );

      expect(cell("C1")).toBe(30); // max(10, 30) = 30 (empty cells in A1, A3)
      expect(cell("C2")).toBe(40); // max(20, 40) = 40 (non-empty cells in A2, A4)
    });
  });

  describe("error handling", () => {
    test("wrong number of arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=MAXIF(A1:A5)"], // Too few arguments
          ["A2", "=MAXIF(A1:A5, 1, B1:B5, 2)"], // Too many arguments
        ])
      );

      expect(cell("A1")).toBe("#VALUE!");
      expect(cell("A2")).toBe("#VALUE!");
    });
  });
});