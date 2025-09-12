import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import {
  FormulaError,
  type SerializedCellValue,
} from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("INDEX function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should return value from single cell array", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "=INDEX(A1, 1, 1)"],
        ])
      );

      expect(cell("B1")).toBe("Apple");
    });

    test("should return value from single cell array with default column", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "=INDEX(A1, 1)"],
        ])
      );

      expect(cell("B1")).toBe("Apple");
    });

    test("should return value from row array", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "Cherry"],
          ["D1", "=INDEX(A1:C1, 1, 2)"],
        ])
      );
      expect(cell("D1", true)).toBe("Banana");
    });

    test("should return value from column array", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", "=INDEX(A1:A3, 2, 1)");

      expect(cell("B1")).toBe("Banana");
    });

    test("should return value from 2D array", () => {
      // Set up 2x3 array
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "Cherry"],
          ["A2", 10],
          ["B2", 20],
          ["C2", 30],
        ])
      );
      setCellContent("D1", "=INDEX(A1:C2, 2, 3)");

      expect(cell("D1")).toBe(30);
    });

    test("should handle large array", () => {
      // Set up 3x3 array
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["B1", 2],
          ["C1", 3],
          ["A2", 4],
          ["B2", 5],
          ["C2", 6],
          ["A3", 7],
          ["B3", 8],
          ["C3", 9],
        ])
      );
      setCellContent("D1", "=INDEX(A1:C3, 3, 2)");

      expect(cell("D1")).toBe(8);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return #VALUE! for string row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=INDEX(A1:A2, "first", 1)');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for string column_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "Banana");
      setCellContent("C1", '=INDEX(A1:B1, 1, "second")');

      expect(cell("C1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for boolean row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, TRUE, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for infinity row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, INFINITY, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("bounds checking", () => {
    test("should return #REF! for row_num out of bounds (too high)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, 3, 1)");

      expect(cell("B1")).toBe(FormulaError.REF);
    });

    test("should return #REF! for column_num out of bounds (too high)", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "=INDEX(A1:B1, 1, 3)"],
        ])
      );

      expect(cell("C1")).toBe(FormulaError.REF);
    });

    test("should return #VALUE! for row_num less than 1", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, 0, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for negative row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, -1, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #REF! for single cell with wrong indices", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "=INDEX(A1, 2, 1)");

      expect(cell("B1")).toBe(FormulaError.REF);
    });
  });

  describe("error handling", () => {
    test("should return #VALUE! for wrong number of arguments", () => {
      setCellContent("B1", "=INDEX()");
      expect(cell("B1")).toBe(FormulaError.VALUE);

      setCellContent("B2", "=INDEX(A1)");
      expect(cell("B2")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for too many arguments", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "=INDEX(A1, 1, 1, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should handle decimal row_num (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", "=INDEX(A1:A3, 2.9, 1)");

      expect(cell("B1")).toBe("Banana");
    });

    test("should handle decimal column_num (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "Banana");
      setCellContent("C1", "Cherry");
      setCellContent("D1", "=INDEX(A1:C1, 1, 2.7)");

      expect(cell("D1")).toBe("Banana");
    });
  });

  describe("edge cases", () => {
    test("should work with mixed data types in array", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Text"],
          ["A2", 42],
          ["A3", true],
        ])
      );
      setCellContent("B1", "=INDEX(A1:A3, 2, 1)");
      setCellContent("B2", "=INDEX(A1:A3, 3, 1)");

      expect(cell("B1")).toBe(42);
      expect(cell("B2")).toBe(true);
    });

    test("should handle empty cells in array", () => {
      setCellContent("A1", "Apple");
      // A2 is empty
      setCellContent("A3", "Cherry");
      setCellContent("B1", "=INDEX(A1:A3, 2, 1)");

      expect(cell("B1")).toBe("");
    });
  });
});
