import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("MATCH function", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  describe("basic functionality", () => {
    test("should find exact match with match_type 0", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Banana", A1:A3, 0)');

      expect(cell("B1")).toBe(2);
    });

    test("should find exact match for numbers", () => {
      // Use SerializedCellValue to set numeric values properly
      engine.setSheetContent(
        sheetName,
        new Map([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
        ])
      );
      setCellContent("B1", "=MATCH(20, A1:A3, 0)");

      expect(cell("B1")).toBe(2);
    });

    test("should default to match_type 1", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Banana", A1:A3)');

      expect(cell("B1")).toBe(2);
    });

    test("should return #N/A when not found", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Grape", A1:A3, 0)');

      expect(cell("B1")).toBe(FormulaError.NA);
    });

    test("should be case-sensitive", () => {
      setCellContent("A1", "apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Apple", A1:A3, 0)');

      expect(cell("B1")).toBe(FormulaError.NA);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return #VALUE! for boolean lookup_value", () => {
      setCellContent("A1", "TRUE");
      setCellContent("A2", "FALSE");
      setCellContent("B1", "=MATCH(TRUE, A1:A2, 0)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for infinity lookup_value", () => {
      setCellContent("A1", "1");
      setCellContent("A2", "2");
      setCellContent("B1", "=MATCH(INFINITY, A1:A2, 0)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for invalid match_type", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, 2)');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for string match_type", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, "exact")');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("error handling", () => {
    test("should return #VALUE! for wrong number of arguments", () => {
      setCellContent("B1", "=MATCH()");
      expect(cell("B1")).toBe(FormulaError.VALUE);

      setCellContent("B2", '=MATCH("Apple")');
      expect(cell("B2")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for too many arguments", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=MATCH("Apple", A1:A1, 0, "extra")');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should handle decimal match_type (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, 0.9)');

      expect(cell("B1")).toBe(1);
    });
  });

  describe("edge cases", () => {
    test("should handle single cell array", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=MATCH("Apple", A1, 0)');

      expect(cell("B1")).toBe(1);
    });

    test("should handle mixed string and number types (strict checking)", () => {
      // Set up mixed array with proper types
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", 10],
        ])
      );
      setCellContent("B1", "=MATCH(10, A1:A2, 0)"); // Number lookup in mixed array

      expect(cell("B1")).toBe(2);
    });
  });

  describe.skip("can use table column as lookup_array", () => {
    test("should find exact match with match_type 0", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Fruit"],
          ["A2", "Stock"],
          ["A3", "Is local"],
          ["A2", "Apple"],
          ["B2", 2],
          ["C2", "Yes"],
          ["A2", "Orange"],
          ["B2", 3],
          ["C2", "No"],
          ["A3", "Banana"],
          ["B3", 1],
          ["C3", "Maybe"],
          ["A4", "Kiwi"],
          ["B4", 4],
          ["C4", "Yes"],
          ["A5", "Pineapple"],
          ["B5", 5],
          ["C5", "No"],
          ["A6", "Pear"],
          ["B6", 6],
          ["C6", "Yes"],
          ["A7", "Strawberry"],
          ["B7", 7],
          ["C7", "No"],
          ["A8", "Watermelon"],
          ["B8", 8],
          ["C8", "Yes"],
          ["A9", "Mango"],
          ["B9", 9],
          ["C9", "No"],
          ["A10", "Pomegranate"],

          ["K1", `=MATCH("Kiwi", A2:A10, 0)`],
          ["L1", `=MATCH("Mango", A:A, 0)`],
        ])
      );

      expect(cell("K1")).toBe(3);
      expect(cell("L1", true)).toBe(8);
    });
  });
});
