import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("COUNTIF function", () => {
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
    test("should count exact string matches", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["A3", "Apple"],
          ["A4", "Cherry"],
          ["A5", "Apple"],
          ["B1", '=COUNTIF(A1:A5, "Apple")'],
        ])
      );

      expect(cell("B1")).toBe(3);
    });

    test("should count exact number matches", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 10],
          ["A4", 30],
          ["A5", 10],
          ["B1", "=COUNTIF(A1:A5, 10)"],
        ])
      );

      expect(cell("B1")).toBe(3);
    });

    test("should be case-sensitive for strings", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "apple"],
          ["A3", "APPLE"],
          ["B1", '=COUNTIF(A1:A3, "Apple")'],
        ])
      );

      expect(cell("B1")).toBe(1);
    });

    test("should handle single cell range", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=COUNTIF(A1, "Apple")');

      expect(cell("B1")).toBe(1);
    });

    test("should return 0 when no matches found", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["A3", "Cherry"],
          ["B1", '=COUNTIF(A1:A3, "Orange")'],
        ])
      );

      expect(cell("B1")).toBe(0);
    });
  });

  describe("comparison operators", () => {
    test("should count values greater than criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 15],
          ["A3", 25],
          ["A4", 8],
          ["A5", 35],
          ["B1", '=COUNTIF(A1:A5, ">10")'],
        ])
      );

      expect(cell("B1")).toBe(3); // 15, 25, 35
    });

    test("should count values less than criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 15],
          ["A3", 25],
          ["A4", 8],
          ["A5", 35],
          ["B1", '=COUNTIF(A1:A5, "<10")'],
        ])
      );

      expect(cell("B1")).toBe(2); // 5, 8
    });

    test("should count values greater than or equal to criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 10],
          ["A3", 15],
          ["A4", 8],
          ["A5", 10],
          ["B1", '=COUNTIF(A1:A5, ">=10")'],
        ])
      );

      expect(cell("B1")).toBe(3); // 10, 15, 10
    });

    test("should count values less than or equal to criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 10],
          ["A3", 15],
          ["A4", 8],
          ["A5", 10],
          ["B1", '=COUNTIF(A1:A5, "<=10")'],
        ])
      );

      expect(cell("B1")).toBe(4); // 5, 10, 8, 10
    });

    test("should count values not equal to criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 10],
          ["A3", 15],
          ["A4", 10],
          ["A5", 20],
          ["B1", '=COUNTIF(A1:A5, "<>10")'],
        ])
      );

      expect(cell("B1")).toBe(3); // 5, 15, 20 (parser correctly identifies 10 as number)
    });

    test("should count non-empty cells using '<>'", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", ""], // Empty string
          ["A3", "Banana"],
          ["A4", ""], // Empty string
          ["A5", "Cherry"],
          ["B1", '=COUNTIF(A1:A5, "<>")'],
        ])
      );

      expect(cell("B1")).toBe(3); // Apple, Banana, Cherry (non-empty strings)
    });

    test("should count non-zero values using string criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 0],
          ["A3", -3],
          ["A4", 0],
          ["A5", 10],
          ["B1", '=COUNTIF(A1:A5, "<>0")'], // String "0" vs numbers
          ["B2", "=COUNTIF(A1:A5, 0)"],     // Number 0 for exact match
        ])
      );

      expect(cell("B1")).toBe(3); // 5, -3, 10 (parser correctly identifies 0 as number)
      expect(cell("B2")).toBe(2); // Two number zeros
    });

    test("should handle decimal comparisons", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 3.14],
          ["A2", 2.71],
          ["A3", 4.5],
          ["A4", 1.0],
          ["B1", '=COUNTIF(A1:A4, ">3")'],
        ])
      );

      expect(cell("B1")).toBe(2); // 3.14, 4.5
    });
  });

  describe("wildcard patterns", () => {
    test("should match asterisk wildcard", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Application"],
          ["A3", "Banana"],
          ["A4", "Apricot"],
          ["B1", '=COUNTIF(A1:A4, "App*")'],
        ])
      );

      expect(cell("B1")).toBe(2); // Apple, Application
    });

    test("should match question mark wildcard", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Cat"],
          ["A2", "Bat"],
          ["A3", "Hat"],
          ["A4", "Boat"],
          ["B1", '=COUNTIF(A1:A4, "?at")'],
        ])
      );

      expect(cell("B1")).toBe(3); // Cat, Bat, Hat
    });

    test("should combine wildcards", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "test1.txt"],
          ["A2", "test2.doc"],
          ["A3", "test3.txt"],
          ["A4", "demo.txt"],
          ["B1", '=COUNTIF(A1:A4, "test?.txt")'],
        ])
      );

      expect(cell("B1")).toBe(2); // test1.txt, test3.txt
    });

    test("should handle wildcards with special regex characters", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "test[1]"],
          ["A2", "test(2)"],
          ["A3", "test{3}"],
          ["A4", "test.4"],
          ["B1", '=COUNTIF(A1:A4, "test*")'],
        ])
      );

      expect(cell("B1")).toBe(4); // All should match
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should NOT compare string numbers with numeric criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "5"],   // String, not number
          ["A2", "15"],  // String, not number
          ["A3", "25"],  // String, not number
          ["B1", '=COUNTIF(A1:A3, ">10")'],
        ])
      );

      expect(cell("B1")).toBe(0); // No matches - strings can't be compared with >
    });

    test("should only count actual numbers in numeric comparisons", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],     // Number
          ["A2", "20"],   // String
          ["A3", 30],     // Number
          ["A4", "text"], // String
          ["B1", '=COUNTIF(A1:A4, ">15")'],
        ])
      );

      expect(cell("B1")).toBe(1); // Only 30 (numbers only)
    });

    test("should handle mixed types with strict matching (no coercion)", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", 10],     // Number 10
          ["A3", "10"],   // String "10"
          ["A4", "Apple"],
          ["B1", '=COUNTIF(A1:A4, "Apple")'], // String criteria matches strings only
          ["B2", '=COUNTIF(A1:A4, "10")'],    // String criteria matches strings only
          ["B3", "=COUNTIF(A1:A4, 10)"],      // Number criteria matches numbers only
        ])
      );

      expect(cell("B1")).toBe(2); // Two "Apple" strings
      expect(cell("B2")).toBe(1); // One "10" string (no coercion, so number 10 doesn't match)
      expect(cell("B3")).toBe(1); // One 10 number (no coercion, so string "10" doesn't match)
    });
  });

  describe("error handling", () => {
    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", "=COUNTIF()");
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", '=COUNTIF(A1:A3)');
      expect(cell("A2")).toBe("#VALUE!");

      setCellContent("A3", '=COUNTIF(A1:A3, "test", "extra")');
      expect(cell("A3")).toBe("#VALUE!");
    });

    test("should handle INFINITY values in range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", "=1/0"], // Results in positive infinity
          ["A3", 20],
          ["A4", 30],
          ["B1", '=COUNTIF(A1:A4, ">15")'],
        ])
      );

      expect(cell("B1")).toBe(3); // 20, 30, and +∞ (which is > 15)
    });

    test("should handle INFINITY in criteria", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["B1", "=1/0"], // Results in "INFINITY" string
          ["C1", "=COUNTIF(A1:A2, B1)"],
        ])
      );

      expect(cell("C1")).toBe(0); // No matches for "INFINITY" string
    });
  });

  describe("edge cases", () => {
    test("should handle empty range", () => {
      setCellContent("A1", '=COUNTIF(B1:B1, "test")');
      // B1 is empty, should count as 0
      expect(cell("A1")).toBe(0);
    });

    test("should handle criteria with leading/trailing spaces", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", " Apple"],
          ["A3", "Apple "],
          ["B1", '=COUNTIF(A1:A3, "Apple")'],
        ])
      );

      expect(cell("B1")).toBe(1); // Only exact match
    });

    test("should handle zero values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 0],
          ["A2", 5],
          ["A3", 0],
          ["A4", -5],
          ["B1", "=COUNTIF(A1:A4, 0)"],
        ])
      );

      expect(cell("B1")).toBe(2);
    });

    test("should handle negative numbers", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", -10],
          ["A2", 5],
          ["A3", -20],
          ["A4", 15],
          ["B1", '=COUNTIF(A1:A4, "<0")'],
        ])
      );

      expect(cell("B1")).toBe(2); // -10, -20
    });

    test("should handle boolean values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", true],   // Boolean true
          ["A2", false],  // Boolean false  
          ["A3", true],   // Boolean true
          ["B1", '=COUNTIF(A1:A3, "TRUE")'],  // Criteria parses to boolean true
          ["B2", '=COUNTIF(A1:A3, "FALSE")'], // Criteria parses to boolean false
        ])
      );

      expect(cell("B1")).toBe(2); // Two boolean true values
      expect(cell("B2")).toBe(1); // One boolean false value
    });
  });

  describe("empty cell counting", () => {
    test("should return INFINITY for empty cells in infinite ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["B1", '=COUNTIF(A:A, "=")'], // Count empty cells in infinite column
        ])
      );

      expect(cell("B1")).toBe("INFINITY"); // Infinite empty cells in column A
    });

    test("should count empty cells in finite ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", ""], // Empty
          ["A3", "Banana"],
          ["A4", ""], // Empty
          ["A5", "Cherry"],
          ["B1", '=COUNTIF(A1:A5, "=")'], // Count empty cells in finite range
        ])
      );

      expect(cell("B1")).toBe(2); // Two empty cells (A2, A4)
    });

    test("should handle mixed empty and non-empty in finite range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Test"],
          ["A2", ""], 
          ["A3", ""], 
          ["B1", '=COUNTIF(A1:A10, "=")'], // Larger finite range
        ])
      );

      expect(cell("B1")).toBe(9); // A2, A3 are empty strings, A4-A10 are undefined (empty) = 9 empty cells
    });

    test("should work with row-based infinite ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Test"],
          ["B1", "Data"],
          ["C1", '=COUNTIF(1:1, "=")'], // Count empty cells in infinite row
        ])
      );

      expect(cell("C1")).toBe("INFINITY"); // Infinite empty cells in row 1
    });
  });

  describe("table references", () => {
    test("should work with table column references", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Fruit"],
          ["B1", "Count"],
          ["A2", "Apple"],
          ["B2", 5],
          ["A3", "Banana"],
          ["B3", 10],
          ["A4", "Apple"],
          ["B4", 3],
          ["A5", "Cherry"],
          ["B5", 7],
          ["C1", "=COUNTIF(FruitTable[Fruit], \"Apple\")"],
        ])
      );

      // Create table
      engine.addTable({
        tableName: "FruitTable",
        sheetName: sheetName,
        start: "A1",
        numRows: { type: "number", value: 4 }, // Header + 4 data rows
        numCols: 2,
      });

      expect(cell("C1")).toBe(2); // Two "Apple" entries
    });
  });
});
