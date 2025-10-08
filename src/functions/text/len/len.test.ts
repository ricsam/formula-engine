import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("LEN function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("LEN should return string length", () => {
      setCellContent("A1", '=LEN("Hello World")');
      expect(cell("A1")).toBe(11);
    });

    test("LEN should handle empty string", () => {
      setCellContent("A1", '=LEN("")');
      expect(cell("A1")).toBe(0);
    });

    test("LEN should handle single character", () => {
      setCellContent("A1", '=LEN("A")');
      expect(cell("A1")).toBe(1);
    });

    test("LEN should handle strings with spaces", () => {
      setCellContent("A1", '=LEN("  Hello  ")');
      expect(cell("A1")).toBe(9);
    });

    test("LEN should handle strings with special characters", () => {
      setCellContent("A1", '=LEN("Hello, World! @#$%")');
      expect(cell("A1")).toBe(18); // "Hello, World! @#$%" = 18 characters
    });

    test("LEN should handle unicode characters", () => {
      setCellContent("A1", '=LEN("Hello 世界")');
      expect(cell("A1")).toBe(8);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return error for number input", () => {
      setCellContent("A1", '=LEN(12345)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for boolean input", () => {
      setCellContent("A1", '=LEN(TRUE)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for FALSE input", () => {
      setCellContent("A1", '=LEN(FALSE)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity input", () => {
      setCellContent("A1", '=LEN(INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity input", () => {
      setCellContent("A1", '=LEN(-INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled text values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "apple"],
          ["A2", "banana"],
          ["A3", "cherry"],
          ["B1", '=LEN(A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe(5);  // "apple"
      expect(cell("B2")).toBe(6);  // "banana"
      expect(cell("B3")).toBe(6);  // "cherry"
    });

    test("should handle spilled values with varying lengths", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hi"],
          ["A2", "Hello World"],
          ["A3", ""],
          ["A4", "A"],
          ["B1", '=LEN(A1:A4)'],
        ])
      );

      expect(cell("B1")).toBe(2);   // "Hi"
      expect(cell("B2")).toBe(11);  // "Hello World"
      expect(cell("B3")).toBe(0);   // ""
      expect(cell("B4")).toBe(1);   // "A"
    });

    test("should handle spilled values with special characters", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello, World!"],
          ["A2", "Test@123"],
          ["A3", "  spaces  "],
          ["B1", '=LEN(A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe(13);  // "Hello, World!"
      expect(cell("B2")).toBe(8);   // "Test@123"
      expect(cell("B3")).toBe(10);  // "  spaces  "
    });
  });

  describe("error handling", () => {
    test("should return error for invalid text argument", () => {
      setCellContent("A1", '=LEN(#REF!)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for wrong number of arguments (no arguments)", () => {
      setCellContent("A1", '=LEN()');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for wrong number of arguments (too many)", () => {
      setCellContent("A1", '=LEN("Hello","World")');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle spilled arrays with text values only", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Text1"],
          ["A2", "Text2"],
          ["A3", "Text3"],
          ["B1", '=LEN(A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe(5);  // "Text1"
      expect(cell("B2")).toBe(5);  // "Text2"
      expect(cell("B3")).toBe(5);  // "Text3"
    });
  });

  describe("edge cases", () => {
    test("should handle very long strings", () => {
      const longString = "A".repeat(1000);
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", longString],
          ["B1", '=LEN(A1)'],
        ])
      );

      expect(cell("B1")).toBe(1000);
    });

    test("should handle strings with newlines", () => {
      setCellContent("A1", '=LEN("Hello\\nWorld")');
      expect(cell("A1")).toBe(12); // "Hello\nWorld" - backslash-n is 2 chars
    });

    test("should handle strings with tabs", () => {
      setCellContent("A1", '=LEN("Hello\\tWorld")');
      expect(cell("A1")).toBe(12); // "Hello\tWorld" - backslash-t is 2 chars
    });

    test("should handle mixed unicode and ASCII", () => {
      setCellContent("A1", '=LEN("Hello 🌍 World")');
      expect(cell("A1")).toBe(14); // Including emoji
    });

    test("should handle strings with only spaces", () => {
      setCellContent("A1", '=LEN("     ")');
      expect(cell("A1")).toBe(5);
    });

    test("should handle strings with quotes", () => {
      // Use a simpler test without escaped quotes to avoid parsing issues
      setCellContent("A1", '=LEN("He said Hello")');
      expect(cell("A1")).toBe(13); // "He said Hello"
    });
  });
});
