import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("Concatenation Operator (&)", () => {
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
    test("should concatenate two strings", () => {
      setCellContent("A1", '="Hello"&"World"');
      expect(cell("A1")).toBe("HelloWorld");
    });

    test("should concatenate with space", () => {
      setCellContent("A1", '="Hello"&" "&"World"');
      expect(cell("A1")).toBe("Hello World");
    });

    test("should concatenate numbers", () => {
      setCellContent("A1", "=123&456");
      expect(cell("A1")).toBe("123456");
    });

    test("should concatenate mixed types", () => {
      setCellContent("A1", '="Value: "&42');
      expect(cell("A1")).toBe("Value: 42");

      setCellContent("A2", '=100&" percent"');
      expect(cell("A2")).toBe("100 percent");
    });

    test("should handle empty strings", () => {
      setCellContent("A1", '=""&"Hello"');
      expect(cell("A1")).toBe("Hello");

      setCellContent("A2", '="Hello"&""');
      expect(cell("A2")).toBe("Hello");
    });

    test("should handle zero", () => {
      setCellContent("A1", '="Count: "&0');
      expect(cell("A1")).toBe("Count: 0");
    });

    test("should handle negative numbers", () => {
      setCellContent("A1", '="Temperature: "&-15');
      expect(cell("A1")).toBe("Temperature: -15");
    });

    test("should handle decimal numbers", () => {
      setCellContent("A1", '="Pi: "&3.14159');
      expect(cell("A1")).toBe("Pi: 3.14159");
    });
  });

  describe("error handling", () => {
    test("should return error for boolean operands", () => {
      setCellContent("A1", '="Hello"&TRUE');
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=FALSE&5");
      expect(cell("A2")).toBe("#VALUE!");
    });

    test("should return error for infinity operands", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // INFINITY
          ["B1", '="Value: "&A1'],
        ])
      );

      expect(cell("B1")).toBe("#VALUE!"); // Can't concatenate INFINITY
    });
  });

  describe("cell references", () => {
    test("should work with cell references", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["B1", " "],
          ["C1", "World"],
          ["D1", "=A1&B1&C1"],
        ])
      );

      expect(cell("D1")).toBe("Hello World");
    });

    test("should work with mixed cell reference types", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Score: "],
          ["B1", 95],
          ["C1", "/"],
          ["D1", 100],
          ["E1", "=A1&B1&C1&D1"],
        ])
      );

      expect(cell("E1")).toBe("Score: 95/100");
    });
  });

  describe("dynamic arrays", () => {
    test("should handle spilled values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Good"],
          ["A3", "Nice"],
          ["B1", "=A1:A3&\" World\""],
        ])
      );

      expect(cell("B1")).toBe("Hello World");
      expect(cell("B2")).toBe("Good World");
      expect(cell("B3")).toBe("Nice World");
    });

    test("should handle multiple spilled arrays", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Good"],
          ["B1", " "],
          ["B2", " "],
          ["C1", "World"],
          ["C2", "Day"],
          ["D1", "=A1:A2&B1:B2&C1:C2"],
        ])
      );

      expect(cell("D1")).toBe("Hello World");
      expect(cell("D2")).toBe("Good Day");
    });
  });

  describe("complex expressions", () => {
    test("should work in complex formulas", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "John"],
          ["B1", 25],
          ["C1", '=IF(B1>=18, A1&" is an adult", A1&" is a minor")'],
        ])
      );

      expect(cell("C1")).toBe("John is an adult");
    });

    test("should chain multiple concatenations", () => {
      setCellContent("A1", '="A"&"B"&"C"&"D"&"E"');
      expect(cell("A1")).toBe("ABCDE");
    });
  });
});
