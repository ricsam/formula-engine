import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("CONCATENATE function", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  describe("basic functionality", () => {
    test("CONCATENATE should join two strings", () => {
      setCellContent("A1", '=CONCATENATE("Hello", "World")');
      expect(cell("A1")).toBe("HelloWorld");
    });

    test("CONCATENATE should join multiple strings", () => {
      setCellContent("A1", '=CONCATENATE("Hello", " ", "Beautiful", " ", "World")');
      expect(cell("A1")).toBe("Hello Beautiful World");
    });

    test("CONCATENATE should handle single argument", () => {
      setCellContent("A1", '=CONCATENATE("Hello")');
      expect(cell("A1")).toBe("Hello");
    });

    test("CONCATENATE should handle empty strings", () => {
      setCellContent("A1", '=CONCATENATE("", "Hello", "")');
      expect(cell("A1")).toBe("Hello");
    });

    test("CONCATENATE should handle all empty strings", () => {
      setCellContent("A1", '=CONCATENATE("", "", "")');
      expect(cell("A1")).toBe("");
    });

    test("CONCATENATE should work with cell references", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["B1", " "],
          ["C1", "World"],
          ["D1", "=CONCATENATE(A1, B1, C1)"],
        ])
      );
      expect(cell("D1")).toBe("Hello World");
    });
  });

  describe("type coercion", () => {
    test("should convert numbers to strings", () => {
      setCellContent("A1", '=CONCATENATE("Hello", 123)');
      expect(cell("A1")).toBe("Hello123");
    });

    test("should convert booleans to strings", () => {
      setCellContent("A1", '=CONCATENATE("Hello", TRUE)');
      expect(cell("A1")).toBe("HelloTRUE");
      
      setCellContent("A2", '=CONCATENATE("Hello", FALSE)');
      expect(cell("A2")).toBe("HelloFALSE");
    });

    test("should handle INFINITY literal", () => {
      setCellContent("A1", '=CONCATENATE("Hello", "INFINITY")');
      expect(cell("A1")).toBe("HelloINFINITY");
    });

    test("should handle mixed types", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["B1", 123],
          ["C1", "=CONCATENATE(A1, B1)"],
        ])
      );
      expect(cell("C1")).toBe("Hello123");
    });

    test("should handle multiple mixed types", () => {
      setCellContent("A1", '=CONCATENATE("Value: ", 42, ", Status: ", TRUE, ", Result: ", 3.14)');
      expect(cell("A1")).toBe("Value: 42, Status: TRUE, Result: 3.14");
    });

    test("should handle negative numbers", () => {
      setCellContent("A1", '=CONCATENATE("Temperature: ", -15, " degrees")');
      expect(cell("A1")).toBe("Temperature: -15 degrees");
    });

    test("should handle zero", () => {
      setCellContent("A1", '=CONCATENATE("Count: ", 0)');
      expect(cell("A1")).toBe("Count: 0");
    });

    test("should handle decimal numbers", () => {
      setCellContent("A1", '=CONCATENATE("Pi is approximately ", 3.14159)');
      expect(cell("A1")).toBe("Pi is approximately 3.14159");
    });
  });

  describe("dynamic arrays", () => {
    test("should handle spilled text values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Good"],
          ["A3", "Nice"],
          ["B1", " "],
          ["B2", " "],
          ["B3", " "],
          ["C1", "World"],
          ["C2", "Day"],
          ["C3", "Work"],
          ["D1", "=CONCATENATE(A1:A3, B1:B3, C1:C3)"],
        ])
      );

      expect(cell("D1")).toBe("Hello World");
      expect(cell("D2")).toBe("Good Day");
      expect(cell("D3")).toBe("Nice Work");
    });

    test("should handle mixed spilled and single values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Good"],
          ["A3", "Nice"],
          ["B1", '=CONCATENATE(A1:A3, " World")'],
        ])
      );

      expect(cell("B1")).toBe("Hello World");
      expect(cell("B2")).toBe("Good World");
      expect(cell("B3")).toBe("Nice World");
    });

    test("should handle single value with spilled values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", " World"],
          ["A2", " Day"],
          ["A3", " Work"],
          ["B1", '=CONCATENATE("Hello", A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe("Hello World");
      expect(cell("B2")).toBe("Hello Day");
      expect(cell("B3")).toBe("Hello Work");
    });

    test("should handle multiple spilled arrays", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Good"],
          ["B1", " "],
          ["B2", " "],
          ["C1", "World"],
          ["C2", "Day"],
          ["D1", "!"],
          ["D2", "!"],
          ["E1", "=CONCATENATE(A1:A2, B1:B2, C1:C2, D1:D2)"],
        ])
      );

      expect(cell("E1")).toBe("Hello World!");
      expect(cell("E2")).toBe("Good Day!");
    });

    test("should handle mixed types in spilled arrays", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Value: "],
          ["A2", "Count: "],
          ["B1", 42],
          ["B2", 100],
          ["C1", "TRUE"],
          ["C2", "FALSE"],
          ["D1", "=CONCATENATE(A1:A2, B1:B2, \", Active: \", C1:C2)"],
        ])
      );

      expect(cell("D1")).toBe("Value: 42, Active: TRUE");
      expect(cell("D2")).toBe("Count: 100, Active: FALSE");
    });
  });

  describe("error handling", () => {
    test("should return error for no arguments", () => {
      setCellContent("A1", "=CONCATENATE()");
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle INFINITY string values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "INFINITY"], // String literal
          ["B1", '=CONCATENATE(A1, "Hello")'],
        ])
      );
      expect(cell("B1")).toBe("INFINITYHello");
    });

    test("should handle INFINITY in spilled values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "INFINITY"], // String literal
          ["B1", '=CONCATENATE(A1:A2, " World")'],
        ])
      );
      expect(cell("B1")).toBe("Hello World");
      expect(cell("B2")).toBe("INFINITY World");
    });
  });

  describe("edge cases", () => {
    test("should handle very long concatenation", () => {
      const longString = "A".repeat(1000);
      setCellContent("A1", `=CONCATENATE("${longString}", "${longString}")`);
      expect(cell("A1")).toBe(longString + longString);
    });

    test("should handle special characters", () => {
      setCellContent("A1", '=CONCATENATE("Hello\\n", "World\\t", "!")');
      expect(cell("A1")).toBe("Hello\\nWorld\\t!");
    });

    test("should handle unicode characters", () => {
      setCellContent("A1", '=CONCATENATE("Hello", " 🌍", " 世界")');
      expect(cell("A1")).toBe("Hello 🌍 世界");
    });

    test("should handle quotes in strings", () => {
      setCellContent("A1", "=CONCATENATE(\"He said 'Hello'\", \" to me\")");
      expect(cell("A1")).toBe("He said 'Hello' to me");
    });
  });
});
