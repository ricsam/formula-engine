import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("LEFT function", () => {
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
    test("LEFT should return leftmost characters", () => {
      setCellContent("A1", '=LEFT("Hello World",5)');
      expect(cell("A1")).toBe("Hello");
    });

    test("LEFT should default to 1 character", () => {
      setCellContent("A1", '=LEFT("Hello")');
      expect(cell("A1")).toBe("H");
    });

    test("should handle requests longer than string", () => {
      setCellContent("A1", '=LEFT("Hi",10)');
      expect(cell("A1")).toBe("Hi");
    });

    test("should return error for negative numbers", () => {
      setCellContent("A1", '=LEFT("Hello",-1)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle empty string", () => {
      setCellContent("A1", '=LEFT("",5)');
      expect(cell("A1")).toBe("");
    });

    test("should handle zero characters", () => {
      setCellContent("A1", '=LEFT("Hello",0)');
      expect(cell("A1")).toBe("");
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return error for number input", () => {
      setCellContent("A1", '=LEFT(12345,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for boolean input", () => {
      setCellContent("A1", '=LEFT(TRUE,2)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for FALSE input", () => {
      setCellContent("A1", '=LEFT(FALSE,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for non-number numChars", () => {
      setCellContent("A1", '=LEFT("Hello",TRUE)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("strict number-only handling", () => {
    test("should return error for infinity numChars", () => {
      setCellContent("A1", '=LEFT("Hello",INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity numChars", () => {
      setCellContent("A1", '=LEFT("Hello",-INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity text input", () => {
      setCellContent("A1", '=LEFT(INFINITY,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity text input", () => {
      setCellContent("A1", '=LEFT(-INFINITY,4)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled text values with single numChars", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana,cherry"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", '=LEFT(A1:A3,3)'],
        ])
      );

      expect(cell("B1")).toBe("app");
      expect(cell("B2")).toBe("dog");
      expect(cell("B3")).toBe("red");
    });

    test("should handle single text with spilled numChars values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 3],
          ["A3", 5],
          ["B1", '=LEFT("Hello World",A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe("H");
      expect(cell("B2")).toBe("Hel");
      expect(cell("B3")).toBe("Hello");
    });

    test("should handle zipped spilled text and numChars values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana,cherry"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", 2],
          ["B2", 3],
          ["B3", 3],
          ["C1", '=LEFT(A1:A3,B1:B3)'],
        ])
      );

      expect(cell("C1")).toBe("ap");
      expect(cell("C2")).toBe("dog");
      expect(cell("C3")).toBe("red");
    });

    test("should work with FIND for comma extraction", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana,cherry"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", '=LEFT(A1:A3,FIND(",",A1:A3)-1)'],
        ])
      );

      expect(cell("B1")).toBe("apple");
      expect(cell("B2")).toBe("dog");
      expect(cell("B3")).toBe("red");
    });

    test("should handle mixed spilled values with strict number checking", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "World"],
          ["A3", "Test"],
          ["B1", 2], // Use number instead of string
          ["B2", 3],
          ["B3", 4], // Use number instead of string
          ["C1", '=LEFT(A1:A3,B1:B3)'],
        ])
      );

      expect(cell("C1")).toBe("He");
      expect(cell("C2")).toBe("Wor");
      expect(cell("C3")).toBe("Test");
    });
  });

  describe("error handling", () => {
    test("should return error for invalid text argument", () => {
      // Error propagation may transform #REF! to #ERROR! through the engine
      setCellContent("A1", '=LEFT(#REF!,3)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for invalid numChars argument", () => {
      // Error propagation may transform #REF! to #ERROR! through the engine  
      setCellContent("A1", '=LEFT("Hello",#REF!)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should handle spilled arrays with text values only", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Text"],
          ["A3", "World"],
          ["B1", '=LEFT(A1:A3,2)'],
        ])
      );

      expect(cell("B1")).toBe("He");
      expect(cell("B2")).toBe("Te");
      expect(cell("B3")).toBe("Wo");
    });

    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", '=LEFT()');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for too many arguments", () => {
      setCellContent("A1", '=LEFT("Hello",2,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("edge cases", () => {
    test("should handle very long strings", () => {
      const longString = "A".repeat(1000);
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", longString],
          ["B1", '=LEFT(A1,500)'],
        ])
      );

      expect(cell("B1")).toBe("A".repeat(500));
    });

    test("should handle decimal numChars (should floor)", () => {
      setCellContent("A1", '=LEFT("Hello",3.9)');
      expect(cell("A1")).toBe("Hel");
    });

    test("should handle very large numbers", () => {
      setCellContent("A1", '=LEFT("Hello",999999)');
      expect(cell("A1")).toBe("Hello");
    });
  });
});