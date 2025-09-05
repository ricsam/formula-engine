import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("IFERROR function", () => {
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
    test("should return value when no error", () => {
      setCellContent("A1", '=IFERROR(5, "Error")');
      expect(cell("A1")).toBe(5);

      setCellContent("A2", '=IFERROR("Hello", "Error")');
      expect(cell("A2")).toBe("Hello");

      setCellContent("A3", "=IFERROR(TRUE, FALSE)");
      expect(cell("A3")).toBe(true);
    });

    test("should return error value when first argument is actual error", () => {
      setCellContent("A1", '=IFERROR(CEILING(5), "Missing arg error")'); // Missing required arg
      expect(cell("A1")).toBe("Missing arg error");

      setCellContent("A2", '=IFERROR(CEILING(5, 0), "Zero sig error")'); // Zero significance  
      expect(cell("A2")).toBe("Zero sig error");

      setCellContent("A3", '=IFERROR("Hello">5, "Comparison error")'); // Invalid comparison
      expect(cell("A3")).toBe("Comparison error");
    });

    test("should handle different error types", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"],           // Results in "INFINITY", not an error
          ["A2", '="Hello">5'],     // #VALUE! (invalid comparison)
          ["A3", "=CEILING(5)"],    // #VALUE! (missing argument)
          ["B1", '=IFERROR(A1, "Div by zero")'],
          ["B2", '=IFERROR(A2, "Value error")'],
          ["B3", '=IFERROR(A3, "Ceiling error")'],
        ])
      );

      expect(cell("B1")).toBe("INFINITY");      // A1 is not an error, returns INFINITY
      expect(cell("B2")).toBe("Value error");   // A2 is an error
      expect(cell("B3")).toBe("Ceiling error"); // A3 is an error
    });

    test("should work with cell references", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 0],
          ["C1", "Safe result"],
          ["D1", "=IFERROR(A1/B1, C1)"],
        ])
      );

      expect(cell("D1")).toBe("INFINITY"); // 10/0 = INFINITY, not an error
    });
  });

  describe("mathematical operations", () => {
    test("should handle actual mathematical errors", () => {
      setCellContent("A1", '=IFERROR(CEILING(2.1, -1), "Sign mismatch")'); // #NUM! error
      expect(cell("A1")).toBe("Sign mismatch");
    });

    test("should handle invalid mathematical operations", () => {
      setCellContent("A1", '=IFERROR(CEILING(5), "Missing argument")');
      expect(cell("A1")).toBe("Missing argument");

      setCellContent("A2", '=IFERROR(CEILING(5, 0), "Zero significance")');
      expect(cell("A2")).toBe("Zero significance");
    });

    test("should return valid mathematical results", () => {
      setCellContent("A1", "=IFERROR(10/2, 0)");
      expect(cell("A1")).toBe(5);

      setCellContent("A2", "=IFERROR(CEILING(4.3, 1), 0)");
      expect(cell("A2")).toBe(5);
    });
  });

  describe("function integration", () => {
    test("should work with COUNTIF", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["A3", "Cherry"],
          ["B1", '=IFERROR(COUNTIF(A1:A3, "Orange"), 0)'],
        ])
      );

      expect(cell("B1")).toBe(0); // COUNTIF returns 0, not an error
    });

    test("should work with IF function", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 0],
          ["C1", '=IFERROR(IF(B1>0, A1/B1, "Zero divisor"), "Error occurred")'],
        ])
      );

      expect(cell("C1")).toBe("Zero divisor"); // IF returns string, no error
    });

    test("should handle nested IFERROR", () => {
      setCellContent("A1", '=IFERROR(IFERROR(CEILING(5), CEILING(3)), "Both failed")');
      expect(cell("A1")).toBe("Both failed"); // Both CEILING calls have missing args
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled values with actual errors", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 2],   // Valid
          ["A2", 0],   // Will cause CEILING error with zero significance
          ["A3", 1],   // Valid
          ["B1", "=IFERROR(CEILING(5, A1:A3), 99)"],
        ])
      );

      expect(cell("B1")).toBe(6);  // CEILING(5, 2) = ceil(5/2)*2 = 3*2 = 6
      expect(cell("B2")).toBe(99); // CEILING(5, 0) -> #DIV/0! -> 99
      expect(cell("B3")).toBe(5);  // CEILING(5, 1) = 5
    });

    test("should handle spilled error values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Default A"],
          ["A2", "Default B"],
          ["A3", "Default C"],
          ["B1", "=IFERROR(CEILING(5), A1:A3)"], // CEILING with missing arg -> error
        ])
      );

      expect(cell("B1")).toBe("Default A");
      expect(cell("B2")).toBe("Default B");
      expect(cell("B3")).toBe("Default C");
    });

    test("should handle multiple spilled arrays", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 3],
          ["A3", 7],
          ["B1", 1],
          ["B2", 0], // Will cause CEILING error
          ["B3", 2],
          ["C1", "Error A"],
          ["C2", "Error B"],
          ["C3", "Error C"],
          ["D1", "=IFERROR(CEILING(A1:A3, B1:B3), C1:C3)"],
        ])
      );

      expect(cell("D1")).toBe(5);        // CEILING(5, 1) = 5
      expect(cell("D2")).toBe("Error B"); // CEILING(3, 0) -> #DIV/0! -> "Error B"
      expect(cell("D3")).toBe(8);        // CEILING(7, 2) = 8
    });
  });

  describe("error handling", () => {
    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", "=IFERROR()");
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=IFERROR(5)");
      expect(cell("A2")).toBe("#VALUE!");

      setCellContent("A3", '=IFERROR(5, "error", "extra")');
      expect(cell("A3")).toBe("#VALUE!");
    });

    test("should handle valid value_if_error", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Safe value"], // Valid value, not error
          ["B1", "=IFERROR(5, A1)"], // 5 is not error, should return 5
        ])
      );

      expect(cell("B1")).toBe(5); // Returns the value (5), not the error handler
    });

    test("should handle error in error handler", () => {
      setCellContent("A1", '=IFERROR(CEILING(5), CEILING(3))'); // Both missing args
      expect(cell("A1")).toBe("#VALUE!"); // Error in error handler propagates
    });
  });

  describe("edge cases", () => {
    test("should handle empty cells", () => {
      setCellContent("B1", '=IFERROR(A1, "Empty")');
      // A1 is empty, which is not an error
      expect(cell("B1")).toBe(""); // Empty cell value, not an error
    });

    test("should handle complex expressions", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["B1", 0],
          ["C1", '=IFERROR(IF(B1=0, "Zero", A1/B1), "Calculation error")'],
        ])
      );

      expect(cell("C1")).toBe("Zero"); // IF returns "Zero", no error
    });

    test("should work with text functions", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["B1", '=IFERROR(LEN(A1), 0)'],
          ["B2", '=IFERROR(FIND("x", A1), 0)'], // "x" not found in "Hello"
        ])
      );

      expect(cell("B1")).toBe(5);  // LEN("Hello") = 5, no error
      expect(cell("B2")).toBe(0);  // FIND returns #VALUE! when not found -> 0
    });

    test("should handle return type preservation", () => {
      setCellContent("A1", "=IFERROR(42, 0)");
      expect(cell("A1")).toBe(42);

      setCellContent("A2", '=IFERROR("Text", "Error")');
      expect(cell("A2")).toBe("Text");

      setCellContent("A3", "=IFERROR(TRUE, FALSE)");
      expect(cell("A3")).toBe(true);
    });
  });
});
