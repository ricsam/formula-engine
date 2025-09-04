import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("IF function", () => {
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
    test("should return true value when condition is true", () => {
      setCellContent("A1", '=IF(TRUE, "Yes", "No")');
      expect(cell("A1")).toBe("Yes");
    });

    test("should return false value when condition is false", () => {
      setCellContent("A1", '=IF(FALSE, "Yes", "No")');
      expect(cell("A1")).toBe("No");
    });

    test("should default to FALSE when no false value provided", () => {
      setCellContent("A1", '=IF(FALSE, "Yes")');
      expect(cell("A1")).toBe(false);
    });

    test("should handle numeric conditions", () => {
      setCellContent("A1", '=IF(1, "Truthy", "Falsy")');
      expect(cell("A1")).toBe("Truthy");

      setCellContent("A2", '=IF(0, "Truthy", "Falsy")');
      expect(cell("A2")).toBe("Falsy");
    });

    test("should handle string conditions", () => {
      setCellContent("A1", '=IF("Hello", "Non-empty", "Empty")');
      expect(cell("A1")).toBe("Non-empty");

      setCellContent("A2", '=IF("", "Non-empty", "Empty")');
      expect(cell("A2")).toBe("Empty");
    });

    test("should work with cell references and comparisons", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 5],
          ["C1", '=IF(A1>B1, "A1 is greater", "B1 is greater")'],
        ])
      );

      expect(cell("C1")).toBe("A1 is greater");
    });
  });

  describe("logical evaluation", () => {
    test("should treat zero as false", () => {
      setCellContent("A1", '=IF(0, "True", "False")');
      expect(cell("A1")).toBe("False");
    });

    test("should treat non-zero numbers as true", () => {
      setCellContent("A1", '=IF(42, "True", "False")');
      expect(cell("A1")).toBe("True");

      setCellContent("A2", '=IF(-1, "True", "False")');
      expect(cell("A2")).toBe("True");

      setCellContent("A3", '=IF(3.14, "True", "False")');
      expect(cell("A3")).toBe("True");
    });

    test("should treat empty string as false", () => {
      setCellContent("A1", '=IF("", "True", "False")');
      expect(cell("A1")).toBe("False");
    });

    test("should treat non-empty strings as true", () => {
      setCellContent("A1", '=IF("Hello", "True", "False")');
      expect(cell("A1")).toBe("True");

      setCellContent("A2", '=IF(" ", "True", "False")');
      expect(cell("A2")).toBe("True"); // Space is not empty
    });

    test("should handle boolean values", () => {
      setCellContent("A1", '=IF(TRUE, "True", "False")');
      expect(cell("A1")).toBe("True");

      setCellContent("A2", '=IF(FALSE, "True", "False")');
      expect(cell("A2")).toBe("False");
    });

    test("should treat INFINITY as true", () => {
      setCellContent("A1", '=IF("INFINITY", "True", "False")');
      expect(cell("A1")).toBe("True");
    });
  });

  describe("return value types", () => {
    test("should return numbers", () => {
      setCellContent("A1", "=IF(TRUE, 42, 0)");
      expect(cell("A1")).toBe(42);

      setCellContent("A2", "=IF(FALSE, 42, 0)");
      expect(cell("A2")).toBe(0);
    });

    test("should return strings", () => {
      setCellContent("A1", '=IF(TRUE, "Hello", "World")');
      expect(cell("A1")).toBe("Hello");

      setCellContent("A2", '=IF(FALSE, "Hello", "World")');
      expect(cell("A2")).toBe("World");
    });

    test("should return booleans", () => {
      setCellContent("A1", "=IF(TRUE, TRUE, FALSE)");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=IF(FALSE, TRUE, FALSE)");
      expect(cell("A2")).toBe(false);
    });

    test("should return mixed types", () => {
      setCellContent("A1", '=IF(TRUE, 42, "Text")');
      expect(cell("A1")).toBe(42);

      setCellContent("A2", '=IF(FALSE, 42, "Text")');
      expect(cell("A2")).toBe("Text");
    });
  });

  describe("nested IF statements", () => {
    test("should handle nested IF in true branch", () => {
      setCellContent("A1", '=IF(TRUE, IF(TRUE, "Inner True", "Inner False"), "Outer False")');
      expect(cell("A1")).toBe("Inner True");
    });

    test("should handle nested IF in false branch", () => {
      setCellContent("A1", '=IF(FALSE, "Outer True", IF(TRUE, "Inner True", "Inner False"))');
      expect(cell("A1")).toBe("Inner True");
    });

    test("should handle multiple levels of nesting", () => {
      setCellContent("A1", '=IF(TRUE, IF(FALSE, "Level 2 True", IF(TRUE, "Level 3 True", "Level 3 False")), "Level 1 False")');
      expect(cell("A1")).toBe("Level 3 True");
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled logical test", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 0],
          ["A3", 5],
          ["B1", '=IF(A1:A3, "True", "False")'],
        ])
      );

      expect(cell("B1")).toBe("True");  // 1 is truthy
      expect(cell("B2")).toBe("False"); // 0 is falsy
      expect(cell("B3")).toBe("True");  // 5 is truthy
    });

    test("should handle spilled true values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Option A"],
          ["A2", "Option B"],
          ["A3", "Option C"],
          ["B1", '=IF(TRUE, A1:A3, "Default")'],
        ])
      );

      expect(cell("B1")).toBe("Option A");
      expect(cell("B2")).toBe("Option B");
      expect(cell("B3")).toBe("Option C");
    });

    test("should handle spilled false values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Default A"],
          ["A2", "Default B"],
          ["A3", "Default C"],
          ["B1", '=IF(FALSE, "True Value", A1:A3)'],
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
          ["A1", 1],
          ["A2", 0],
          ["A3", 1],
          ["B1", "True A"],
          ["B2", "True B"],
          ["B3", "True C"],
          ["C1", "False A"],
          ["C2", "False B"],
          ["C3", "False C"],
          ["D1", "=IF(A1:A3, B1:B3, C1:C3)"],
        ])
      );

      expect(cell("D1")).toBe("True A");  // 1 -> True A
      expect(cell("D2")).toBe("False B"); // 0 -> False B
      expect(cell("D3")).toBe("True C");  // 1 -> True C
    });
  });

  describe("error handling", () => {
    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", "=IF()");
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=IF(TRUE)");
      expect(cell("A2")).toBe("#VALUE!");

      setCellContent("A3", '=IF(TRUE, "A", "B", "C")');
      expect(cell("A3")).toBe("#VALUE!");
    });

    test("should propagate errors from logical test", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Results in INFINITY
          ["B1", '=IF(A1, "True", "False")'],
        ])
      );

      expect(cell("B1")).toBe("True"); // INFINITY is truthy
    });

    test("should propagate errors from true value", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Results in INFINITY
          ["B1", "=IF(TRUE, A1, \"False\")"],
        ])
      );

      expect(cell("B1")).toBe("INFINITY");
    });

    test("should propagate errors from false value", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Results in INFINITY
          ["B1", "=IF(FALSE, \"True\", A1)"],
        ])
      );

      expect(cell("B1")).toBe("INFINITY");
    });

    test("should handle errors in spilled arrays", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", "=1/0"], // Error in logical test array
          ["A3", 0],
          ["B1", '=IF(A1:A3, "True", "False")'],
        ])
      );

      expect(cell("B1")).toBe("True");     // 1 is truthy
      expect(cell("B2")).toBe("True");     // INFINITY is truthy
      expect(cell("B3")).toBe("False");    // 0 is falsy
    });
  });

  describe("edge cases", () => {
    test("should handle empty cells in logical test", () => {
      setCellContent("B1", '=IF(A1, "True", "False")');
      // A1 is empty, should be falsy
      expect(cell("B1")).toBe("False");
    });

    test("should handle complex expressions with comparisons", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 5],
          ["C1", 15],
          ["D1", '=IF(A1+B1>C1, "Sum is greater", "Sum is not greater")'],
        ])
      );

      expect(cell("D1")).toBe("Sum is not greater"); // 10+5 = 15, not > 15
    });

    test("should handle string equality comparisons", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Apple"],
          ["C1", "Banana"],
          ["D1", '=IF(A1=B1, "Match", "No match")'],
          ["D2", '=IF(A1=C1, "Match", "No match")'],
        ])
      );

      expect(cell("D1")).toBe("Match");
      expect(cell("D2")).toBe("No match");
    });

    test("should work with other functions and comparisons", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "World"],
          ["B1", '=IF(LEN(A1)>LEN(A2), CONCATENATE(A1, " is longer"), CONCATENATE(A2, " is longer"))'],
        ])
      );

      expect(cell("B1")).toBe("World is longer"); // "Hello" (5) vs "World" (5), equal so not >
    });
  });
});
