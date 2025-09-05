import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("CEILING function", () => {
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
    test("should round up to nearest integer with significance 1", () => {
      setCellContent("A1", "=CEILING(4.3, 1)");
      expect(cell("A1")).toBe(5);

      setCellContent("A2", "=CEILING(4.7, 1)");
      expect(cell("A2")).toBe(5);

      setCellContent("A3", "=CEILING(4.0, 1)");
      expect(cell("A3")).toBe(4);
    });

    test("should handle negative numbers with negative significance", () => {
      setCellContent("A1", "=CEILING(-2.1, -1)");
      expect(cell("A1")).toBe(-3); // Excel result

      setCellContent("A2", "=CEILING(-2.9, -1)");
      expect(cell("A2")).toBe(-3); // Excel result

      setCellContent("A3", "=CEILING(-3.0, -1)");
      expect(cell("A3")).toBe(-3); // Excel result
    });

    test("should handle zero", () => {
      setCellContent("A1", "=CEILING(0, 1)");
      expect(cell("A1")).toBe(0);

      setCellContent("A2", "=CEILING(0, 0.5)");
      expect(cell("A2")).toBe(0);
    });

    test("should handle positive integers", () => {
      setCellContent("A1", "=CEILING(5, 1)");
      expect(cell("A1")).toBe(5);

      setCellContent("A2", "=CEILING(-5, -1)");
      expect(cell("A2")).toBe(-5);
    });
  });

  describe("custom significance", () => {
    test("should round up to nearest 0.5", () => {
      setCellContent("A1", "=CEILING(4.3, 0.5)");
      expect(cell("A1")).toBe(4.5);

      setCellContent("A2", "=CEILING(4.1, 0.5)");
      expect(cell("A2")).toBe(4.5);

      setCellContent("A3", "=CEILING(4.5, 0.5)");
      expect(cell("A3")).toBe(4.5);
    });

    test("should round up to nearest 10", () => {
      setCellContent("A1", "=CEILING(23, 10)");
      expect(cell("A1")).toBe(30);

      setCellContent("A2", "=CEILING(20, 10)");
      expect(cell("A2")).toBe(20);

      setCellContent("A3", "=CEILING(1, 10)");
      expect(cell("A3")).toBe(10);
    });

    test("should handle decimal significance", () => {
      setCellContent("A1", "=CEILING(1.234, 0.01)");
      expect(cell("A1")).toBe(1.24);

      setCellContent("A2", "=CEILING(1.236, 0.01)");
      expect(cell("A2")).toBe(1.24);
    });

    test("should handle fractional significance", () => {
      setCellContent("A1", "=CEILING(5.5, 0.25)");
      expect(cell("A1")).toBe(5.5);

      setCellContent("A2", "=CEILING(5.6, 0.25)");
      expect(cell("A2")).toBe(5.75);
    });
  });

  describe("negative significance", () => {
    test("should handle negative significance with negative numbers", () => {
      setCellContent("A1", "=CEILING(-2.1, -1)");
      expect(cell("A1")).toBe(-3); // Excel result: rounds away from zero

      setCellContent("A2", "=CEILING(-2.9, -1)");
      expect(cell("A2")).toBe(-3); // Excel result: rounds away from zero
    });

    test("should return error for mismatched signs", () => {
      setCellContent("A1", "=CEILING(2.1, -1)"); // Positive number, negative significance
      expect(cell("A1")).toBe("#NUM!"); // Excel gives #NUM! error

      setCellContent("A2", "=CEILING(3.5, -0.5)"); // Positive number, negative significance
      expect(cell("A2")).toBe("#NUM!"); // Excel gives #NUM! error

      setCellContent("A3", "=CEILING(-2.1, 1)"); // Negative number, positive significance
      expect(cell("A3")).toBe("#NUM!"); // Excel gives #NUM! error
    });
  });

  describe("error handling", () => {
    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", "=CEILING()");
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=CEILING(1)"); // Missing significance argument
      expect(cell("A2")).toBe("#VALUE!");

      setCellContent("A3", "=CEILING(1, 2, 3)");
      expect(cell("A3")).toBe("#VALUE!");
    });

    test("should return error for zero significance", () => {
      setCellContent("A1", "=CEILING(5, 0)");
      expect(cell("A1")).toBe("#DIV/0!");
    });

    test("should return error for non-number arguments", () => {
      setCellContent("A1", '=CEILING("Hello", 1)');
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=CEILING(5, TRUE)");
      expect(cell("A2")).toBe("#VALUE!");
    });

    test("should handle infinity arguments", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Positive infinity
          ["B1", "=CEILING(A1, 1)"],
          ["B2", "=CEILING(5, A1)"],
        ])
      );

      expect(cell("B1")).toBe("INFINITY"); // CEILING(+∞, 1) = +∞
      expect(cell("B2")).toBe("#VALUE!"); // Can't use infinity as significance
    });
  });

  describe("dynamic arrays", () => {
    test("should handle spilled number values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 4.3],
          ["A2", 5.7],
          ["A3", 2.1],
          ["B1", "=CEILING(A1:A3, 1)"],
        ])
      );

      expect(cell("B1")).toBe(5);  // CEILING(4.3, 1)
      expect(cell("B2")).toBe(6);  // CEILING(5.7, 1)
      expect(cell("B3")).toBe(3);  // CEILING(2.1, 1)
    });

    test("should handle spilled significance values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 0.5],
          ["A2", 0.25],
          ["A3", 0.1],
          ["B1", "=CEILING(4.3, A1:A3)"],
        ])
      );

      expect(cell("B1")).toBe(4.5);   // CEILING(4.3, 0.5)
      expect(cell("B2")).toBe(4.5);   // CEILING(4.3, 0.25) = ceil(4.3/0.25)*0.25 = ceil(17.2)*0.25 = 18*0.25 = 4.5
      expect(cell("B3")).toBe(4.3);   // CEILING(4.3, 0.1) = 4.3 (Excel result)
    });

    test("should handle multiple spilled arrays", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 4.3],
          ["A2", 5.7],
          ["B1", 0.5],
          ["B2", 0.25],
          ["C1", "=CEILING(A1:A2, B1:B2)"],
        ])
      );

      expect(cell("C1")).toBe(4.5); // CEILING(4.3, 0.5)
      expect(cell("C2")).toBe(5.75); // CEILING(5.7, 0.25) = ceil(5.7/0.25)*0.25 = ceil(22.8)*0.25 = 23*0.25 = 5.75
    });
  });

  describe("edge cases", () => {
    test("should handle very small numbers", () => {
      setCellContent("A1", "=CEILING(0.0001, 0.001)");
      expect(cell("A1")).toBe(0.001);
    });

    test("should handle large numbers", () => {
      setCellContent("A1", "=CEILING(999999.9, 1000000)");
      expect(cell("A1")).toBe(1000000);
    });

    test("should handle precision edge cases (Excel verified)", () => {
      setCellContent("A1", "=CEILING(4.3, 0.1)");
      expect(cell("A1")).toBe(4.3); // Excel result

      setCellContent("A2", "=CEILING(4.31, 0.1)");
      expect(cell("A2")).toBe(4.4); // Excel result

      setCellContent("A3", "=CEILING(1.1, 0.1)");
      expect(cell("A3")).toBe(1.1); // Excel result

      setCellContent("A4", "=CEILING(1.11, 0.1)");
      expect(cell("A4")).toBeCloseTo(1.2, 10); // Excel result (handle floating point precision)
    });

    test("should work with cell references", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 4.3],
          ["B1", 0.5],
          ["C1", "=CEILING(A1, B1)"],
        ])
      );

      expect(cell("C1")).toBe(4.5);
    });
  });
});
