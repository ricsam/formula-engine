import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("COUNT function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should count numeric values only", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", "text"], // Should not be counted
          ["A3", 20],
          ["A4", true], // Should not be counted
          ["A5", 30],
          ["B1", "=COUNT(A1:A5)"],
        ])
      );

      expect(cell("B1")).toBe(3); // Only counts A1, A3, A5
    });

    test("should count direct numeric arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=COUNT(1, 2, 3)"],
          ["A2", '=COUNT(10, "text", 20, true, 30)'],
        ])
      );

      expect(cell("A1")).toBe(3); // Counts all three numbers
      expect(cell("A2")).toBe(3); // Counts 10, 20, 30; ignores "text" and true
    });

    test("should count infinity values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Positive infinity
          ["A2", "=-1/0"], // Negative infinity
          ["A3", 10],
          ["A4", "text"],
          ["B1", "=COUNT(A1:A4)"],
        ])
      );

      expect(cell("B1")).toBe(3); // Counts both infinities and the number
    });

    test("should handle empty ranges", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=COUNT(B1:B5)"], // Empty range
        ])
      );

      expect(cell("A1")).toBe(0);
    });

    test("should not count boolean values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", true],
          ["A2", false],
          ["A3", 10],
          ["B1", "=COUNT(A1:A3)"],
        ])
      );

      expect(cell("B1")).toBe(1); // Only counts the number 10
    });

    test("should not count text that looks like numbers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "123"],
          ["A2", "45.67"],
          ["A3", 89],
          ["B1", "=COUNT(A1:A3)"],
        ])
      );

      expect(cell("B1")).toBe(1); // Only counts the actual number 89
    });
  });

  describe("error handling", () => {
    test("should propagate errors immediately", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", "=CEILING()"], // This creates a #VALUE! error
          ["A3", 20],
          ["B1", "=COUNT(A1:A3)"],
        ])
      );

      expect(cell("B1")).toBe("#VALUE!"); // Should propagate the error from A2
    });

    test("should handle no arguments", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=COUNT()"],
        ])
      );

      expect(cell("A1")).toBe(0); // COUNT with no arguments returns 0
    });
  });

  describe("edge cases", () => {
    test("should handle zero values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 0],
          ["A2", -0],
          ["A3", 10],
          ["B1", "=COUNT(A1:A3)"],
        ])
      );

      expect(cell("B1")).toBe(3); // Zero is still a number
    });

    test("should handle negative numbers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", -10],
          ["A2", -20.5],
          ["A3", 15],
          ["B1", "=COUNT(A1:A3)"],
        ])
      );

      expect(cell("B1")).toBe(3); // All are numbers
    });
  });
});
