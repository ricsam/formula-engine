import { describe, expect, test, beforeEach } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { parseCellReference } from "src/core/utils";
import type { SerializedCellValue } from "src/core/types";

describe("AVERAGEIFS function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  test("should return error for insufficient arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=AVERAGEIFS(A1:A5, B1:B5)"], // Missing criteria
      ])
    );

    expect(cell("A1")).toBe("#VALUE!");
  });

  test("should return error for even number of arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=AVERAGEIFS(A1:A5, B1:B5, 1, C1:C5)"], // Missing second criteria
      ])
    );

    expect(cell("A1")).toBe("#VALUE!");
  });

  test("should handle basic single criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", "Apple"],
        ["B2", "Banana"],
        ["B3", "Apple"],
        ["C1", '=AVERAGEIFS(A1:A3, B1:B3, "Apple")'],
      ])
    );

    expect(cell("C1")).toBe(20); // (10 + 30) / 2
  });

  test("should handle multiple criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10], // values to average
        ["A2", 20],
        ["A3", 30],
        ["A4", 40],
        ["B1", "Apple"], // first criteria range
        ["B2", "Banana"],
        ["B3", "Apple"],
        ["B4", "Apple"],
        ["C1", 5], // second criteria range
        ["C2", 15],
        ["C3", 25],
        ["C4", 35],
        ["D1", '=AVERAGEIFS(A1:A4, B1:B4, "Apple", C1:C4, ">10")'],
      ])
    );

    expect(cell("D1")).toBe(35); // Only A3 and A4 match both criteria: (30 + 40) / 2
  });

  test("should handle numeric criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 100],
        ["A2", 200],
        ["A3", 300],
        ["B1", 50],
        ["B2", 150],
        ["B3", 250],
        ["C1", '=AVERAGEIFS(A1:A3, B1:B3, ">100")'],
      ])
    );

    expect(cell("C1")).toBe(250); // Only A2 and A3 match: (200 + 300) / 2
  });

  test("should handle comparison operators", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["A4", 40],
        ["B1", 5],
        ["B2", 15],
        ["B3", 25],
        ["B4", 35],
        ["C1", '=AVERAGEIFS(A1:A4, B1:B4, ">=20")'], // Test >=
        ["C2", '=AVERAGEIFS(A1:A4, B1:B4, "<=15")'], // Test <=
        ["C3", '=AVERAGEIFS(A1:A4, B1:B4, "<>15")'], // Test <>
      ])
    );

    expect(cell("C1")).toBe(35); // A3 and A4: (30 + 40) / 2
    expect(cell("C2")).toBe(15); // A1 and A2: (10 + 20) / 2
    expect(cell("C3")).toBeCloseTo(26.666666666666668); // A1, A3, A4: (10 + 30 + 40) / 3
  });

  test("should return error when no values match criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["B1", "Apple"],
        ["B2", "Banana"],
        ["C1", '=AVERAGEIFS(A1:A2, B1:B2, "Orange")'],
      ])
    );

    expect(cell("C1")).toBe("#DIV/0!");
  });

  test("should handle wildcard criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 100],
        ["A2", 200],
        ["A3", 300],
        ["B1", "Apple"],
        ["B2", "Apricot"],
        ["B3", "Banana"],
        ["C1", '=AVERAGEIFS(A1:A3, B1:B3, "Ap*")'],
      ])
    );

    expect(cell("C1")).toBe(150); // A1 and A2: (100 + 200) / 2
  });

  test("should handle three criteria pairs", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10], // values to average
        ["A2", 20],
        ["A3", 30],
        ["A4", 40],
        ["B1", "Apple"], // first criteria
        ["B2", "Apple"],
        ["B3", "Apple"],
        ["B4", "Banana"],
        ["C1", 5], // second criteria
        ["C2", 15],
        ["C3", 25],
        ["C4", 35],
        ["D1", "Red"], // third criteria
        ["D2", "Blue"],
        ["D3", "Red"],
        ["D4", "Red"],
        ["E1", '=AVERAGEIFS(A1:A4, B1:B4, "Apple", C1:C4, ">10", D1:D4, "Red")'],
      ])
    );

    expect(cell("E1")).toBe(30); // Only A3 matches all criteria: 30 / 1
  });

  test("should ignore non-numeric values in average range", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", "text"], // non-numeric
        ["A3", 30],
        ["B1", "Yes"],
        ["B2", "Yes"],
        ["B3", "Yes"],
        ["C1", '=AVERAGEIFS(A1:A3, B1:B3, "Yes")'],
      ])
    );

    expect(cell("C1")).toBe(20); // Only A1 and A3 are numeric: (10 + 30) / 2
  });

  test("should handle empty cells in criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", "Apple"],
        // B2 is empty
        ["B3", "Banana"],
        ["C1", '=AVERAGEIFS(A1:A3, B1:B3, "")'],
      ])
    );

    expect(cell("C1")).toBe(20); // Only A2 matches empty criteria: 20 / 1
  });

  test("should handle single value ranges", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 100],
        ["B1", "Test"],
        ["C1", '=AVERAGEIFS(A1, B1, "Test")'],
      ])
    );

    expect(cell("C1")).toBe(100); // Single matching value
  });

  test("should handle mixed numeric and text criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 50],
        ["A2", 75],
        ["A3", 100],
        ["A4", 125],
        ["B1", "Product A"],
        ["B2", "Product B"],
        ["B3", "Product A"],
        ["B4", "Product A"],
        ["C1", 80],
        ["C2", 90],
        ["C3", 110],
        ["C4", 120],
        ["D1", '=AVERAGEIFS(A1:A4, B1:B4, "Product A", C1:C4, ">=100")'],
      ])
    );

    expect(cell("D1")).toBe(112.5); // Only A3 and A4 match: (100 + 125) / 2
  });
});