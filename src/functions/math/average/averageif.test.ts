import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("AVERAGEIF function", () => {
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

  test("basic exact match criteria - 2 arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 10],
        ["A4", 30],
        ["B1", "=AVERAGEIF(A1:A4, 10)"], // Average cells equal to 10
      ])
    );

    expect(cell("B1")).toBe(10); // (10 + 10) / 2 = 10
  });

  test("basic exact match criteria - 3 arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Banana"],
        ["A3", "Apple"],
        ["A4", "Cherry"],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["B4", 300],
        ["C1", "=AVERAGEIF(A1:A4, \"Apple\", B1:B4)"], // Average B values where A = "Apple"
      ])
    );

    expect(cell("C1")).toBe(125); // (100 + 150) / 2 = 125
  });

  test("comparison operators", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 5],
        ["A2", 15],
        ["A3", 25],
        ["A4", 35],
        ["B1", "=AVERAGEIF(A1:A4, \">10\")"], // Greater than 10
        ["B2", "=AVERAGEIF(A1:A4, \"<=25\")"], // Less than or equal to 25
        ["B3", "=AVERAGEIF(A1:A4, \"<>15\")"], // Not equal to 15
      ])
    );

    expect(cell("B1")).toBe(25); // (15 + 25 + 35) / 3 = 25
    expect(cell("B2")).toBe(15); // (5 + 15 + 25) / 3 = 15
    expect(cell("B3")).toBeCloseTo(21.666666666666668); // (5 + 25 + 35) / 3 ≈ 21.67
  });

  test("wildcard patterns", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Apricot"],
        ["A3", "Banana"],
        ["A4", "Application"],
        ["B1", 10],
        ["B2", 20],
        ["B3", 30],
        ["B4", 40],
        ["C1", "=AVERAGEIF(A1:A4, \"App*\", B1:B4)"], // Starts with "App"
        ["C2", "=AVERAGEIF(A1:A4, \"A?ple\", B1:B4)"], // Matches "Apple" with ? wildcard
      ])
    );

    expect(cell("C1")).toBe(25); // (10 + 20 + 40) / 3 = 25 (Apple, Apricot, Application)
    expect(cell("C2")).toBe(10); // Only "Apple" matches "A?ple"
  });

  test("mixed data types - only numbers averaged", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Category"],
        ["A2", "Category"],
        ["A3", "Category"],
        ["A4", "Other"],
        ["B1", 10],
        ["B2", "text"],
        ["B3", 30],
        ["B4", 40],
        ["C1", "=AVERAGEIF(A1:A4, \"Category\", B1:B4)"], // Should ignore "text" in B2
      ])
    );

    expect(cell("C1")).toBe(20); // (10 + 30) / 2 = 20, ignoring "text"
  });

  test("no matches found", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", "=AVERAGEIF(A1:A3, 100)"], // No values equal 100
      ])
    );

    expect(cell("B1")).toBe(FormulaError.DIV0);
  });

  test("all matching values are non-numeric", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Apple"],
        ["A3", "Banana"],
        ["B1", "text1"],
        ["B2", "text2"],
        ["B3", 100],
        ["C1", "=AVERAGEIF(A1:A3, \"Apple\", B1:B3)"], // Matching values are text
      ])
    );

    expect(cell("C1")).toBe(FormulaError.DIV0);
  });

  test("single cell references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["B1", 100],
        ["C1", "=AVERAGEIF(A1, 10, B1)"], // Single cell criteria and average
        ["C2", "=AVERAGEIF(A1, 5, B1)"], // No match
      ])
    );

    expect(cell("C1")).toBe(100); // A1 matches 10, so average B1 = 100
    expect(cell("C2")).toBe(FormulaError.DIV0); // A1 doesn't match 5
  });

  test("error handling - wrong number of arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=AVERAGEIF(A1:A3)"], // Too few arguments
        ["A2", "=AVERAGEIF(A1:A3, 10, B1:B3, C1:C3)"], // Too many arguments
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
  });

  test("criteria as cell reference", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 10],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["D1", 10], // Criteria value
        ["C1", "=AVERAGEIF(A1:A3, D1, B1:B3)"], // Use D1 as criteria
      ])
    );

    expect(cell("C1")).toBe(125); // (100 + 150) / 2 = 125
  });

  test("empty cells handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", ""],
        ["A2", "Apple"],
        ["A3", ""],
        ["A4", "Banana"],
        ["B1", 10],
        ["B2", 20],
        ["B3", 30],
        ["B4", 40],
        ["C1", "=AVERAGEIF(A1:A4, \"=\", B1:B4)"], // Match empty cells
        ["C2", "=AVERAGEIF(A1:A4, \"<>\", B1:B4)"], // Match non-empty cells
      ])
    );

    expect(cell("C1")).toBe(20); // (10 + 30) / 2 = 20 (empty cells in A1, A3)
    expect(cell("C2")).toBe(30); // (20 + 40) / 2 = 30 (non-empty cells in A2, A4)
  });

  test("boolean values handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", true],
        ["A2", false],
        ["A3", true],
        ["B1", 10],
        ["B2", 20],
        ["B3", 30],
        ["C1", "=AVERAGEIF(A1:A3, TRUE, B1:B3)"], // Match TRUE values
        ["C2", "=AVERAGEIF(A1:A3, FALSE, B1:B3)"], // Match FALSE values
      ])
    );

    expect(cell("C1")).toBe(20); // (10 + 30) / 2 = 20
    expect(cell("C2")).toBe(20); // Only B2 = 20
  });

  test("decimal precision", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["A2", 1],
        ["A3", 2],
        ["B1", 1.5],
        ["B2", 2.5],
        ["B3", 3.5],
        ["C1", "=AVERAGEIF(A1:A3, 1, B1:B3)"], // Should be (1.5 + 2.5) / 2 = 2
      ])
    );

    expect(cell("C1")).toBe(2);
  });

  test("large numbers", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 1000000],
        ["A2", 2000000],
        ["A3", 1000000],
        ["B1", 1000000000],
        ["B2", 2000000000],
        ["B3", 1500000000],
        ["C1", "=AVERAGEIF(A1:A3, 1000000, B1:B3)"], // Large number averaging
      ])
    );

    expect(cell("C1")).toBe(1250000000); // (1000000000 + 1500000000) / 2
  });
});
