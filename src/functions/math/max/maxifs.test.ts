import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("MAXIFS function", () => {
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

  test("basic single criteria", () => {
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
        ["C1", "=MAXIFS(B1:B4, A1:A4, \"Apple\")"], // Max of B where A = "Apple"
      ])
    );

    expect(cell("C1")).toBe(150); // Max of 100, 150 = 150
  });

  test("multiple criteria", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Apple"],
        ["A3", "Banana"],
        ["A4", "Apple"],
        ["B1", 5],
        ["B2", 15],
        ["B3", 25],
        ["B4", 35],
        ["C1", 100],
        ["C2", 200],
        ["C3", 150],
        ["C4", 300],
        ["D1", "=MAXIFS(C1:C4, A1:A4, \"Apple\", B1:B4, \">10\")"], // Max of C where A="Apple" AND B>10
      ])
    );

    expect(cell("D1")).toBe(300); // Only C2 (200) and C4 (300) match both criteria, max is 300
  });

  test("comparison operators", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 5],
        ["A2", 15],
        ["A3", 25],
        ["A4", 35],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["B4", 300],
        ["C1", "=MAXIFS(B1:B4, A1:A4, \">10\")"], // Max where A > 10
        ["C2", "=MAXIFS(B1:B4, A1:A4, \"<=25\")"], // Max where A <= 25
        ["C3", "=MAXIFS(B1:B4, A1:A4, \"<>15\")"], // Max where A <> 15
      ])
    );

    expect(cell("C1")).toBe(300); // Max of B2, B3, B4 = 300
    expect(cell("C2")).toBe(200); // Max of B1, B2, B3 = 200
    expect(cell("C3")).toBe(300); // Max of B1, B3, B4 = 300
  });

  test("wildcard patterns", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Apricot"],
        ["A3", "Banana"],
        ["A4", "Application"],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["B4", 300],
        ["C1", "=MAXIFS(B1:B4, A1:A4, \"App*\")"], // Max where A starts with "App"
        ["C2", "=MAXIFS(B1:B4, A1:A4, \"A?ple\")"], // Max where A matches "A?ple"
      ])
    );

    expect(cell("C1")).toBe(300); // Max of B1, B2, B4 = 300 (Apple, Apricot, Application)
    expect(cell("C2")).toBe(100); // Only B1 matches "A?ple" (Apple)
  });

  test("no matches found", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Banana"],
        ["A3", "Cherry"],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["C1", "=MAXIFS(B1:B3, A1:A3, \"Orange\")"], // No matches
      ])
    );

    expect(cell("C1")).toBe(FormulaError.VALUE);
  });

  test("mixed data types - only numbers considered", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Category"],
        ["A2", "Category"],
        ["A3", "Category"],
        ["A4", "Other"],
        ["B1", 100],
        ["B2", "text"],
        ["B3", 150],
        ["B4", 200],
        ["C1", "=MAXIFS(B1:B4, A1:A4, \"Category\")"], // Should ignore "text" in B2
      ])
    );

    expect(cell("C1")).toBe(150); // Max of 100, 150 = 150, ignoring "text"
  });

  test("single cell references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["B1", 100],
        ["C1", "=MAXIFS(B1, A1, \"Apple\")"], // Single cell max and criteria
        ["C2", "=MAXIFS(B1, A1, \"Banana\")"], // No match
      ])
    );

    expect(cell("C1")).toBe(100); // A1 matches "Apple", so max B1 = 100
    expect(cell("C2")).toBe(FormulaError.VALUE); // A1 doesn't match "Banana"
  });

  test("error handling - wrong number of arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=MAXIFS(B1:B3)"], // Too few arguments
        ["A2", "=MAXIFS(B1:B3, A1:A3)"], // Even number of arguments (missing criteria)
        ["A3", "=MAXIFS(B1:B3, A1:A3, \"Apple\", C1:C3)"], // Even number of arguments (missing last criteria)
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
    expect(cell("A3")).toBe(FormulaError.VALUE);
  });

  test("criteria as cell references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Banana"],
        ["A3", "Apple"],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["D1", "Apple"], // Criteria value
        ["D2", 10], // Numeric criteria
        ["C1", "=MAXIFS(B1:B3, A1:A3, D1)"], // Use D1 as criteria
        ["C2", "=MAXIFS(B1:B3, B1:B3, \">\" & D2)"], // Dynamic criteria construction
      ])
    );

    expect(cell("C1")).toBe(150); // Max of B1, B3 = 150
    expect(cell("C2")).toBe(200); // Max where B > 10 = 200
  });

  test("empty cells handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", ""],
        ["A2", "Apple"],
        ["A3", ""],
        ["A4", "Banana"],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["B4", 300],
        ["C1", "=MAXIFS(B1:B4, A1:A4, \"=\")"], // Max where A is empty
        ["C2", "=MAXIFS(B1:B4, A1:A4, \"<>\")"], // Max where A is not empty
      ])
    );

    expect(cell("C1")).toBe(150); // Max of B1, B3 = 150 (empty cells in A1, A3)
    expect(cell("C2")).toBe(300); // Max of B2, B4 = 300 (non-empty cells in A2, A4)
  });

  test("boolean values handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", true],
        ["A2", false],
        ["A3", true],
        ["B1", 100],
        ["B2", 200],
        ["B3", 150],
        ["C1", "=MAXIFS(B1:B3, A1:A3, TRUE)"], // Max where A is TRUE
        ["C2", "=MAXIFS(B1:B3, A1:A3, FALSE)"], // Max where A is FALSE
      ])
    );

    expect(cell("C1")).toBe(150); // Max of B1, B3 = 150
    expect(cell("C2")).toBe(200); // Only B2 = 200
  });

  test("three criteria pairs", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Apple"],
        ["A2", "Apple"],
        ["A3", "Banana"],
        ["A4", "Apple"],
        ["B1", 5],
        ["B2", 15],
        ["B3", 25],
        ["B4", 35],
        ["C1", "Red"],
        ["C2", "Blue"],
        ["C3", "Red"],
        ["C4", "Red"],
        ["D1", 100],
        ["D2", 200],
        ["D3", 150],
        ["D4", 300],
        ["E1", "=MAXIFS(D1:D4, A1:A4, \"Apple\", B1:B4, \">10\", C1:C4, \"Red\")"], // Three criteria
      ])
    );

    expect(cell("E1")).toBe(300); // Only D4 matches all three criteria: Apple, >10, Red
  });

  test("decimal precision", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["A2", 1],
        ["A3", 2],
        ["B1", 1.5],
        ["B2", 2.7],
        ["B3", 3.2],
        ["C1", "=MAXIFS(B1:B3, A1:A3, 1)"], // Max where A = 1
      ])
    );

    expect(cell("C1")).toBe(2.7); // Max of 1.5, 2.7 = 2.7
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
        ["C1", "=MAXIFS(B1:B3, A1:A3, 1000000)"], // Large number criteria
      ])
    );

    expect(cell("C1")).toBe(1500000000); // Max of B1, B3 = 1500000000
  });

  test("infinity handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=1/0"], // Positive infinity
        ["A2", "=-1/0"], // Negative infinity
        ["A3", 10],
        ["B1", 100],
        ["B2", 200],
        ["B3", 300],
        ["C1", "=MAXIFS(B1:B3, A1:A3, INFINITY)"], // Max where A is +infinity
        ["C2", "=MAXIFS(B1:B3, A1:A3, -INFINITY)"], // Max where A is -infinity
      ])
    );

    expect(cell("C1", true)).toBe(100); // Should be B1 where A1 is +infinity
    expect(cell("C2")).toBe(200); // Should be B2 where A2 is -infinity
  });

  test("nested function calls", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", 100],
        ["B2", 200],
        ["B3", 300],
        ["C1", "=MAXIFS(B1:B3, A1:A3, MAX(A1:A2))"], // Use MAX result as criteria
      ])
    );

    // MAX(A1:A2) = 20, so find max of B where A = 20
    expect(cell("C1")).toBe(200); // Only B2 matches A2 = 20
  });
});
