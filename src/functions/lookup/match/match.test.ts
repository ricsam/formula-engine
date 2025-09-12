import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("MATCH function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should find exact match with match_type 0", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Banana", A1:A3, 0)');

      expect(cell("B1")).toBe(2);
    });

    test("should find exact match for numbers", () => {
      // Use SerializedCellValue to set numeric values properly
      engine.setSheetContent(
        sheetAddress,
        new Map([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
        ])
      );
      setCellContent("B1", "=MATCH(20, A1:A3, 0)");

      expect(cell("B1")).toBe(2);
    });

    test("should default to match_type 1", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Banana", A1:A3)');

      expect(cell("B1")).toBe(2);
    });

    test("should return #N/A when not found", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Grape", A1:A3, 0)');

      expect(cell("B1")).toBe(FormulaError.NA);
    });

    test("should be case-sensitive", () => {
      setCellContent("A1", "apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Apple", A1:A3, 0)');

      expect(cell("B1")).toBe(FormulaError.NA);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return #VALUE! for boolean lookup_value", () => {
      setCellContent("A1", "TRUE");
      setCellContent("A2", "FALSE");
      setCellContent("B1", "=MATCH(TRUE, A1:A2, 0)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for infinity lookup_value", () => {
      setCellContent("A1", "1");
      setCellContent("A2", "2");
      setCellContent("B1", "=MATCH(INFINITY, A1:A2, 0)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for invalid match_type", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, 2)');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for string match_type", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, "exact")');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("error handling", () => {
    test("should return #VALUE! for wrong number of arguments", () => {
      setCellContent("B1", "=MATCH()");
      expect(cell("B1")).toBe(FormulaError.VALUE);

      setCellContent("B2", '=MATCH("Apple")');
      expect(cell("B2")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for too many arguments", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=MATCH("Apple", A1:A1, 0, "extra")');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should handle decimal match_type (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, 0.9)');

      expect(cell("B1")).toBe(1);
    });
  });

  describe("edge cases", () => {
    test("should handle single cell array", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=MATCH("Apple", A1, 0)');

      expect(cell("B1")).toBe(1);
    });

    test("should handle mixed string and number types (strict checking)", () => {
      // Set up mixed array with proper types
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", 10],
        ])
      );
      setCellContent("B1", "=MATCH(10, A1:A2, 0)"); // Number lookup in mixed array

      expect(cell("B1")).toBe(2);
    });
  });

  describe("can use table column as lookup_array", () => {
    test("should find exact match with match_type 0", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Fruit"],
          ["A2", "Stock"],
          ["A3", "Is local"],
          ["A2", "Apple"],
          ["B2", 2],
          ["C2", "Yes"],
          ["A2", "Orange"],
          ["B2", 3],
          ["C2", "No"],
          ["A3", "Banana"],
          ["B3", 1],
          ["C3", "Maybe"],
          ["A4", "Kiwi"],
          ["B4", 4],
          ["C4", "Yes"],
          ["A5", "Pineapple"],
          ["B5", 5],
          ["C5", "No"],
          ["A6", "Pear"],
          ["B6", 6],
          ["C6", "Yes"],
          ["A7", "Strawberry"],
          ["B7", 7],
          ["C7", "No"],
          ["A8", "Watermelon"],
          ["B8", 8],
          ["C8", "Yes"],
          ["A9", "Mango"],
          ["B9", 9],
          ["C9", "No"],
          ["A10", "Pomegranate"],

          ["K1", `=MATCH("Kiwi", A2:A10, 0)`],
          ["L1", `=MATCH("Mango", A:A, 0)`],
        ])
      );

      // expect(cell("K1")).toBe(3);
      expect(cell("L1", true)).toBe(9);
    });
  });

  describe("structured table references", () => {
    test("should work with table column references across sheets", () => {
      // Create a separate sheet for the ORDERinput table
      const inputSheetName = "InputSheet";
      engine.addSheet({ workbookName, sheetName: inputSheetName });

      // Set up data on the input sheet FIRST
      engine.setSheetContent(
        { workbookName, sheetName: inputSheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "OrderID"],
          ["B1", "Amount"],
          ["A2", "ORD-001"],
          ["B2", 100],
          ["A3", "ORD-002"],
          ["B3", 200],
          ["A4", "ORD-003"],
          ["B4", 300],
          ["A5", "ORD-004"],
          ["B5", 400],
          ["A6", "ORD-005"],
          ["B6", "=INVALID_FUNCTION()"],
        ])
      );

      // Set up data on the main sheet FIRST (only CurrentTable data)
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          // CurrentTable data - expand to 3 columns to include the formula column
          ["D7", "OrderID"],
          ["E7", "Status"],
          ["F7", "MatchResult"],
          ["D8", "ORD-002"],
          ["E8", "PENDING"],
          ["F8", `=MATCH([@[OrderID]], ORDERinput[OrderID], 0)`], // Should find ORD-002 at position 2
          ["D9", "ORD-004"],
          ["E9", "COMPLETE"],
          ["F9", `=MATCH([@[OrderID]], ORDERinput[OrderID], 0)`], // Should find ORD-004 at position 4

          // Test the cross-sheet table reference separately
          ["G1", `=ORDERinput[OrderID]`], // Test if cross-sheet structured reference works
        ])
      );

      // NOW create the tables after the sheet content is set
      // Create the ORDERinput table on the input sheet with infinite rows
      engine.addTable({
        tableName: "ORDERinput",
        sheetName: inputSheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "infinity", sign: "positive" }, // Infinite rows
        numCols: 2, // OrderID and Amount columns
      });

      // Create another table on the main sheet for the current row reference with infinite rows
      engine.addTable({
        tableName: "CurrentTable",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "D7",
        numRows: { type: "infinity", sign: "positive" }, // Infinite rows
        numCols: 3, // OrderID, Status, and MatchResult columns
      });

      expect(cell("F8", true)).toBe(2); // ORD-002 is at position 2 in ORDERinput
      expect(cell("F9", true)).toBe(4); // ORD-004 is at position 4 in ORDERinput
    });

    test("should handle empty table columns gracefully", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "OrderID"],
          ["B1", `=MATCH("TEST", EmptyTable[OrderID], 0)`],
        ])
      );

      // Create an empty table
      engine.addTable({
        tableName: "EmptyTable",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 0 }, // No data rows, just header
        numCols: 1,
      });

      expect(cell("B1")).toBe("#VALUE!"); // Should return error for empty table
    });
  });
});
