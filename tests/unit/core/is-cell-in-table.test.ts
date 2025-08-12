import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("FormulaEngine - isCellInTable", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });
  
  test("should return undefined when no tables exist", () => {
    const result = engine.isCellInTable({
      sheetName,
      colIndex: 0,
      rowIndex: 0,
    });

    expect(result).toBeUndefined();
  });

  test("should return undefined when cell is not in any table", () => {
    // Set up table data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Name"],
        ["B1", "Age"],
        ["C1", "City"],
        ["A2", "John"],
        ["B2", 25],
        ["C2", "NYC"],
        ["A3", "Jane"],
        ["B3", 30],
        ["C3", "LA"],
      ])
    );

    // Create table A1:C3 (3 rows including header)
    engine.addTable({
      tableName: "Table1",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 3,
    });

    // Test cell outside table
    const result = engine.isCellInTable(address("D1"));

    expect(result).toBeUndefined();
  });

  test("should return table when cell is in table header", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Name"],
        ["B1", "Age"],
        ["C1", "City"],
      ])
    );

    const table = engine.addTable({
      tableName: "Table1",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 3,
    });

    const result = engine.isCellInTable(address("B1"));

    expect(result).toBe(table);
    expect(result?.name).toBe("Table1");
  });

  test("should return table when cell is in table data", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Name"],
        ["B1", "Age"],
        ["A2", "John"],
        ["B2", 25],
        ["A3", "Jane"],
        ["B3", 30],
      ])
    );

    const table = engine.addTable({
      tableName: "Table1",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 2,
    });

    // Test data cell
    const result = engine.isCellInTable(address("B2"));

    expect(result).toBe(table);
    expect(result?.name).toBe("Table1");
  });

  test("should handle table boundaries correctly", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["B2", "Header1"],
        ["C2", "Header2"],
        ["B3", "Data1"],
        ["C3", "Data2"],
      ])
    );

    const table = engine.addTable({
      tableName: "Table1",
      sheetName,
      start: "B2", // Start at B2, not A1
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });

    // Test cells around the table
    expect(
      engine.isCellInTable(address("A2")) // A2
    ).toBeUndefined();

    expect(
      engine.isCellInTable(address("B1")) // B1
    ).toBeUndefined();

    expect(
      engine.isCellInTable(address("B2")) // B2 (start, header)
    ).toBe(table);

    expect(
      engine.isCellInTable(address("B3")) // B3 // first row
    ).toBe(table);

    expect(
      engine.isCellInTable(address("B4")) // B4 // last row
    ).toBe(table);

    expect(
      engine.isCellInTable(address("B5")) // B5 // after last row
    ).toBe(table);

    expect(
      engine.isCellInTable(address("C3")) // C3 (end)
    ).toBe(table);

    expect(
      engine.isCellInTable(address("D2")) // D2
    ).toBeUndefined();


  });

  test("should handle infinite table rows", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Header1"],
        ["B1", "Header2"],
      ])
    );

    const table = engine.addTable({
      tableName: "Table1",
      sheetName,
      start: "A1",
      numRows: { type: "infinity", sign: "positive" },
      numCols: 2,
    });

    // Test that cells far down are still in the table
    expect(
      engine.isCellInTable(address("A1")) // A1
    ).toBe(table);

    expect(
      engine.isCellInTable(address("B1001")) // B1001
    ).toBe(table);

    // But outside column range should not be in table
    expect(
      engine.isCellInTable(address("C101")) // C101
    ).toBeUndefined();
  });

  test("should return correct table when multiple tables exist", () => {
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        // Table1 data
        ["A1", "Name"],
        ["B1", "Age"],
        ["A2", "John"],
        ["B2", 25],
        // Table2 data
        ["D1", "Product"],
        ["E1", "Price"],
        ["D2", "Widget"],
        ["E2", 10],
      ])
    );

    const table1 = engine.addTable({
      tableName: "Table1",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });

    const table2 = engine.addTable({
      tableName: "Table2",
      sheetName,
      start: "D1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });

    // Test cells in first table
    expect(
      engine.isCellInTable(address("A1")) // A1
    ).toBe(table1);

    expect(
      engine.isCellInTable(address("B2")) // B2
    ).toBe(table1);

    // Test cells in second table
    expect(
      engine.isCellInTable(address("D1")) // D1
    ).toBe(table2);

    expect(
      engine.isCellInTable(address("E2")) // E2
    ).toBe(table2);

    // Test cell between tables
    expect(
      engine.isCellInTable(address("C1")) // C1
    ).toBeUndefined();
  });

  test("should handle different sheet correctly", () => {
    const sheet1Name = engine.addSheet("Sheet1").name;
    const sheet2Name = engine.addSheet("Sheet2").name;

    engine.setSheetContent(
      sheet1Name,
      new Map<string, SerializedCellValue>([
        ["A1", "Header"],
        ["A2", "Data"],
      ])
    );

    const table = engine.addTable({
      tableName: "Table1",
      sheetName: sheet1Name,
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 1,
    });

    // Cell A1 in Sheet1 should be in table
    expect(
      engine.isCellInTable({
        sheetName: sheet1Name,
        colIndex: 0,
        rowIndex: 0,
      })
    ).toBe(table);

    // Cell A1 in Sheet2 should not be in table
    expect(
      engine.isCellInTable({
        sheetName: sheet2Name,
        colIndex: 0,
        rowIndex: 0,
      })
    ).toBeUndefined();
  });
});
