// @ts-nocheck
import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { getCellReference, parseCellReference } from "src/core/utils";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { dependencyNodeToKey } from "src/core/utils/dependency-node-key";

describe("Tables", () => {
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

  test("should create table and use in formulas", () => {
    // Set up data first
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 10],
        ["A3", "Gadget"],
        ["B3", 15],
      ])
    );

    // Create table (3 rows total: 1 header + 2 data rows)
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 }, // 2 data rows (excluding header)
      numCols: 2,
    });

    // Now add the formula that references the table
    setCellContent("D1", "=SUM(Products[Price])");

    expect(cell("D1")).toBe(25); // 10 + 15
  });

  test("should handle table column references", () => {
    // Set up data first
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Name"],
        ["B1", "Score"],
        ["A2", "Alice"],
        ["B2", 85],
        ["A3", "Bob"],
        ["B3", 92],
      ])
    );

    // Create table (3 rows total: 1 header + 2 data rows)
    engine.addTable({
      tableName: "Scores",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 }, // 2 data rows (excluding header)
      numCols: 2,
    });

    // Add formulas after table is created - use available functions
    setCellContent("D1", "=SUM(Scores[Score])");
    setCellContent("D2", "=INDEX(Scores[Score], 1)"); // First score (85)
    setCellContent("D3", "=INDEX(Scores[Score], 2)"); // Second score (92)

    expect(cell("D1")).toBe(177); // 85 + 92
    expect(cell("D2")).toBe(85); // First score
    expect(cell("D3")).toBe(92); // Second score
  });

  test("should update formulas when table is renamed", () => {
    // Set up data first
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Item"],
        ["B1", "Value"],
        ["A2", "X"],
        ["B2", 100],
      ])
    );

    engine.addTable({
      tableName: "Data",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 }, // 1 data row (excluding header)
      numCols: 2,
    });

    // Add formula after table is created
    setCellContent("C1", "=SUM(Data[Value])");
    expect(cell("C1")).toBe(100);

    // Rename table
    engine.renameTable({ oldName: "Data", newName: "NewData" });

    // Formula should still work with new name
    expect(cell("C1")).toBe(100);
  });

  test("should show error when table is removed", () => {
    // Set up data first
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Category"],
        ["B1", "Amount"],
        ["A2", "Sales"],
        ["B2", 500],
      ])
    );

    engine.addTable({
      tableName: "TempTable",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 }, // 1 data row (excluding header)
      numCols: 2,
    });

    // Add formula after table is created
    setCellContent("C1", "=SUM(TempTable[Amount])");
    expect(cell("C1")).toBe(500);

    // Remove table
    engine.removeTable({ tableName: "TempTable" });

    // Formula should now show error
    const result = cell("C1");
    expect(typeof result === "string" && result.startsWith("#")).toBe(true);
  });

  test("should handle cross-sheet table references", () => {
    // Create table on Sheet1
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Revenue"],
        ["A2", "Alpha"],
        ["B2", 1000],
        ["D1", `=Revenue[Revenue]`],
      ])
    );

    engine.addTable({
      tableName: "Revenue",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 }, // 1 data row (excluding header)
      numCols: 2,
    });

    const sheet2 = "Sheet2";
    engine.addSheet(sheet2);

    // Reference table from Sheet2 - add formula after table is created
    engine.setSheetContent(
      sheet2,
      new Map<string, SerializedCellValue>([
        ["A1", `=Revenue[Revenue]`],
      ])
    );

    expect(cell("D1")).toBe(1000);

    expect(
      engine.getCellValue({ sheetName: sheet2, ...parseCellReference("A1") }, true)
    ).toBe(1000);
  });

  test("should update named expressions when table is renamed", () => {
    // Create table
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 150],
      ])
    );

    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    // Add global named expression that references the table
    engine.addNamedExpression({
      expressionName: "TOTAL_PRICE",
      expression: "SUM(Products[Price])*1.1",
    });

    // Add sheet-scoped named expression that references the table
    engine.addNamedExpression({
      expressionName: "DISCOUNTED_PRICE",
      expression: "SUM(Products[Price])*0.9",
      sheetName,
    });

    // Use named expressions in formulas
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 150],
        ["C1", "=TOTAL_PRICE"],
        ["C2", "=DISCOUNTED_PRICE"],
      ])
    );

    // Verify initial values
    expect(cell("C1")).toBe(165); // 150 * 1.1
    expect(cell("C2")).toBe(135); // 150 * 0.9

    // Rename the table
    engine.renameTable({
      oldName: "Products",
      newName: "Inventory",
    });

    // Verify that formulas still work after table rename
    expect(cell("C1")).toBe(165); // Should still work
    expect(cell("C2")).toBe(135); // Should still work

    // Verify that named expressions were updated
    const globalExpr = engine.getGlobalNamedExpressionsSerialized().get("TOTAL_PRICE");
    const sheetExpr = engine.getNamedExpressionsSerialized(sheetName).get("DISCOUNTED_PRICE");
    
    expect(globalExpr.expression).toBe("SUM(Inventory[Price])*1.1");
    expect(sheetExpr.expression).toBe("SUM(Inventory[Price])*0.9");
  });

  test("should update table location with updateTable", () => {
    // Set up initial data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 100],
        ["A3", "Gadget"],
        ["B3", 200],
        // New location data
        ["D1", "Item"],
        ["E1", "Cost"],
        ["D2", "Tool"],
        ["E2", 300],
        ["D3", "Part"],
        ["E3", 400],
      ])
    );

    // Create initial table
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });

    // Add formula that references the table
    setCellContent("F1", "=SUM(Products[Price])");
    expect(cell("F1")).toBe(300); // 100 + 200

    // Update table to new location
    engine.updateTable({
      tableName: "Products",
      start: "D1",
    });

    // Add new formula after table update to reference new location
    setCellContent("F2", "=SUM(Products[Cost])");
    expect(cell("F2")).toBe(700); // 300 + 400 (from new location)
    
    // Verify table structure was updated
    const tables = engine.getTablesSerialized();
    const table = tables.get("Products");
    expect(table.start.rowIndex).toBe(0); // D1 row
    expect(table.start.colIndex).toBe(3); // D1 column
  });

  test("should update table size with updateTable", () => {
    // Set up data with more rows
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 10],
        ["A3", "Gadget"],
        ["B3", 20],
        ["A4", "Tool"],
        ["B4", 30],
        ["A5", "Part"],
        ["B5", 40],
      ])
    );

    // Create table with 2 data rows initially
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });

    setCellContent("D1", "=SUM(Products[Price])");
    expect(cell("D1")).toBe(30); // 10 + 20

    // Expand table to include more rows
    engine.updateTable({
      tableName: "Products",
      numRows: { type: "number", value: 4 }, // Now 4 data rows
    });

    // Formula should now include all 4 rows
    expect(cell("D1")).toBe(100); // 10 + 20 + 30 + 40
  });

  test("should update table columns with updateTable", () => {
    // Set up data with more columns
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["C1", "Quantity"],
        ["A2", "Widget"],
        ["B2", 10],
        ["C2", 5],
        ["A3", "Gadget"],
        ["B3", 20],
        ["C3", 3],
      ])
    );

    // Create table with 2 columns initially
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });

    // Should only have Price column available
    setCellContent("D1", "=SUM(Products[Price])");
    expect(cell("D1")).toBe(30); // 10 + 20

    // Expand table to include Quantity column
    engine.updateTable({
      tableName: "Products",
      numCols: 3,
    });

    // Now should have Quantity column available
    setCellContent("D2", "=SUM(Products[Quantity])");
    expect(cell("D2")).toBe(8); // 5 + 3
  });

  test("should update multiple table properties at once", () => {
    // Set up initial data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Old"],
        ["B1", "Data"],
        ["A2", "X"],
        ["B2", 1],
        // New location with different structure
        ["D1", "New"],
        ["E1", "Values"],
        ["F1", "Extra"],
        ["D2", "Y"],
        ["E2", 10],
        ["F2", 100],
        ["D3", "Z"],
        ["E3", 20],
        ["F3", 200],
        ["D4", "W"],
        ["E4", 30],
        ["F4", 300],
      ])
    );

    // Create initial table
    engine.addTable({
      tableName: "Data",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    setCellContent("G1", "=SUM(Data[Data])");
    expect(cell("G1")).toBe(1);

    // Update table to new location, size, and columns all at once
    engine.updateTable({
      tableName: "Data",
      start: "D1",
      numRows: { type: "number", value: 3 },
      numCols: 3,
    });

    // Add new formulas after table update to reference new structure
    setCellContent("G3", "=SUM(Data[Values])");
    expect(cell("G3")).toBe(60); // 10 + 20 + 30 (Values column)

    // Should have access to the new Extra column
    setCellContent("G2", "=SUM(Data[Extra])");
    expect(cell("G2")).toBe(600); // 100 + 200 + 300
    
    // Verify table structure was updated
    const tables = engine.getTablesSerialized();
    const table = tables.get("Data");
    expect(table.start.rowIndex).toBe(0); // D1 row
    expect(table.start.colIndex).toBe(3); // D1 column
    expect(table.headers.size).toBe(3); // 3 columns now
  });

  test("should move table to different sheet", () => {
    const sheet2 = "Sheet2";
    engine.addSheet(sheet2);

    // Set up data on both sheets
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 100],
      ])
    );

    engine.setSheetContent(
      sheet2,
      new Map<string, SerializedCellValue>([
        ["A1", "Item"],
        ["B1", "Cost"],
        ["A2", "Tool"],
        ["B2", 200],
      ])
    );

    // Create table on first sheet
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    // Add formula that references the table
    setCellContent("C1", "=SUM(Products[Price])");
    expect(cell("C1")).toBe(100);

    // Move table to second sheet
    engine.updateTable({
      tableName: "Products",
      sheetName: sheet2,
    });

    // Verify table was moved to second sheet
    const tables = engine.getTablesSerialized();
    const table = tables.get("Products");
    expect(table.sheetName).toBe(sheet2);
    
    // Add new formula after table move to reference new sheet data
    setCellContent("C2", "=SUM(Products[Cost])");
    expect(cell("C2")).toBe(200);
  });

  test("should throw error when updating non-existent table", () => {
    expect(() => {
      engine.updateTable({
        tableName: "NonExistent",
        start: "A1",
      });
    }).toThrow("Table not found");
  });

  test("should handle basic table creation and updates", () => {
    // Set up data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 10],
        ["A3", "Gadget"],
        ["B3", 20],
        ["A4", "Tool"],
        ["B4", 30],
      ])
    );

    // Create table with finite rows
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 3 }, // 3 data rows
      numCols: 2,
    });

    setCellContent("D1", "=SUM(Products[Price])");
    expect(cell("D1")).toBe(60); // 10 + 20 + 30

    // Update table to reduce rows
    engine.updateTable({
      tableName: "Products",
      numRows: { type: "number", value: 2 }, // Only 2 data rows now
    });

    // Should now sum only first 2 rows
    expect(cell("D1")).toBe(30); // 10 + 20
    
    // Verify table structure was updated
    const tables = engine.getTablesSerialized();
    const table = tables.get("Products");
    expect(table.endRow.type).toBe("number");
    expect(table.endRow.type === "number" ? table.endRow.value : 0).toBe(2); // 2 data rows + header = row 3 (0-indexed: 2)
  });

  test("should preserve table properties when not specified in update", () => {
    // Set up data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["C1", "Quantity"],
        ["A2", "Widget"],
        ["B2", 10],
        ["C2", 5],
        ["A3", "Gadget"],
        ["B3", 20],
        ["C3", 3],
      ])
    );

    // Create table
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 3,
    });

    setCellContent("D1", "=SUM(Products[Price])");
    setCellContent("D2", "=SUM(Products[Quantity])");
    expect(cell("D1")).toBe(30); // 10 + 20
    expect(cell("D2")).toBe(8);  // 5 + 3

    // Update only the number of rows, preserve other properties
    engine.updateTable({
      tableName: "Products",
      numRows: { type: "number", value: 1 }, // Reduce to 1 row
    });

    // Should still have all columns but fewer rows
    expect(cell("D1")).toBe(10); // Only first row now
    expect(cell("D2")).toBe(5);  // Only first row now
  });

  test("should handle edge case: update table to same values", () => {
    // Set up data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 100],
      ])
    );

    // Create table
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    setCellContent("C1", "=SUM(Products[Price])");
    expect(cell("C1")).toBe(100);

    // Update table with same values
    engine.updateTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    // Should still work the same
    expect(cell("C1")).toBe(100);
  });

  test("should handle table serialization methods", () => {
    // Set up data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 100],
      ])
    );

    // Create table
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    // Test getTablesSerialized method
    const serializedTables = engine.getTablesSerialized();
    expect(serializedTables.size).toBe(1);
    expect(serializedTables.has("Products")).toBe(true);
    
    const table = serializedTables.get("Products");
    expect(table).toBeDefined();
    expect(table.name).toBe("Products");
    expect(table.sheetName).toBe(sheetName);
    expect(table.start.rowIndex).toBe(0);
    expect(table.start.colIndex).toBe(0);
  });

  test("should handle table events", () => {
    let tablesUpdatedCount = 0;
    let lastUpdatedTables: Map<string, any> | null = null;

    // Listen for table update events
    const unsubscribe = engine.on("tables-updated", (tables) => {
      tablesUpdatedCount++;
      lastUpdatedTables = tables;
    });

    // Set up data
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 100],
      ])
    );

    // Create table - should trigger event
    engine.addTable({
      tableName: "Products",
      sheetName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    expect(tablesUpdatedCount).toBe(1);
    expect(lastUpdatedTables.has("Products")).toBe(true);

    // Update table - should trigger event
    engine.updateTable({
      tableName: "Products",
      numRows: { type: "number", value: 1 },
    });

    expect(tablesUpdatedCount).toBe(2);

    // Rename table - should trigger event
    engine.renameTable({ oldName: "Products", newName: "Items" });

    expect(tablesUpdatedCount).toBe(3);
    expect(lastUpdatedTables.has("Items")).toBe(true);
    expect(lastUpdatedTables.has("Products")).toBe(false);

    // Remove table - should trigger event
    engine.removeTable({ tableName: "Items" });

    expect(tablesUpdatedCount).toBe(4);
    expect(lastUpdatedTables.size).toBe(0);

    unsubscribe();
  });
});
