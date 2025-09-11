import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { parseCellReference } from "src/core/utils";
import { type SerializedCellValue } from "src/core/types";

describe("Sheets", () => {
  let engine: FormulaEngine;
  const workbookName = "TestWorkbook";
  const sheetName = "TestSheet";
  const sheetAddress = { workbookName, sheetName };

  const cell = (sheetName: string, ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (
    sheetName: string,
    ref: string,
    content: SerializedCellValue
  ) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
  });

  test("should add new sheets", () => {
    // Add first sheet
    const sheet1 = engine.addSheet({ workbookName, sheetName: "Sheet1" });
    expect(sheet1.name).toBe("Sheet1");
    expect(sheet1.index).toBe(0);
    expect(engine.workbookManager.getSheets(workbookName).size).toBe(1);

    // Add second sheet
    const sheet2 = engine.addSheet({ workbookName, sheetName: "Sheet2" });
    expect(sheet2.name).toBe("Sheet2");
    expect(sheet2.index).toBe(1);
    expect(engine.workbookManager.getSheets(workbookName).size).toBe(2);

    // Verify sheets exist
    expect(engine.workbookManager.getSheets(workbookName).has("Sheet1")).toBe(
      true
    );
    expect(engine.workbookManager.getSheets(workbookName).has("Sheet2")).toBe(
      true
    );
  });

  test("should throw error when adding sheet with existing name", () => {
    engine.addSheet({ workbookName, sheetName: "DuplicateSheet" });

    expect(() => {
      engine.addSheet({ workbookName, sheetName: "DuplicateSheet" });
    }).toThrow("Sheet already exists");
  });

  test("should remove sheets", () => {
    // Add multiple sheets
    engine.addSheet({ workbookName, sheetName: "Sheet1" });
    engine.addSheet({ workbookName, sheetName: "Sheet2" });
    engine.addSheet({ workbookName, sheetName: "Sheet3" });
    expect(engine.workbookManager.getSheets(workbookName).size).toBe(3);

    // Remove middle sheet
    engine.removeSheet({ workbookName, sheetName: "Sheet2" });
    expect(engine.workbookManager.getSheets(workbookName).size).toBe(2);
    expect(engine.workbookManager.getSheets(workbookName).has("Sheet1")).toBe(
      true
    );
    expect(engine.workbookManager.getSheets(workbookName).has("Sheet2")).toBe(
      false
    );
    expect(engine.workbookManager.getSheets(workbookName).has("Sheet3")).toBe(
      true
    );
  });

  test("should throw error when removing non-existent sheet", () => {
    expect(() => {
      engine.removeSheet({ workbookName, sheetName: "NonExistent" });
    }).toThrow("Sheet not found");
  });

  test("should clean up related data when removing sheet", () => {
    const sheetName = "TestSheet";
    engine.addSheet({ workbookName, sheetName });

    // Add some data to the sheet
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "Test"],
        ["B1", 100],
      ])
    );

    // Add sheet-scoped named expression
    engine.addNamedExpression({
      expressionName: "LOCAL_RATE",
      expression: "0.10",
      sheetName,
      workbookName,
    });

    // Add table
    engine.addTable({
      tableName: "TestTable",
      sheetName,
      workbookName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    // Verify data exists
    expect(engine.getSheetExpressionsSerialized(sheetAddress).size).toBe(1);
    expect(engine.getTablesSerialized().size).toBe(1);

    // Remove sheet
    engine.removeSheet(sheetAddress);

    // Verify cleanup
    expect(engine.workbookManager.getSheets(workbookName).has(sheetName)).toBe(
      false
    );
    expect(engine.getSheetExpressionsSerialized(sheetAddress).size).toBe(0);
    // Tables should be automatically removed when sheet is removed
    expect(engine.getTablesSerialized().has("TestTable")).toBe(false);
  });

  test("should rename sheets", () => {
    const originalName = "OriginalSheet";
    const newName = "RenamedSheet";

    engine.addSheet({ workbookName, sheetName: originalName });

    // Add some content
    engine.setSheetContent(
      { workbookName, sheetName: originalName },
      new Map<string, SerializedCellValue>([
        ["A1", "Content"],
        ["B1", 42],
      ])
    );

    // Rename sheet
    engine.renameSheet({ workbookName, sheetName: originalName, newSheetName: newName });

    // Verify rename
    expect(engine.workbookManager.getSheets(workbookName).has(originalName)).toBe(false);
    expect(engine.workbookManager.getSheets(workbookName).has(newName)).toBe(true);

    // Verify content is preserved
    expect(cell(newName, "A1")).toBe("Content");
    expect(cell(newName, "B1")).toBe(42);
  });

  test("should throw error when renaming non-existent sheet", () => {
    expect(() => {
      engine.renameSheet({ workbookName, sheetName: "NonExistent", newSheetName: "NewName" });
    }).toThrow("Sheet not found");
  });

  test("should throw error when renaming to existing sheet name", () => {
    engine.addSheet({ workbookName, sheetName: "Sheet1" });
    engine.addSheet({ workbookName, sheetName: "Sheet2" });

    expect(() => {
      engine.renameSheet({ workbookName, sheetName: "Sheet1", newSheetName: "Sheet2" });
    }).toThrow("Sheet with new name already exists");
  });

  test("should update cross-sheet references when renaming sheet", () => {
    engine.addSheet({ workbookName, sheetName: "Sheet1" });
    engine.addSheet({ workbookName, sheetName: "Sheet2" });

    // Set up data on Sheet1
    engine.setSheetContent(
      { workbookName, sheetName: "Sheet1" },
      new Map([
        ["A1", 100],
        ["B1", 200],
      ])
    );

    // Set up formulas on Sheet2 that reference Sheet1
    engine.setSheetContent(
      { workbookName, sheetName: "Sheet2" },
      new Map([
        ["A1", "=Sheet1!A1"],
        ["B1", "=SUM(Sheet1!A1:B1)"],
        ["C1", "=Sheet1!A1 + Sheet1!B1"],
      ])
    );

    // Verify initial values
    expect(cell("Sheet2", "A1")).toBe(100);
    expect(cell("Sheet2", "B1")).toBe(300);
    expect(cell("Sheet2", "C1")).toBe(300);

    // Rename Sheet1 to DataSheet
    engine.renameSheet({ workbookName, sheetName: "Sheet1", newSheetName: "DataSheet" });

    // Formulas should still work with new sheet name
    expect(cell("Sheet2", "A1")).toBe(100);
    expect(cell("Sheet2", "B1")).toBe(300);
    expect(cell("Sheet2", "C1")).toBe(300);
  });

  test("should update named expressions when renaming sheet", () => {
    engine.addSheet({ workbookName, sheetName: "Sheet1" });

    // Add sheet-scoped named expression
    engine.addNamedExpression({
      expressionName: "LOCAL_VALUE",
      expression: "0.15",
      sheetName: "Sheet1",
      workbookName,
    });

    // Add global named expression that references the sheet
    engine.addNamedExpression({
      expressionName: "GLOBAL_REF",
      expression: "Sheet1!A1 * 2",
    });

    // Set up data
    engine.setSheetContent(
      { workbookName, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([
        ["A1", 50],
        ["B1", "=LOCAL_VALUE * 100"], // Uses sheet-scoped expression
      ])
    );

    expect(cell("Sheet1", "B1")).toBe(15); // 0.15 * 100

    // Rename sheet
    engine.renameSheet({ workbookName, sheetName: "Sheet1", newSheetName: "NewSheet" });

    // Sheet-scoped named expression should still work
    expect(cell("NewSheet", "B1")).toBe(15);

    // Verify sheet-scoped named expression was moved
    expect(engine.getSheetExpressionsSerialized({ workbookName, sheetName: "Sheet1" }).size).toBe(0);
    expect(engine.getSheetExpressionsSerialized({ workbookName, sheetName: "NewSheet" }).size).toBe(1);
    expect(
      engine.getSheetExpressionsSerialized({ workbookName, sheetName: "NewSheet" }).has("LOCAL_VALUE")
    ).toBe(true);

    // Note: Currently, global named expressions are not automatically updated when sheets are renamed
    // This might be a limitation that could be addressed in the future
    const globalExpressions = engine.getGlobalNamedExpressionsSerialized();
    expect(globalExpressions.get("GLOBAL_REF")!.expression).toBe(
      "Sheet1!A1 * 2"
    );
  });

  test("should handle sheet content operations", () => {
    const sheetName = "ContentSheet";
    engine.addSheet({ workbookName, sheetName });

    // Set sheet content in bulk
    const content = new Map<string, SerializedCellValue>([
      ["A1", "Header1"],
      ["B1", "Header2"],
      ["A2", 100],
      ["B2", 200],
      ["A3", "=A2+B2"],
    ]);

    engine.setSheetContent({ workbookName, sheetName }, content);

    // Verify content was set
    expect(cell(sheetName, "A1")).toBe("Header1");
    expect(cell(sheetName, "B1")).toBe("Header2");
    expect(cell(sheetName, "A2")).toBe(100);
    expect(cell(sheetName, "B2")).toBe(200);
    expect(cell(sheetName, "A3")).toBe(300); // Formula result

    // Test individual cell setting
    setCellContent(sheetName, "C1", "New Value");
    expect(cell(sheetName, "C1")).toBe("New Value");

    // Test formula setting
    setCellContent(sheetName, "C2", "=A2*2");
    expect(cell(sheetName, "C2")).toBe(200); // 100 * 2
  });

  test("should throw error when setting content on non-existent sheet", () => {
    expect(() => {
      engine.setSheetContent({ workbookName, sheetName: "NonExistent" }, new Map<string, SerializedCellValue>([["A1", "test"]]));
    }).toThrow("Sheet not found");

    expect(() => {
      setCellContent("NonExistent", "A1", "test");
    }).toThrow("Sheet not found");
  });

  test("should handle sheet re-evaluation", () => {
    engine.addSheet({ workbookName, sheetName: "Sheet1" });
    engine.addSheet({ workbookName, sheetName: "Sheet2" });

    // Set up interdependent data
    engine.setSheetContent(
      { workbookName, sheetName: "Sheet1" },
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["B1", "=A1*2"],
      ])
    );

    engine.setSheetContent({ workbookName, sheetName: "Sheet2" }, new Map<string, SerializedCellValue>([["A1", "=Sheet1!B1*3"]]));

    expect(cell("Sheet1", "B1")).toBe(20); // 10 * 2
    expect(cell("Sheet2", "A1")).toBe(60); // 20 * 3

    // Update Sheet1 data
    setCellContent("Sheet1", "A1", 20);

    // Re-evaluate specific sheet
    engine.reevaluateSheet({ workbookName, sheetName: "Sheet1" });
    expect(cell("Sheet1", "B1")).toBe(40); // 20 * 2

    // Re-evaluate all sheets
    engine.reevaluate();
    expect(cell("Sheet2", "A1")).toBe(120); // 40 * 3
  });

  test("should throw error when re-evaluating non-existent sheet", () => {
    expect(() => {
      engine.reevaluateSheet({ workbookName, sheetName: "NonExistent" });
    }).toThrow("Sheet not found");
  });

  test("should handle sheet events", () => {
    let sheetAddedCount = 0;
    let sheetRemovedCount = 0;
    let sheetRenamedCount = 0;
    let lastSheetAdded: string = '';
    let lastSheetRemoved: string = '';
    let lastSheetRenamed: { oldName: string; newName: string } = { oldName: '', newName: '' };

    // Listen for sheet events
    const unsubscribeAdded = engine.on("sheet-added", (event) => {
      sheetAddedCount++;
      lastSheetAdded = event.sheetName;
    });

    const unsubscribeRemoved = engine.on("sheet-removed", (event) => {
      sheetRemovedCount++;
      lastSheetRemoved = event.sheetName;
    });

    const unsubscribeRenamed = engine.on("sheet-renamed", (event) => {
      sheetRenamedCount++;
      lastSheetRenamed = {
        oldName: event.oldSheetName,
        newName: event.newSheetName,
      };
    });

    // Add sheet - should trigger event
    engine.addSheet({ workbookName, sheetName: "EventSheet1" });
    expect(sheetAddedCount).toBe(1);
    expect(lastSheetAdded).toBe("EventSheet1");

    // Add another sheet - should trigger event
    engine.addSheet({ workbookName, sheetName: "EventSheet2" });
    expect(sheetAddedCount).toBe(2);
    expect(lastSheetAdded).toBe("EventSheet2");

    // Rename sheet - should trigger event
    engine.renameSheet({ workbookName, sheetName: "EventSheet1", newSheetName: "RenamedSheet" });
    expect(sheetRenamedCount).toBe(1);
    expect(lastSheetRenamed.oldName).toBe("EventSheet1");
    expect(lastSheetRenamed.newName).toBe("RenamedSheet");

    // Remove sheet - should trigger event
    engine.removeSheet({ workbookName, sheetName: "RenamedSheet" });
    expect(sheetRemovedCount).toBe(1);
    expect(lastSheetRemoved).toBe("RenamedSheet");

    // Clean up
    unsubscribeAdded();
    unsubscribeRemoved();
    unsubscribeRenamed();
  });

  test("should handle complex multi-sheet scenarios", () => {
    // Create multiple sheets with interdependencies
    engine.addSheet({ workbookName, sheetName: "Data" });
    engine.addSheet({ workbookName, sheetName: "Calculations" });
    engine.addSheet({ workbookName, sheetName: "Summary" });

    // Set up base data
    engine.setSheetContent(
      { workbookName, sheetName: "Data" },
      new Map<string, SerializedCellValue>([
        ["A1", "Product"],
        ["B1", "Price"],
        ["A2", "Widget"],
        ["B2", 100],
        ["A3", "Gadget"],
        ["B3", 150],
      ])
    );

    // Set up calculations that reference data
    engine.setSheetContent(
      { workbookName, sheetName: "Calculations" },
      new Map([
        ["A1", "Tax"],
        ["B1", "Total"],
        ["A2", "=Data!B2*0.1"],
        ["B2", "=Data!B2+A2"],
        ["A3", "=Data!B3*0.1"],
        ["B3", "=Data!B3+A3"],
      ])
    );

    // Set up summary that references calculations
    engine.setSheetContent(
      { workbookName, sheetName: "Summary" },
      new Map([
        ["A1", "Grand Total"],
        ["B1", "=SUM(Calculations!B2:B3)"],
      ])
    );

    // Verify calculations
    expect(cell("Calculations", "A2")).toBe(10); // 100 * 0.1
    expect(cell("Calculations", "B2")).toBe(110); // 100 + 10
    expect(cell("Calculations", "A3")).toBe(15); // 150 * 0.1
    expect(cell("Calculations", "B3")).toBe(165); // 150 + 15
    expect(cell("Summary", "B1")).toBe(275); // 110 + 165

    // Rename Data sheet
    engine.renameSheet({ workbookName, sheetName: "Data", newSheetName: "Products" });

    // All calculations should still work
    expect(cell("Calculations", "A2")).toBe(10);
    expect(cell("Calculations", "B2")).toBe(110);
    expect(cell("Summary", "B1")).toBe(275);

    // Update data on renamed sheet
    setCellContent("Products", "B2", 200);

    // Calculations should update
    expect(cell("Calculations", "A2")).toBe(20); // 200 * 0.1
    expect(cell("Calculations", "B2")).toBe(220); // 200 + 20
    expect(cell("Summary", "B1")).toBe(385); // 220 + 165
  });

  test("should handle edge cases and error conditions", () => {
    const sheetName = "EdgeCaseSheet";
    engine.addSheet({ workbookName, sheetName });

    // Test empty sheet content
    engine.setSheetContent({ workbookName, sheetName }, new Map());
    expect(engine.workbookManager.getSheets(workbookName).get(sheetName)?.content.size).toBe(0);

    // Test setting content with formulas that reference non-existent sheets
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map([["A1", "=NonExistentSheet!A1"]])
    );

    // Should result in error
    const result = cell(sheetName, "A1");
    expect(typeof result === "string" && result.startsWith("#")).toBe(true);

    // Test renaming sheet to same name (should throw error)
    expect(() => {
      engine.renameSheet({ workbookName, sheetName, newSheetName: sheetName });
    }).toThrow("Sheet with new name already exists");
  });

  test("should preserve sheet indices when adding and removing sheets", () => {
    // Add multiple sheets
    const sheet1 = engine.addSheet({ workbookName, sheetName: "First" });
    const sheet2 = engine.addSheet({ workbookName, sheetName: "Second" });
    const sheet3 = engine.addSheet({ workbookName, sheetName: "Third" });

    expect(sheet1.index).toBe(0);
    expect(sheet2.index).toBe(1);
    expect(sheet3.index).toBe(2);

    // Remove middle sheet
    engine.removeSheet({ workbookName, sheetName: "Second" });

    // Add new sheet - should get next available index
    const sheet4 = engine.addSheet({ workbookName, sheetName: "Fourth" });
    expect(sheet4.index).toBe(2); // Should reuse the index from removed sheet
  });
});
