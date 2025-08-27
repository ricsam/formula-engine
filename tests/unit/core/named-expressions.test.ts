// @ts-nocheck
import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";
import { FormulaEngine } from "../../../src/core/engine";

describe("Named Expressions", () => {
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

  test("should add and use global named expressions", () => {
    // Add a global named expression
    engine.addNamedExpression({
      expressionName: "TAX_RATE",
      expression: "0.08",
    });

    // Use it in a formula
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*TAX_RATE"]
      ])
    );

    expect(cell("B1")).toBe(80); // 1000 * 0.08
  });

  test("should add and use sheet-scoped named expressions", () => {
    // Add a sheet-scoped named expression
    engine.addNamedExpression({
      expressionName: "COMMISSION",
      expression: "0.05",
      sheetName,
    });

    // Use it in a formula
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 2000],
        ["B1", "=A1*COMMISSION"]
      ])
    );

    expect(cell("B1")).toBe(100); // 2000 * 0.05
  });

  test("should update formulas when global named expression is edited", () => {
    // Add named expression and use it
    engine.addNamedExpression({ expressionName: "RATE", expression: "0.1" });
    
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 500],
        ["B1", "=A1*RATE"]
      ])
    );

    expect(cell("B1")).toBe(50); // 500 * 0.1

    // Update the named expression
    engine.updateNamedExpression({
      expressionName: "RATE",
      expression: "0.2",
    });

    // Formula should automatically update
    expect(cell("B1")).toBe(100); // 500 * 0.2
  });

  test("should update formulas when sheet-scoped named expression is edited", () => {
    // Add sheet-scoped named expression and use it
    engine.addNamedExpression({
      expressionName: "DISCOUNT",
      expression: "0.15",
      sheetName,
    });
    
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*DISCOUNT"]
      ])
    );

    expect(cell("B1")).toBe(150); // 1000 * 0.15

    // Update the named expression
    engine.updateNamedExpression({
      expressionName: "DISCOUNT",
      expression: "0.25",
      sheetName,
    });

    // Formula should automatically update
    expect(cell("B1")).toBe(250); // 1000 * 0.25
  });

  test("should show error when named expression is removed", () => {
    // Add named expression and use it
    engine.addNamedExpression({
      expressionName: "TEMP_RATE",
      expression: "0.12",
    });
    
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 800],
        ["B1", "=A1*TEMP_RATE"]
      ])
    );

    expect(cell("B1")).toBe(96); // 800 * 0.12

    // Remove the named expression
    engine.removeNamedExpression({ expressionName: "TEMP_RATE" });

    // Formula should now show error
    const result = cell("B1");
    expect(typeof result === "string" && result.startsWith("#")).toBe(true);
  });

  test("should handle global vs sheet-scoped named expression precedence", () => {
    // Add global named expression
    engine.addNamedExpression({
      expressionName: "PRIORITY",
      expression: "0.1",
    });

    // Add sheet-scoped with same name
    engine.addNamedExpression({
      expressionName: "PRIORITY",
      expression: "0.2",
      sheetName,
    });

    // Use in formula - sheet-scoped should take precedence
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 100],
        ["B1", "=A1*PRIORITY"]
      ])
    );

    expect(cell("B1")).toBe(20); // 100 * 0.2 (sheet-scoped)
  });

  test("should handle named expressions across sheets", () => {
    const sheet2 = "Sheet2";
    engine.addSheet(sheet2);

    // Add global named expression
    engine.addNamedExpression({
      expressionName: "GLOBAL_RATE",
      expression: "0.08",
    });

    // Add sheet-scoped to Sheet1
    engine.addNamedExpression({
      expressionName: "LOCAL_RATE",
      expression: "0.05",
      sheetName,
    });

    // Set up data on Sheet1
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*GLOBAL_RATE"],
        ["C1", "=A1*LOCAL_RATE"]
      ])
    );

    // Set up data on Sheet2
    engine.setSheetContent(
      sheet2,
      new Map<string, SerializedCellValue>([
        ["A1", 500],
        ["B1", "=A1*GLOBAL_RATE"],
        ["C1", "=A1*LOCAL_RATE"] // Should error since LOCAL_RATE is scoped to Sheet1
      ])
    );

    expect(cell("B1")).toBe(80); // Sheet1: 1000 * 0.08
    expect(engine.getCellValue({ sheetName: sheet2, rowIndex: 0, colIndex: 1 })).toBe(40); // Sheet2: 500 * 0.08

    // Use sheet-scoped on Sheet1 (should work)
    expect(cell("C1")).toBe(50); // 1000 * 0.05

    // Try to use sheet-scoped on Sheet2 (should error)
    const result = engine.getCellValue({ sheetName: sheet2, rowIndex: 0, colIndex: 2 });
    expect(typeof result === "string" && result.startsWith("#")).toBe(true);
  });

  test("should rename global named expressions", () => {
    // Add global named expression
    engine.addNamedExpression({
      expressionName: "OLD_RATE",
      expression: "0.15",
    });

    // Use it in a formula
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*OLD_RATE"]
      ])
    );

    expect(cell("B1")).toBe(150); // 1000 * 0.15

    // Rename the named expression
    engine.renameNamedExpression({
      expressionName: "OLD_RATE",
      newName: "NEW_RATE",
    });

    // Formula should still work with new name
    expect(cell("B1")).toBe(150);

    // Verify old name no longer exists
    const globalExpressions = engine.getGlobalNamedExpressionsSerialized();
    expect(globalExpressions.has("OLD_RATE")).toBe(false);
    expect(globalExpressions.has("NEW_RATE")).toBe(true);
    expect(globalExpressions.get("NEW_RATE").expression).toBe("0.15");
  });

  test("should rename sheet-scoped named expressions", () => {
    // Add sheet-scoped named expression
    engine.addNamedExpression({
      expressionName: "OLD_DISCOUNT",
      expression: "0.20",
      sheetName,
    });

    // Use it in a formula
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 500],
        ["B1", "=A1*OLD_DISCOUNT"]
      ])
    );

    expect(cell("B1")).toBe(100); // 500 * 0.20

    // Rename the named expression
    engine.renameNamedExpression({
      expressionName: "OLD_DISCOUNT",
      newName: "NEW_DISCOUNT",
      sheetName,
    });

    // Formula should still work with new name
    expect(cell("B1")).toBe(100);

    // Verify old name no longer exists
    const sheetExpressions = engine.getNamedExpressionsSerialized(sheetName);
    expect(sheetExpressions.has("OLD_DISCOUNT")).toBe(false);
    expect(sheetExpressions.has("NEW_DISCOUNT")).toBe(true);
    expect(sheetExpressions.get("NEW_DISCOUNT").expression).toBe("0.2");
  });

  test("should handle renaming named expressions that reference other named expressions", () => {
    // Add base named expressions
    engine.addNamedExpression({
      expressionName: "BASE_RATE",
      expression: "0.10",
    });

    engine.addNamedExpression({
      expressionName: "MULTIPLIER",
      expression: "BASE_RATE * 2",
    });

    // Use in formula
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*MULTIPLIER"]
      ])
    );

    expect(cell("B1")).toBe(200); // 1000 * (0.10 * 2)

    // Rename BASE_RATE
    engine.renameNamedExpression({
      expressionName: "BASE_RATE",
      newName: "FOUNDATION_RATE",
    });

    // Formula should still work
    expect(cell("B1")).toBe(200);

    // Verify MULTIPLIER was updated to reference new name
    const globalExpressions = engine.getGlobalNamedExpressionsSerialized();
    expect(globalExpressions.get("MULTIPLIER").expression).toBe("FOUNDATION_RATE*2");
  });

  test("should throw error when renaming non-existent named expression", () => {
    expect(() => {
      engine.renameNamedExpression({
        expressionName: "NON_EXISTENT",
        newName: "NEW_NAME",
      });
    }).toThrow("Named expression 'NON_EXISTENT' does not exist");
  });

  test("should throw error when renaming to existing name", () => {
    // Add two named expressions
    engine.addNamedExpression({
      expressionName: "RATE_A",
      expression: "0.10",
    });

    engine.addNamedExpression({
      expressionName: "RATE_B",
      expression: "0.20",
    });

    // Try to rename RATE_A to RATE_B (should fail)
    expect(() => {
      engine.renameNamedExpression({
        expressionName: "RATE_A",
        newName: "RATE_B",
      });
    }).toThrow("Named expression 'RATE_B' already exists");
  });

  test("should throw error when updating non-existent named expression", () => {
    expect(() => {
      engine.updateNamedExpression({
        expressionName: "NON_EXISTENT",
        expression: "0.15",
      });
    }).toThrow("Named expression 'NON_EXISTENT' does not exist");
  });

  test("should handle named expression serialization and bulk operations", () => {
    // Add multiple global named expressions
    engine.addNamedExpression({ expressionName: "RATE_1", expression: "0.10" });
    engine.addNamedExpression({ expressionName: "RATE_2", expression: "0.20" });

    // Add multiple sheet-scoped named expressions
    engine.addNamedExpression({ expressionName: "LOCAL_1", expression: "0.05", sheetName });
    engine.addNamedExpression({ expressionName: "LOCAL_2", expression: "0.15", sheetName });

    // Test serialization
    const globalExpressions = engine.getGlobalNamedExpressionsSerialized();
    const sheetExpressions = engine.getNamedExpressionsSerialized(sheetName);

    expect(globalExpressions.size).toBe(2);
    expect(globalExpressions.has("RATE_1")).toBe(true);
    expect(globalExpressions.has("RATE_2")).toBe(true);

    expect(sheetExpressions.size).toBe(2);
    expect(sheetExpressions.has("LOCAL_1")).toBe(true);
    expect(sheetExpressions.has("LOCAL_2")).toBe(true);

    // Test bulk setting of global named expressions
    const newGlobalExpressions = new Map([
      ["BULK_1", { name: "BULK_1", expression: "0.30" }],
      ["BULK_2", { name: "BULK_2", expression: "0.40" }],
    ]);

    engine.setGlobalNamedExpressions(newGlobalExpressions);

    const updatedGlobalExpressions = engine.getGlobalNamedExpressionsSerialized();
    expect(updatedGlobalExpressions.size).toBe(2);
    expect(updatedGlobalExpressions.has("RATE_1")).toBe(false); // Old ones should be gone
    expect(updatedGlobalExpressions.has("BULK_1")).toBe(true);
    expect(updatedGlobalExpressions.has("BULK_2")).toBe(true);

    // Test bulk setting of sheet-scoped named expressions
    const newSheetExpressions = new Map([
      ["SHEET_BULK_1", { name: "SHEET_BULK_1", expression: "0.50" }],
    ]);

    engine.setNamedExpressions(sheetName, newSheetExpressions);

    const updatedSheetExpressions = engine.getNamedExpressionsSerialized(sheetName);
    expect(updatedSheetExpressions.size).toBe(1);
    expect(updatedSheetExpressions.has("LOCAL_1")).toBe(false); // Old ones should be gone
    expect(updatedSheetExpressions.has("SHEET_BULK_1")).toBe(true);
  });

  test("should handle named expression events", () => {
    let globalExpressionsUpdatedCount = 0;
    let lastUpdatedGlobalExpressions: Map<string, any> | null = null;

    // Listen for global named expression update events
    const unsubscribe = engine.on("global-named-expressions-updated", (expressions) => {
      globalExpressionsUpdatedCount++;
      lastUpdatedGlobalExpressions = expressions;
    });

    // Add global named expression - should trigger event
    engine.addNamedExpression({
      expressionName: "EVENT_RATE",
      expression: "0.12",
    });

    expect(globalExpressionsUpdatedCount).toBe(1);
    expect(lastUpdatedGlobalExpressions.has("EVENT_RATE")).toBe(true);

    // Update global named expression - should trigger event
    engine.updateNamedExpression({
      expressionName: "EVENT_RATE",
      expression: "0.15",
    });

    expect(globalExpressionsUpdatedCount).toBe(2);

    // Rename global named expression - should trigger event
    engine.renameNamedExpression({
      expressionName: "EVENT_RATE",
      newName: "RENAMED_EVENT_RATE",
    });

    expect(globalExpressionsUpdatedCount).toBe(3);
    expect(lastUpdatedGlobalExpressions.has("RENAMED_EVENT_RATE")).toBe(true);
    expect(lastUpdatedGlobalExpressions.has("EVENT_RATE")).toBe(false);

    // Remove global named expression - should trigger event
    engine.removeNamedExpression({ expressionName: "RENAMED_EVENT_RATE" });

    expect(globalExpressionsUpdatedCount).toBe(4);
    expect(lastUpdatedGlobalExpressions.size).toBe(0);

    // Set global named expressions in bulk - should trigger event
    engine.setGlobalNamedExpressions(new Map([
      ["BULK_EVENT", { name: "BULK_EVENT", expression: "0.25" }],
    ]));

    expect(globalExpressionsUpdatedCount).toBe(5);
    expect(lastUpdatedGlobalExpressions.has("BULK_EVENT")).toBe(true);

    unsubscribe();
  });

  test("should handle complex named expression dependencies", () => {
    // Create a chain of named expressions
    engine.addNamedExpression({ expressionName: "BASE", expression: "10" });
    engine.addNamedExpression({ expressionName: "DOUBLE", expression: "BASE * 2" });
    engine.addNamedExpression({ expressionName: "TRIPLE", expression: "BASE * 3" });
    engine.addNamedExpression({ expressionName: "COMBINED", expression: "DOUBLE + TRIPLE" });

    // Use in formula
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", "=COMBINED"],
        ["B1", "=BASE + DOUBLE + TRIPLE"]
      ])
    );

    expect(cell("A1")).toBe(50); // (10 * 2) + (10 * 3) = 20 + 30 = 50
    expect(cell("B1")).toBe(60); // 10 + 20 + 30 = 60

    // Update BASE - should cascade through all dependent expressions
    engine.updateNamedExpression({ expressionName: "BASE", expression: "20" });

    expect(cell("A1")).toBe(100); // (20 * 2) + (20 * 3) = 40 + 60 = 100
    expect(cell("B1")).toBe(120); // 20 + 40 + 60 = 120

    // Rename BASE - should update all references
    engine.renameNamedExpression({ expressionName: "BASE", newName: "FOUNDATION" });

    expect(cell("A1")).toBe(100); // Should still work
    expect(cell("B1")).toBe(120); // Should still work

    // Verify all expressions were updated
    const expressions = engine.getGlobalNamedExpressionsSerialized();
    expect(expressions.get("DOUBLE").expression).toBe("FOUNDATION*2");
    expect(expressions.get("TRIPLE").expression).toBe("FOUNDATION*3");
  });

  test("should handle bulk replacement with setGlobalNamedExpressions and setNamedExpressions", () => {
    // Set up initial named expressions
    engine.addNamedExpression({ expressionName: "OLD_GLOBAL", expression: "0.10" });
    engine.addNamedExpression({ expressionName: "OLD_LOCAL", expression: "0.20", sheetName });

    // Use them in formulas
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*OLD_GLOBAL"],
        ["C1", "=A1*OLD_LOCAL"]
      ])
    );

    expect(cell("B1")).toBe(100); // 1000 * 0.10
    expect(cell("C1")).toBe(200); // 1000 * 0.20

    // Replace all global named expressions
    const newGlobalExpressions = new Map([
      ["NEW_GLOBAL_1", { name: "NEW_GLOBAL_1", expression: "0.15" }],
      ["NEW_GLOBAL_2", { name: "NEW_GLOBAL_2", expression: "0.25" }],
    ]);
    engine.setGlobalNamedExpressions(newGlobalExpressions);

    // Replace all sheet-scoped named expressions
    const newSheetExpressions = new Map([
      ["NEW_LOCAL_1", { name: "NEW_LOCAL_1", expression: "0.30" }],
      ["NEW_LOCAL_2", { name: "NEW_LOCAL_2", expression: "0.40" }],
    ]);
    engine.setNamedExpressions(sheetName, newSheetExpressions);

    // Old expressions should be gone, add new formulas with new expressions
    engine.setSheetContent(
      sheetName,
      new Map<string, SerializedCellValue>([
        ["A1", 1000],
        ["B1", "=A1*OLD_GLOBAL"], // Should error now
        ["C1", "=A1*OLD_LOCAL"],  // Should error now
        ["D1", "=A1*NEW_GLOBAL_1"],
        ["D2", "=A1*NEW_GLOBAL_2"],
        ["E1", "=A1*NEW_LOCAL_1"],
        ["E2", "=A1*NEW_LOCAL_2"]
      ])
    );

    // Old expressions should error
    const oldGlobalResult = cell("B1");
    const oldLocalResult = cell("C1");
    expect(typeof oldGlobalResult === "string" && oldGlobalResult.startsWith("#")).toBe(true);
    expect(typeof oldLocalResult === "string" && oldLocalResult.startsWith("#")).toBe(true);

    // New expressions should work
    expect(cell("D1")).toBe(150); // 1000 * 0.15
    expect(cell("D2")).toBe(250); // 1000 * 0.25
    expect(cell("E1")).toBe(300); // 1000 * 0.30
    expect(cell("E2")).toBe(400); // 1000 * 0.40

    // Verify old expressions are gone and new ones exist
    const globalExpressions = engine.getGlobalNamedExpressionsSerialized();
    const sheetExpressions = engine.getNamedExpressionsSerialized(sheetName);

    expect(globalExpressions.has("OLD_GLOBAL")).toBe(false);
    expect(globalExpressions.has("NEW_GLOBAL_1")).toBe(true);
    expect(globalExpressions.has("NEW_GLOBAL_2")).toBe(true);
    expect(globalExpressions.size).toBe(2);

    expect(sheetExpressions.has("OLD_LOCAL")).toBe(false);
    expect(sheetExpressions.has("NEW_LOCAL_1")).toBe(true);
    expect(sheetExpressions.has("NEW_LOCAL_2")).toBe(true);
    expect(sheetExpressions.size).toBe(2);
  });
});
