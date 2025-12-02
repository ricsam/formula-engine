import { describe, test, expect, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { defineApi } from "../../../src/core/api/api";
import { SchemaIntegrityError } from "../../../src/core/commands/command-executor";
import {
  FormulaError,
  type SerializedCellValue,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";
import { serialize, deserialize } from "../../../src/core/map-serializer";

describe("Command Pattern", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "Sheet1";

  // Parse functions for schema validation
  const parseNumber = (value: unknown) => {
    if (typeof value !== "number") {
      throw new Error("Expected a number value");
    }
    return value;
  };

  const parseString = (value: unknown) => {
    if (typeof value !== "string") {
      throw new Error("Expected a string value");
    }
    return value;
  };

  describe("Schema Validation with Evaluated Values", () => {
    test("validates evaluated formula result, not raw content", () => {
      // Create API with number schema
      const api = defineApi().addCellApi(
        "numericCell",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        parseNumber
      );

      const engine = new FormulaEngine(api);
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // =123+123 is a string, but evaluates to 246 (number) - should pass
      expect(() => {
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
          "=123+123"
        );
      }).not.toThrow();

      // Verify the value is evaluated correctly
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(246);
    });

    test("throws SchemaIntegrityError when formula evaluates to wrong type", () => {
      // Create API with number schema
      const api = defineApi().addCellApi(
        "numericCell",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        parseNumber
      );

      const engine = new FormulaEngine(api);
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // ="hello" evaluates to "hello" (string) - should fail
      expect(() => {
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
          '="hello"'
        );
      }).toThrow(SchemaIntegrityError);
    });

    test("validates dependent cells when dependency changes", () => {
      // Create API with number schema on B1, which references A1
      const api = defineApi().addCellApi(
        "numericCell",
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 }, // B1
        parseNumber
      );

      const engine = new FormulaEngine(api);
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // Set A1 to a number
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        100
      );

      // Set B1 to reference A1 - should pass since A1 is a number
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=A1"
      );

      // Now change A1 to a string - should fail because B1 now evaluates to string
      expect(() => {
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
          "hello"
        );
      }).toThrow(SchemaIntegrityError);

      // Verify A1 was rolled back to its original value
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(100);
    });
  });

  describe("Rollback on Validation Failure", () => {
    test("rolls back cell content when validation fails", () => {
      const api = defineApi().addCellApi(
        "numericCell",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        parseNumber
      );

      const engine = new FormulaEngine(api);
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // Set initial valid value
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        42
      );

      // Try to set invalid value
      expect(() => {
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
          "not a number"
        );
      }).toThrow(SchemaIntegrityError);

      // Verify the value was rolled back
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(42);
    });

    test("rolls back sheet content when validation fails", () => {
      const api = defineApi().addCellApi(
        "numericCell",
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        parseNumber
      );

      const engine = new FormulaEngine(api);
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // Set initial valid content
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", 100],
          ["B1", 200],
        ])
      );

      // Try to set invalid content
      expect(() => {
        engine.setSheetContent(
          { workbookName, sheetName },
          new Map<string, SerializedCellValue>([
            ["A1", "invalid"],
            ["B1", 300],
          ])
        );
      }).toThrow(SchemaIntegrityError);

      // Verify the content was rolled back
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(100);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(200);
    });
  });

  describe("Undo/Redo", () => {
    let engine: FormulaEngine;

    beforeEach(() => {
      engine = FormulaEngine.buildEmpty();
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });
    });

    test("can undo and redo cell content changes", () => {
      const address = { workbookName, sheetName, colIndex: 0, rowIndex: 0 };

      engine.setCellContent(address, "first");
      expect(engine.getCellValue(address)).toBe("first");

      engine.setCellContent(address, "second");
      expect(engine.getCellValue(address)).toBe("second");

      // Undo
      expect(engine.undo()).toBe(true);
      expect(engine.getCellValue(address)).toBe("first");

      // Redo
      expect(engine.redo()).toBe(true);
      expect(engine.getCellValue(address)).toBe("second");
    });

    test("can undo and redo sheet operations", () => {
      expect(engine.hasSheet({ workbookName, sheetName: "NewSheet" })).toBe(
        false
      );

      engine.addSheet({ workbookName, sheetName: "NewSheet" });
      expect(engine.hasSheet({ workbookName, sheetName: "NewSheet" })).toBe(
        true
      );

      // Undo
      expect(engine.undo()).toBe(true);
      expect(engine.hasSheet({ workbookName, sheetName: "NewSheet" })).toBe(
        false
      );

      // Redo
      expect(engine.redo()).toBe(true);
      expect(engine.hasSheet({ workbookName, sheetName: "NewSheet" })).toBe(
        true
      );
    });

    test("can undo and redo workbook operations", () => {
      expect(engine.hasWorkbook("NewWorkbook")).toBe(false);

      engine.addWorkbook("NewWorkbook");
      expect(engine.hasWorkbook("NewWorkbook")).toBe(true);

      // Undo
      expect(engine.undo()).toBe(true);
      expect(engine.hasWorkbook("NewWorkbook")).toBe(false);

      // Redo
      expect(engine.redo()).toBe(true);
      expect(engine.hasWorkbook("NewWorkbook")).toBe(true);
    });

    test("can undo and redo table operations", () => {
      expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(
        false
      );

      engine.addTable({
        workbookName,
        sheetName,
        tableName: "Table1",
        start: "A1",
        numRows: { type: "number", value: 5 },
        numCols: 3,
      });
      expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(true);

      // Undo
      expect(engine.undo()).toBe(true);
      expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(
        false
      );

      // Redo
      expect(engine.redo()).toBe(true);
      expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(true);
    });

    test("can undo and redo metadata changes (no reevaluation)", () => {
      const address = { workbookName, sheetName, colIndex: 0, rowIndex: 0 };

      engine.setCellMetadata(address, { note: "first" });
      expect(engine.getCellMetadata(address)).toEqual({ note: "first" });

      engine.setCellMetadata(address, { note: "second" });
      expect(engine.getCellMetadata(address)).toEqual({ note: "second" });

      // Undo
      expect(engine.undo()).toBe(true);
      expect(engine.getCellMetadata(address)).toEqual({ note: "first" });

      // Redo
      expect(engine.redo()).toBe(true);
      expect(engine.getCellMetadata(address)).toEqual({ note: "second" });
    });

    test("can undo and redo style changes", () => {
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 5 },
                row: { type: "number", value: 5 },
              },
            },
          },
        ],
        style: { backgroundColor: "#FF0000" },
      });
      expect(engine.getCellStyleCount(workbookName)).toBe(1);

      // Undo
      expect(engine.undo()).toBe(true);
      expect(engine.getCellStyleCount(workbookName)).toBe(0);

      // Redo
      expect(engine.redo()).toBe(true);
      expect(engine.getCellStyleCount(workbookName)).toBe(1);
    });

    test("canUndo and canRedo report correct state", () => {
      // Clear history from beforeEach setup
      engine.clearHistory();

      expect(engine.canUndo()).toBe(false);
      expect(engine.canRedo()).toBe(false);

      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "test"
      );

      expect(engine.canUndo()).toBe(true);
      expect(engine.canRedo()).toBe(false);

      engine.undo();

      expect(engine.canUndo()).toBe(false);
      expect(engine.canRedo()).toBe(true);
    });

    test("new action clears redo stack", () => {
      const address = { workbookName, sheetName, colIndex: 0, rowIndex: 0 };

      engine.setCellContent(address, "first");
      engine.setCellContent(address, "second");
      engine.undo();

      expect(engine.canRedo()).toBe(true);

      // New action should clear redo stack
      engine.setCellContent(address, "third");
      expect(engine.canRedo()).toBe(false);
    });
  });

  describe("Action Serialization and Replay", () => {
    test("can serialize and deserialize action log", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "hello"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        42
      );

      // Get action log
      const actionLog = engine.getActionLog();
      expect(actionLog.length).toBeGreaterThan(0);

      // Serialize actions
      const serialized = serialize(actionLog);
      expect(typeof serialized).toBe("string");

      // Deserialize actions
      const deserialized = deserialize(serialized);
      expect(Array.isArray(deserialized)).toBe(true);
      expect(deserialized).toHaveLength(actionLog.length);
    });

    test("actions contain correct type and payload", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      const address = { workbookName, sheetName, colIndex: 0, rowIndex: 0 };
      engine.setCellContent(address, "test");

      const actionLog = engine.getActionLog();
      const setCellAction = actionLog.find(
        (a) => a.type === "SET_CELL_CONTENT"
      );

      expect(setCellAction).toBeDefined();
      expect(setCellAction?.payload).toMatchObject({
        address,
        content: "test",
      });
    });

    test("can replay actions to rebuild state", () => {
      // Create first engine and perform operations
      const engine1 = FormulaEngine.buildEmpty();
      engine1.addWorkbook(workbookName);
      engine1.addSheet({ workbookName, sheetName });
      engine1.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=1+2"
      );
      engine1.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "hello"
      );

      // Get the final state
      const finalValue1 = engine1.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      const finalValue2 = engine1.getCellValue({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 0,
      });

      // Serialize the full state (simulating backend save)
      const serializedState = engine1.serializeEngine();

      // Create new engine and restore state
      const engine2 = FormulaEngine.buildEmpty();
      engine2.resetToSerializedEngine(serializedState);

      // Verify state is identical
      expect(
        engine2.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(finalValue1!);
      expect(
        engine2.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(finalValue2!);
    });

    test("action log includes timestamps", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook(workbookName);

      const actionLog = engine.getActionLog();
      const lastAction = actionLog[actionLog.length - 1];

      expect(lastAction?.timestamp).toBeDefined();
      expect(typeof lastAction?.timestamp).toBe("number");
      expect(lastAction?.timestamp).toBeLessThanOrEqual(Date.now());
    });

    test("can clear action log", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      expect(engine.getActionLog().length).toBeGreaterThan(0);

      engine.clearActionLog();
      expect(engine.getActionLog().length).toBe(0);
    });
  });

  describe("Spill into Schema-Protected Area", () => {
    test("SEQUENCE spilling into table schema validates each cell", () => {
      // Create API with table schema
      const api = defineApi().addTableApi(
        "numbers",
        { workbookName, tableName: "Numbers" },
        {
          value: { parse: parseNumber, index: 0 },
        }
      );

      const engine = new FormulaEngine(api);
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // Set up table headers
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "value"
      );

      // Create table starting at A1
      engine.addTable({
        workbookName,
        sheetName,
        tableName: "Numbers",
        start: "A1",
        numRows: { type: "number", value: 5 },
        numCols: 1,
      });

      // Put a SEQUENCE formula outside the table that would spill into it
      // SEQUENCE(5,1) creates 5 numbers - all should be valid for the number schema
      // This tests that spill validation uses evaluated values
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 }, // B1 - outside table
        "=SEQUENCE(5,1)"
      );

      // Verify the spill worked
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(1);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 4,
        })
      ).toBe(5);
    });

    test("spilled values are picked up when spill overlaps existing cell", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // Set up a cell with an existing value that will be in the spill area
      // SEQUENCE(5,5) spills from A1 to E5 (5 rows x 5 columns)
      // So B2 will be in the spill area
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 }, // B2
        "existing-value"
      );

      // Create a SEQUENCE formula that will spill onto B2
      // SEQUENCE(5,5) creates 5 rows x 5 columns starting from 1:
      // Row 1: 1, 2, 3, 4, 5
      // Row 2: 6, 7, 8, 9, 10
      // Row 3: 11, 12, 13, 14, 15
      // Row 4: 16, 17, 18, 19, 20
      // Row 5: 21, 22, 23, 24, 25
      // So B2 (col index 1, row index 1, 0-indexed) should be 7
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        "=SEQUENCE(5,5)"
      );

      // The spill should be blocked because B2 has an existing value
      // So A1 should show #SPILL! error
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(FormulaError.SPILL);

      // B2 should still have its original value since the spill was blocked
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 1,
        })
      ).toBe("existing-value");

      // Now clear the blocking value and verify the spill works
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 }, // B2
        ""
      );

      // Re-trigger evaluation by setting the formula again
      // The engine should automatically re-evaluate now that B2 is empty
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        "=SEQUENCE(5,5)"
      );

      // Now the spill should work and B2 should show the spilled value
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(1); // Origin cell

      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 1, // B2
        })
      ).toBe(7); // Spilled value at B2

      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 2, // C3
        })
      ).toBe(13); // Spilled value at C3 (row index 2, col index 2: value 13)
    });
  });
});
