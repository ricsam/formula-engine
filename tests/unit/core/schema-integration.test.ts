import { describe, test, expect, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { createHeader, defineSchema } from "../../../src/core/schema/schema";
import { SchemaIntegrityError } from "../../../src/core/commands/command-executor";

describe("Schema Integration", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "Sheet1";

  // Parse functions
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

  // Define the schema
  const userSchema = defineSchema()
    .addTableSchema(
      "users",
      {
        workbookName,
        tableName: "Users",
      },
      {
        id: createHeader(0, parseNumber),
        name: createHeader(1, parseString),
        email: createHeader(2, parseString),
        age: createHeader(3, parseNumber),
      }
    )
    .addCellSchema(
      "config",
      {
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 0,
      },
      (value) => {
        return parseString(value);
      }
    )
    .addGridSchema(
      "plate",
      { workbookName, sheetName: "GridSheet" },
      { start: { col: 0, row: 0 }, end: { col: 2, row: 2 } }, // 3x3 grid (A1:C3)
      (value) => {
        return parseNumber(value);
      }
    );

  describe("defineSchema", () => {
    test("schema property is undefined at runtime", () => {
      expect(userSchema.schema).toBeUndefined();
    });

    test("declaration contains the schema definitions", () => {
      expect(userSchema.declaration.users).toBeDefined();
      expect(userSchema.declaration.users.type).toBe("table");
      expect(userSchema.declaration.config).toBeDefined();
      expect(userSchema.declaration.config.type).toBe("cell");
      expect(userSchema.declaration.plate).toBeDefined();
      expect(userSchema.declaration.plate.type).toBe("grid");
    });
  });

  describe("Engine with Schema", () => {
    let engine: FormulaEngine<any, typeof userSchema>;

    beforeEach(() => {
      engine = new FormulaEngine(userSchema);

      // Set up workbook and sheet
      engine.addWorkbook(workbookName);
      engine.addSheet({ workbookName, sheetName });

      // Create the Users table with headers
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map([
          ["A1", "id"],
          ["B1", "name"],
          ["C1", "email"],
          ["D1", "age"],
        ])
      );

      // Add the table definition
      engine.addTable({
        workbookName,
        sheetName,
        tableName: "Users",
        start: "A1",
        numRows: { type: "infinity", sign: "positive" },
        numCols: 4,
      });
    });

    test("engine.schema is defined when schema is provided", () => {
      expect(engine.schema).toBeDefined();
      expect(engine.schema.users).toBeDefined();
      expect(engine.schema.config).toBeDefined();
      expect(engine.schema.plate).toBeDefined();
    });

    test("engine.schema is defined when no schema is provided", () => {
      const engineWithoutSchema = new FormulaEngine();
      expect(engineWithoutSchema.schema).toBeDefined();
    });

    describe("TableOrm operations", () => {
      test("append adds a new row", () => {
        const user = engine.schema.users.append({
          id: 1,
          name: "John Doe",
          email: "john@example.com",
          age: 30,
        });

        expect(user).toEqual({
          id: 1,
          name: "John Doe",
          email: "john@example.com",
          age: 30,
        });

        // Verify data was written to cells
        expect(
          engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: 0,
            rowIndex: 1,
          })
        ).toBe(1);
        expect(
          engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: 1,
            rowIndex: 1,
          })
        ).toBe("John Doe");
        expect(
          engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: 2,
            rowIndex: 1,
          })
        ).toBe("john@example.com");
        expect(
          engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: 3,
            rowIndex: 1,
          })
        ).toBe(30);
      });

      test("findWhere finds a row by filter", () => {
        // Add some users
        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });
        engine.schema.users.append({
          id: 2,
          name: "Bob",
          email: "bob@example.com",
          age: 35,
        });

        const user = engine.schema.users.findWhere({ id: 2 });
        expect(user).toEqual({
          id: 2,
          name: "Bob",
          email: "bob@example.com",
          age: 35,
        });
      });

      test("findWhere returns undefined when not found", () => {
        const user = engine.schema.users.findWhere({ id: 999 });
        expect(user).toBeUndefined();
      });

      test("findAllWhere returns all matching rows", () => {
        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });
        engine.schema.users.append({
          id: 2,
          name: "Bob",
          email: "bob@example.com",
          age: 25,
        });
        engine.schema.users.append({
          id: 3,
          name: "Charlie",
          email: "charlie@example.com",
          age: 30,
        });

        const users = engine.schema.users.findAllWhere({});
        expect(users).toHaveLength(3);
      });

      test("updateWhere updates matching rows", () => {
        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });

        const updated = engine.schema.users.updateWhere({ id: 1 }, { age: 26 });
        expect(updated).toBe(1);

        const user = engine.schema.users.findWhere({ id: 1 });
        expect(user?.age).toBe(26);
      });

      test("removeWhere deletes matching rows", () => {
        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });
        engine.schema.users.append({
          id: 2,
          name: "Bob",
          email: "bob@example.com",
          age: 35,
        });

        const removed = engine.schema.users.removeWhere({ id: 1 });
        expect(removed).toBe(1);

        const user = engine.schema.users.findWhere({ id: 1 });
        expect(user).toBeUndefined();
      });

      test("count returns the number of rows", () => {
        expect(engine.schema.users.count()).toBe(0);

        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });
        expect(engine.schema.users.count()).toBe(1);

        engine.schema.users.append({
          id: 2,
          name: "Bob",
          email: "bob@example.com",
          age: 35,
        });
        expect(engine.schema.users.count()).toBe(2);
      });
    });

    describe("CellOrm operations", () => {
      test("read returns the cell value", () => {
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
          "test-config"
        );

        const value = engine.schema.config.read();
        expect(value).toBe("test-config");
      });

      test("write sets the cell value", () => {
        engine.schema.config.write("new-config");

        const value = engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 0,
        });
        expect(value).toBe("new-config");
      });
    });

    describe("GridOrm operations", () => {
      const gridSheetName = "GridSheet";

      beforeEach(() => {
        // Add the grid sheet
        engine.addSheet({ workbookName, sheetName: gridSheetName });
      });

      test("columns getter returns column-major 2D array", () => {
        // Set up a 3x3 grid of values (A1:C3)
        // Column 0: 1, 4, 7
        // Column 1: 2, 5, 8
        // Column 2: 3, 6, 9
        engine.setSheetContent(
          { workbookName, sheetName: gridSheetName },
          new Map([
            ["A1", 1], ["B1", 2], ["C1", 3],
            ["A2", 4], ["B2", 5], ["C2", 6],
            ["A3", 7], ["B3", 8], ["C3", 9],
          ])
        );

        const columns = engine.schema.plate.columns;

        // columns[colIndex][rowIndex]
        expect(columns).toEqual([
          [1, 4, 7], // Column A
          [2, 5, 8], // Column B
          [3, 6, 9], // Column C
        ]);
      });

      test("rows getter returns row-major 2D array", () => {
        // Set up a 3x3 grid of values (A1:C3)
        engine.setSheetContent(
          { workbookName, sheetName: gridSheetName },
          new Map([
            ["A1", 1], ["B1", 2], ["C1", 3],
            ["A2", 4], ["B2", 5], ["C2", 6],
            ["A3", 7], ["B3", 8], ["C3", 9],
          ])
        );

        const rows = engine.schema.plate.rows;

        // rows[rowIndex][colIndex]
        expect(rows).toEqual([
          [1, 2, 3], // Row 1
          [4, 5, 6], // Row 2
          [7, 8, 9], // Row 3
        ]);
      });

      test("setValue writes a value at the specified position", () => {
        engine.schema.plate.setValue(42, { col: 1, row: 1 });

        // Verify the cell was written
        expect(engine.getCellValue({ workbookName, sheetName: gridSheetName, colIndex: 1, rowIndex: 1 })).toBe(42);
      });

      test("getValue reads a value at the specified position", () => {
        // Set up a value
        engine.setCellContent(
          { workbookName, sheetName: gridSheetName, colIndex: 2, rowIndex: 0 },
          99
        );

        const value = engine.schema.plate.getValue({ col: 2, row: 0 });
        expect(value).toBe(99);
      });

      test("setValue throws error for out of bounds position", () => {
        expect(() => {
          engine.schema.plate.setValue(1, { col: 5, row: 0 }); // col 5 is out of bounds (grid is 3x3)
        }).toThrow('out of bounds');

        expect(() => {
          engine.schema.plate.setValue(1, { col: 0, row: 10 }); // row 10 is out of bounds
        }).toThrow('out of bounds');

        expect(() => {
          engine.schema.plate.setValue(1, { col: -1, row: 0 }); // negative col
        }).toThrow('out of bounds');
      });

      test("getValue throws error for out of bounds position", () => {
        expect(() => {
          engine.schema.plate.getValue({ col: 5, row: 0 });
        }).toThrow('out of bounds');
      });
    });

    describe("Schema validation", () => {
      test("setCellContent throws SchemaIntegrityError for invalid data in table range", () => {
        // First add a valid user to establish the table has data
        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });

        // Try to write invalid data (string where number expected) to the age column
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 3, rowIndex: 1 },
            "not-a-number"
          );
        }).toThrow(SchemaIntegrityError);
      });

      test("setCellContent throws SchemaIntegrityError for invalid data in cell schema", () => {
        // Try to write a number where string expected
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
            123
          );
        }).toThrow(SchemaIntegrityError);
      });

      test("setCellContent allows valid data", () => {
        // Write valid data to table
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 3, rowIndex: 1 },
            30
          );
        }).not.toThrow();

        // Write valid data to cell schema
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
            "valid-string"
          );
        }).not.toThrow();
      });

      test("setCellContent throws SchemaIntegrityError for invalid data in grid range", () => {
        // Add the grid sheet
        engine.addSheet({ workbookName, sheetName: "GridSheet" });

        // Try to write invalid data (string where number expected) to grid cell
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName: "GridSheet", colIndex: 1, rowIndex: 1 },
            "not-a-number"
          );
        }).toThrow(SchemaIntegrityError);
      });

      test("setCellContent allows valid data in grid range", () => {
        // Add the grid sheet
        engine.addSheet({ workbookName, sheetName: "GridSheet" });

        // Write valid number to grid cell
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName: "GridSheet", colIndex: 1, rowIndex: 1 },
            42
          );
        }).not.toThrow();
      });
    });

    describe("Schema lifecycle", () => {
      test("schema is invalidated when table is deleted", () => {
        engine.schema.users.append({
          id: 1,
          name: "Alice",
          email: "alice@example.com",
          age: 25,
        });

        // Delete the table
        engine.removeTable({ workbookName, tableName: "Users" });

        // Writing invalid data should now be allowed (schema is invalidated)
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 3, rowIndex: 1 },
            "not-a-number"
          );
        }).not.toThrow();
      });
    });

    describe("Spill validation into schema-protected cells", () => {
      test("spilling formula into cell with number schema constraint throws error when invalid", () => {
        // Create a cell API schema that requires a number at B2 (colIndex 1, rowIndex 1)
        const cellSchema = defineSchema().addCellSchema(
          "numberCell",
          {
            workbookName,
            sheetName,
            colIndex: 1, // B2
            rowIndex: 1,
          },
          (value) => {
            if (typeof value !== "number") {
              throw new Error("Expected a number value");
            }
            return value;
          }
        );

        const engineWithCellSchema = new FormulaEngine(cellSchema);
        engineWithCellSchema.addWorkbook(workbookName);
        engineWithCellSchema.addSheet({ workbookName, sheetName });

        engineWithCellSchema.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 1 }, // A2
          "=E1:H1"
        );

        expect(() => {
          engineWithCellSchema.setCellContent(
            { workbookName, sheetName, colIndex: 5, rowIndex: 0 }, // F1
            "string"
          );
        }).toThrow(SchemaIntegrityError);
        expect(() => {
          engineWithCellSchema.setCellContent(
            { workbookName, sheetName, colIndex: 5, rowIndex: 0 }, // F1
            123
          );
        }).not.toThrow(SchemaIntegrityError);
      });

      test("spilling formula into table area returns #SPILL! error on origin cell", () => {
        // Excel behavior: formulas cannot spill into tables - they get #SPILL! error
        const engine = new FormulaEngine();
        engine.addWorkbook(workbookName);
        engine.addSheet({ workbookName, sheetName });

        // Set up table headers at B2
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
          "value"
        );

        // Create table starting at B2 with infinite rows
        engine.addTable({
          workbookName,
          sheetName,
          tableName: "Numbers",
          start: "B2",
          numRows: { type: "infinity", sign: "positive" },
          numCols: 1,
        });

        // Fill E1:H1 with values so the spill has something to reference
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 4, rowIndex: 0 },
          1
        ); // E1
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
          2
        ); // F1
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 6, rowIndex: 0 },
          3
        ); // G1
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 7, rowIndex: 0 },
          4
        ); // H1

        // Add spilling formula at A5 that references E1:H1
        // This would spill into B5 which is in the table - should get #SPILL! error
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 4 }, // A5
          "=E1:H1"
        );

        // The origin cell (A5) should show #SPILL! error
        const originValue = engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 4,
        });
        expect(originValue).toBe("#SPILL!");

        // The table cell (B5) should be empty (no spill occurred)
        const tableCell = engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 4,
        });
        expect(tableCell).toBe("");
      });

      test("spilling formula inside a table returns #SPILL! error", () => {
        // Excel behavior: spilling formulas cannot exist inside tables
        const engine = new FormulaEngine();
        engine.addWorkbook(workbookName);
        engine.addSheet({ workbookName, sheetName });

        // Set up table headers at A1
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
          "formula"
        );
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
          "col2"
        );

        // Create table starting at A1 with 2 columns
        engine.addTable({
          workbookName,
          sheetName,
          tableName: "TestTable",
          start: "A1",
          numRows: { type: "infinity", sign: "positive" },
          numCols: 2,
        });

        // Fill some source data outside the table (E1:H1)
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 4, rowIndex: 0 },
          1
        );
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
          2
        );
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 6, rowIndex: 0 },
          3
        );
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 7, rowIndex: 0 },
          4
        );

        // Add a spilling formula inside the table (A2 is in the table data area)
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 0, rowIndex: 1 }, // A2
          "=E1:H1"
        );

        // The formula cell inside the table should show #SPILL! error
        const cellValue = engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 1,
        });
        expect(cellValue).toBe("#SPILL!");
      });
    });
  });
});
