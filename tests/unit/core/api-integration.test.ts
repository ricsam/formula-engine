import { describe, test, expect, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { defineApi } from "../../../src/core/api/api";
import { SchemaValidationError } from "../../../src/core/managers/api-schema-manager";

describe("API Integration", () => {
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

  // Define the API schema
  const userApi = defineApi()
    .addTableApi(
      "users",
      {
        workbookName,
        tableName: "Users",
      },
      {
        id: {
          parse: (value) => parseNumber(value),
          index: 0,
        },
        name: {
          parse: (value) => parseString(value),
          index: 1,
        },
        email: {
          parse: (value) => parseString(value),
          index: 2,
        },
        age: {
          parse: (value) => parseNumber(value),
          index: 3,
        },
      },
      {
        get(id: number) {
          return this.findWhere({ id });
        },
        getAll() {
          return this.findAllWhere({});
        },
        create(newUser: { name: string; email: string; age: number }) {
          return this.append({
            id: this.count() + 1,
            ...newUser,
          });
        },
        update(id: number, update: { name?: string; email?: string; age?: number }) {
          return this.updateWhere({ id }, update);
        },
        delete(id: number) {
          return this.removeWhere({ id });
        },
        count() {
          return this.count();
        },
      }
    )
    .addCellApi(
      "config",
      {
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 0,
      },
      (value) => {
        return parseString(value);
      },
      {
        get() {
          return this.read();
        },
        set(value: string) {
          this.write(value);
        },
      }
    );

  describe("defineApi", () => {
    test("api property is undefined at runtime", () => {
      expect(userApi.api).toBeUndefined();
    });

    test("declaration contains the schema definitions", () => {
      expect(userApi.declaration.users).toBeDefined();
      expect(userApi.declaration.users.type).toBe("table");
      expect(userApi.declaration.config).toBeDefined();
      expect(userApi.declaration.config.type).toBe("cell");
    });
  });

  describe("Engine with API", () => {
    let engine: FormulaEngine<any, typeof userApi>;

    beforeEach(() => {
      engine = new FormulaEngine(userApi);

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

    test("engine.api is defined when API is provided", () => {
      expect(engine.api).toBeDefined();
      expect(engine.api.users).toBeDefined();
      expect(engine.api.config).toBeDefined();
    });

    test("engine.api is undefined when no API is provided", () => {
      const engineWithoutApi = new FormulaEngine();
      expect(engineWithoutApi.api).toBeUndefined();
    });

    describe("TableOrm operations", () => {
      test("append adds a new row", () => {
        const user = engine.api.users.create({
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
        expect(engine.getCellValue({ workbookName, sheetName, colIndex: 0, rowIndex: 1 })).toBe(1);
        expect(engine.getCellValue({ workbookName, sheetName, colIndex: 1, rowIndex: 1 })).toBe("John Doe");
        expect(engine.getCellValue({ workbookName, sheetName, colIndex: 2, rowIndex: 1 })).toBe("john@example.com");
        expect(engine.getCellValue({ workbookName, sheetName, colIndex: 3, rowIndex: 1 })).toBe(30);
      });

      test("findWhere finds a row by filter", () => {
        // Add some users
        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });
        engine.api.users.create({ name: "Bob", email: "bob@example.com", age: 35 });

        const user = engine.api.users.get(2);
        expect(user).toEqual({
          id: 2,
          name: "Bob",
          email: "bob@example.com",
          age: 35,
        });
      });

      test("findWhere returns undefined when not found", () => {
        const user = engine.api.users.get(999);
        expect(user).toBeUndefined();
      });

      test("findAllWhere returns all matching rows", () => {
        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });
        engine.api.users.create({ name: "Bob", email: "bob@example.com", age: 25 });
        engine.api.users.create({ name: "Charlie", email: "charlie@example.com", age: 30 });

        const users = engine.api.users.getAll();
        expect(users).toHaveLength(3);
      });

      test("updateWhere updates matching rows", () => {
        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });

        const updated = engine.api.users.update(1, { age: 26 });
        expect(updated).toBe(1);

        const user = engine.api.users.get(1);
        expect(user?.age).toBe(26);
      });

      test("removeWhere deletes matching rows", () => {
        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });
        engine.api.users.create({ name: "Bob", email: "bob@example.com", age: 35 });

        const removed = engine.api.users.delete(1);
        expect(removed).toBe(1);

        const user = engine.api.users.get(1);
        expect(user).toBeUndefined();
      });

      test("count returns the number of rows", () => {
        expect(engine.api.users.count()).toBe(0);

        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });
        expect(engine.api.users.count()).toBe(1);

        engine.api.users.create({ name: "Bob", email: "bob@example.com", age: 35 });
        expect(engine.api.users.count()).toBe(2);
      });
    });

    describe("CellOrm operations", () => {
      test("read returns the cell value", () => {
        engine.setCellContent(
          { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
          "test-config"
        );

        const value = engine.api.config.get();
        expect(value).toBe("test-config");
      });

      test("write sets the cell value", () => {
        engine.api.config.set("new-config");

        const value = engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 0,
        });
        expect(value).toBe("new-config");
      });
    });

    describe("Schema validation", () => {
      test("setCellContent throws SchemaValidationError for invalid data in table range", () => {
        // First add a valid user to establish the table has data
        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });

        // Try to write invalid data (string where number expected) to the age column
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 3, rowIndex: 1 },
            "not-a-number"
          );
        }).toThrow(SchemaValidationError);
      });

      test("setCellContent throws SchemaValidationError for invalid data in cell schema", () => {
        // Try to write a number where string expected
        expect(() => {
          engine.setCellContent(
            { workbookName, sheetName, colIndex: 5, rowIndex: 0 },
            123
          );
        }).toThrow(SchemaValidationError);
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
    });

    describe("Schema lifecycle", () => {
      test("schema is invalidated when table is deleted", () => {
        engine.api.users.create({ name: "Alice", email: "alice@example.com", age: 25 });

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
  });
});

