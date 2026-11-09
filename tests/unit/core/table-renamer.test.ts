import { test, expect, describe } from "bun:test";
import { 
  renameTableInFormula, 
  renameTableColumnsInFormula,
  formulaReferencesTable, 
  getReferencedTableNames 
} from "../../../src/core/table-renamer";

describe("Table Renamer", () => {
  describe("renameTableInFormula", () => {
    const rename = (
      formula: string,
      oldTableName: string,
      newTableName: string,
      workbooks: { tableWorkbookName?: string; formulaWorkbookName?: string } = {}
    ) =>
      renameTableInFormula({
        formula,
        oldTableName,
        newTableName,
        tableWorkbookName: workbooks.tableWorkbookName ?? "Book",
        formulaWorkbookName: workbooks.formulaWorkbookName ?? "Book",
      });

    test("should rename simple table reference", () => {
      const formula = "SUM(Products[Price])";
      const result = rename(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price])");
    });

    test("should rename multiple references to same table", () => {
      const formula = "SUM(Products[Price]) + AVERAGE(Products[Quantity])";
      const result = rename(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price])+AVERAGE(Items[Quantity])");
    });

    test("should only rename matching table names", () => {
      const formula = "SUM(Products[Price]) + SUM(Orders[Total])";
      const result = rename(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price])+SUM(Orders[Total])");
    });

    test("should handle complex formulas with nested references", () => {
      const formula = "IF(Products[Price] > 100, Products[Price] * 0.9, Products[Price])";
      const result = rename(formula, "Products", "Inventory");
      
      expect(result).toBe("IF(Inventory[Price]>100,Inventory[Price]*0.9,Inventory[Price])");
    });

    test("should handle formulas with no table references", () => {
      const formula = "A1 + B1 * 2";
      const result = rename(formula, "Products", "Items");
      
      expect(result).toBe("A1+B1*2");
    });

    test("should handle formulas with different table names", () => {
      const formula = "SUM(Orders[Total])";
      const result = rename(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Orders[Total])");
    });

    test("should not rename a matching table in another workbook", () => {
      const formula = "SUM([Book]Products[Price])+SUM([Other]Products[Price])";
      const result = rename(formula, "Products", "Items");
      expect(result).toBe(
        "SUM([Book]Items[Price])+SUM([Other]Products[Price])"
      );
    });

    test("should rename workbook-qualified references in formulas of other workbooks", () => {
      const formula =
        "SUM([Book]Products[Price])+SUM(Products[Price])+[Book]Products[@Price]";
      const result = rename(formula, "Products", "Items", {
        tableWorkbookName: "Book",
        formulaWorkbookName: "Collation",
      });
      // The unqualified reference names the Collation workbook's own table.
      expect(result).toBe(
        "SUM([Book]Items[Price])+SUM(Products[Price])+[Book]Items[@Price]"
      );
    });

    test("should rename unqualified references when the formula has no workbook", () => {
      // A global named expression is evaluated in the workbook that uses it.
      const result = renameTableInFormula({
        formula: "SUM(Products[Price])+SUM([Other]Products[Price])",
        oldTableName: "Products",
        newTableName: "Items",
        tableWorkbookName: "Book",
      });
      expect(result).toBe("SUM(Items[Price])+SUM([Other]Products[Price])");
    });

    test("should handle current row references", () => {
      const formula = "Products[@Price] * 1.1";
      const result = rename(formula, "Products", "Items");
      
      expect(result).toBe("Items[@Price]*1.1");
    });

    test("should handle column range references", () => {
      const formula = "SUM(Products[[Price]:[Quantity]])";
      const result = rename(formula, "Products", "Items");

      // Column ranges always use double-bracket syntax to avoid ambiguity with colons in column names
      expect(result).toBe("SUM(Items[[Price]:[Quantity]])");
    });

    test("should return original formula if parsing fails", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = rename(invalidFormula, "Products", "Items");
      
      expect(result).toBe(invalidFormula);
    });
  });

  describe("renameTableColumnsInFormula", () => {
    test("should rename columns of a table in another workbook", () => {
      const result = renameTableColumnsInFormula({
        formula:
          "SUM([Book]Products[Price])+SUM([Other]Products[Price])+SUM(Products[Price])",
        tableName: "Products",
        tableWorkbookName: "Book",
        formulaWorkbookName: "Collation",
        columnRenames: new Map([["Price", "Cost"]]),
      });
      expect(result).toBe(
        "SUM([Book]Products[Cost])+SUM([Other]Products[Price])+SUM(Products[Price])"
      );
    });

    test("should rename current-row references and column spans", () => {
      const result = renameTableColumnsInFormula({
        formula: "[Book]Products[@Price]+SUM([Book]Products[[Price]:[Qty]])",
        tableName: "Products",
        tableWorkbookName: "Book",
        formulaWorkbookName: "Collation",
        columnRenames: new Map([
          ["Price", "Cost"],
          ["Qty", "Quantity"],
        ]),
      });
      expect(result).toBe(
        "[Book]Products[@Cost]+SUM([Book]Products[[Cost]:[Quantity]])"
      );
    });
  });

  describe("formulaReferencesTable", () => {
    test("should return true for formula with table reference", () => {
      const formula = "SUM(Products[Price])";
      const result = formulaReferencesTable(formula, "Products");
      
      expect(result).toBe(true);
    });

    test("should return false for formula without table reference", () => {
      const formula = "A1 + B1";
      const result = formulaReferencesTable(formula, "Products");
      
      expect(result).toBe(false);
    });

    test("should return false for formula with different table", () => {
      const formula = "SUM(Orders[Total])";
      const result = formulaReferencesTable(formula, "Products");
      
      expect(result).toBe(false);
    });

    test("should return true for formula with multiple references to same table", () => {
      const formula = "Products[Price] + Products[Tax]";
      const result = formulaReferencesTable(formula, "Products");
      
      expect(result).toBe(true);
    });

    test("should return true for formula with current row reference", () => {
      const formula = "Products[@Price] * 1.1";
      const result = formulaReferencesTable(formula, "Products");
      
      expect(result).toBe(true);
    });

    test("should return false for invalid formula", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = formulaReferencesTable(invalidFormula, "Products");
      
      expect(result).toBe(false);
    });
  });

  describe("getReferencedTableNames", () => {
    test("should return single table name", () => {
      const formula = "SUM(Products[Price])";
      const result = getReferencedTableNames(formula);
      
      expect(result).toEqual(["Products"]);
    });

    test("should return multiple unique table names", () => {
      const formula = "SUM(Products[Price]) + SUM(Orders[Total])";
      const result = getReferencedTableNames(formula);
      
      expect(result.sort()).toEqual(["Orders", "Products"]);
    });

    test("should return unique table names (no duplicates)", () => {
      const formula = "Products[Price] + Products[Tax] + Orders[Total]";
      const result = getReferencedTableNames(formula);
      
      expect(result.sort()).toEqual(["Orders", "Products"]);
    });

    test("should return empty array for formula with no table references", () => {
      const formula = "A1 + B1 * 2";
      const result = getReferencedTableNames(formula);
      
      expect(result).toEqual([]);
    });

    test("should handle complex nested formulas", () => {
      const formula = "IF(Products[Price] > Orders[MinPrice], Inventory[Stock], 0)";
      const result = getReferencedTableNames(formula);
      
      expect(result.sort()).toEqual(["Inventory", "Orders", "Products"]);
    });

    test("should return empty array for invalid formula", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = getReferencedTableNames(invalidFormula);
      
      expect(result).toEqual([]);
    });

    test("should handle current row and column range references", () => {
      const formula = "Products[@Price] + SUM(Orders[[Total]:[Tax]])";
      const result = getReferencedTableNames(formula);
      
      expect(result.sort()).toEqual(["Orders", "Products"]);
    });
  });
});
