import { test, expect, describe } from "bun:test";
import { 
  renameTableInFormula, 
  formulaReferencesTable, 
  getReferencedTableNames 
} from "../../../src/core/table-renamer";

describe("Table Renamer", () => {
  describe("renameTableInFormula", () => {
    test("should rename simple table reference", () => {
      const formula = "SUM(Products[Price])";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price])");
    });

    test("should rename multiple references to same table", () => {
      const formula = "SUM(Products[Price]) + AVERAGE(Products[Quantity])";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price])+AVERAGE(Items[Quantity])");
    });

    test("should only rename matching table names", () => {
      const formula = "SUM(Products[Price]) + SUM(Orders[Total])";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price])+SUM(Orders[Total])");
    });

    test("should handle complex formulas with nested references", () => {
      const formula = "IF(Products[Price] > 100, Products[Price] * 0.9, Products[Price])";
      const result = renameTableInFormula(formula, "Products", "Inventory");
      
      expect(result).toBe("IF(Inventory[Price]>100,Inventory[Price]*0.9,Inventory[Price])");
    });

    test("should handle formulas with no table references", () => {
      const formula = "A1 + B1 * 2";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("A1+B1*2");
    });

    test("should handle formulas with different table names", () => {
      const formula = "SUM(Orders[Total])";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Orders[Total])");
    });

    test("should handle current row references", () => {
      const formula = "Products[@Price] * 1.1";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("Items[@Price]*1.1");
    });

    test("should handle column range references", () => {
      const formula = "SUM(Products[[Price]:[Quantity]])";
      const result = renameTableInFormula(formula, "Products", "Items");
      
      expect(result).toBe("SUM(Items[Price:Quantity])");
    });

    test("should return original formula if parsing fails", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = renameTableInFormula(invalidFormula, "Products", "Items");
      
      expect(result).toBe(invalidFormula);
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
