import { test, expect, describe } from "bun:test";
import { 
  renameSheetInFormula, 
  formulaReferencesSheet, 
  getReferencedSheetNames 
} from "../../../src/core/sheet-renamer";

describe("Sheet Renamer", () => {
  describe("renameSheetInFormula", () => {
    test("should rename simple cross-sheet reference", () => {
      const formula = "Sheet1!A1+10";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("DataSheet!A1+10");
    });

    test("should rename multiple references to same sheet", () => {
      const formula = "Sheet1!A1+Sheet1!B2*2";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("DataSheet!A1+DataSheet!B2*2");
    });

    test("should only rename matching sheet names", () => {
      const formula = "Sheet1!A1+Sheet2!B1";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("DataSheet!A1+Sheet2!B1");
    });

    test("should handle range references with sheet names", () => {
      const formula = "SUM(Sheet1!A1:B10)";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("SUM(DataSheet!A1:B10)");
    });

    test("should handle complex formulas with nested references", () => {
      const formula = "IF(Sheet1!A1>0,SUM(Sheet1!B1:B10),Sheet2!C1)";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("IF(DataSheet!A1>0,SUM(DataSheet!B1:B10),Sheet2!C1)");
    });

    test("should handle formulas with no sheet references", () => {
      const formula = "A1+B1*2";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("A1+B1*2");
    });

    test("should handle formulas with different sheet names", () => {
      const formula = "Sheet2!A1+Sheet3!B1";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("Sheet2!A1+Sheet3!B1");
    });

    test("should return original formula for unsupported 3D ranges", () => {
      const formula = "SUM(Sheet1:Sheet3!A1)";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      // 3D ranges are not supported yet, so formula should remain unchanged
      expect(result).toBe("SUM(Sheet1:Sheet3!A1)");
    });

    test("should return original formula if parsing fails", () => {
      const formula = "INVALID(SYNTAX";
      const result = renameSheetInFormula({ formula, oldSheetName: "Sheet1", newSheetName: "DataSheet" });
      expect(result).toBe("INVALID(SYNTAX");
    });
  });

  describe("formulaReferencesSheet", () => {
    test("should return true for formula with sheet reference", () => {
      const formula = "Sheet1!A1+10";
      expect(formulaReferencesSheet(formula, "Sheet1")).toBe(true);
    });

    test("should return false for formula without sheet reference", () => {
      const formula = "A1+B1*2";
      expect(formulaReferencesSheet(formula, "Sheet1")).toBe(false);
    });

    test("should return false for formula with different sheet", () => {
      const formula = "Sheet2!A1+10";
      expect(formulaReferencesSheet(formula, "Sheet1")).toBe(false);
    });

    test("should return true for formula with multiple references to same sheet", () => {
      const formula = "Sheet1!A1+Sheet1!B2";
      expect(formulaReferencesSheet(formula, "Sheet1")).toBe(true);
    });

    test("should return true for formula with range reference", () => {
      const formula = "SUM(Sheet1!A1:B10)";
      expect(formulaReferencesSheet(formula, "Sheet1")).toBe(true);
    });

    test("should return false for invalid formula", () => {
      const formula = "INVALID(SYNTAX";
      expect(formulaReferencesSheet(formula, "Sheet1")).toBe(false);
    });
  });

  describe("getReferencedSheetNames", () => {
    test("should return single sheet name", () => {
      const formula = "Sheet1!A1+10";
      const result = getReferencedSheetNames(formula);
      expect(result).toEqual(["Sheet1"]);
    });

    test("should return multiple unique sheet names", () => {
      const formula = "Sheet1!A1+Sheet2!B1+Sheet3!C1";
      const result = getReferencedSheetNames(formula);
      expect(result.sort()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    });

    test("should return unique sheet names (no duplicates)", () => {
      const formula = "Sheet1!A1+Sheet1!B1+Sheet2!C1";
      const result = getReferencedSheetNames(formula);
      expect(result.sort()).toEqual(["Sheet1", "Sheet2"]);
    });

    test("should return empty array for formula with no sheet references", () => {
      const formula = "A1+B1*2";
      const result = getReferencedSheetNames(formula);
      expect(result).toEqual([]);
    });

    test("should handle complex nested formulas", () => {
      const formula = "IF(Sheet1!A1>0,SUM(Sheet2!B1:B10),Sheet3!C1)";
      const result = getReferencedSheetNames(formula);
      expect(result.sort()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    });

    test("should return empty array for invalid formula", () => {
      const formula = "INVALID(SYNTAX";
      const result = getReferencedSheetNames(formula);
      expect(result).toEqual([]);
    });

    test("should return empty array for unsupported 3D ranges", () => {
      const formula = "SUM(Sheet1:Sheet3!A1)";
      const result = getReferencedSheetNames(formula);
      // 3D ranges are not supported yet, so no sheet names should be detected
      expect(result).toEqual([]);
    });
  });

});
