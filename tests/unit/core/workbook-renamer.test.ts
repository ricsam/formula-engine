import { test, expect, describe } from "bun:test";
import { 
  renameWorkbookInFormula,
  formulaReferencesWorkbook,
  getReferencedWorkbookNames
} from "../../../src/core/workbook-renamer";

describe("Workbook Renamer", () => {
  describe("renameWorkbookInFormula", () => {
    test("should rename simple cross-workbook reference", () => {
      const formula = "[MyWorkbook]Sheet1!A1+10";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("[NewWorkbook]Sheet1!A1+10");
    });

    test("should rename workbook sheet alias", () => {
      const formula = "[MyWorkbook]Sheet1";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("[NewWorkbook]Sheet1!A1:INFINITY");
    });

    test("should rename multiple references to same workbook", () => {
      const formula = "[MyWorkbook]Sheet1!A1+[MyWorkbook]Sheet2!B2*2";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("[NewWorkbook]Sheet1!A1+[NewWorkbook]Sheet2!B2*2");
    });

    test("should only rename matching workbook names", () => {
      const formula = "[MyWorkbook]Sheet1!A1+[OtherWorkbook]Sheet1!B1";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("[NewWorkbook]Sheet1!A1+[OtherWorkbook]Sheet1!B1");
    });

    test("should handle range references with workbook names", () => {
      const formula = "SUM([MyWorkbook]Sheet1!A1:B10)";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("SUM([NewWorkbook]Sheet1!A1:B10)");
    });

    test("should handle 3D range references with workbook names", () => {
      const formula = "SUM([MyWorkbook]Sheet1:Sheet3!A1:C5)";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("SUM([NewWorkbook]Sheet1:Sheet3!A1:C5)");
    });

    test("should handle table references with workbook names", () => {
      const formula = "[MyWorkbook]Sheet1!Table1[Column1]";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("[NewWorkbook]Sheet1!Table1[Column1]");
    });

    test("should handle complex workbook names with spaces and special characters", () => {
      const formula = "[Budget 2024]'Financial Data'!A1";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "Budget 2024", 
        newWorkbookName: "Budget 2025" 
      });
      expect(result).toBe("[Budget 2025]'Financial Data'!A1");
    });

    test("should handle complex formulas with nested workbook references", () => {
      const formula = "IF([MyWorkbook]Sheet1!A1>0,SUM([MyWorkbook]Sheet1!B1:B10),[OtherWorkbook]Sheet1!C1)";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("IF([NewWorkbook]Sheet1!A1>0,SUM([NewWorkbook]Sheet1!B1:B10),[OtherWorkbook]Sheet1!C1)");
    });

    test("should handle formulas with no workbook references", () => {
      const formula = "Sheet1!A1+B1*2";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("Sheet1!A1+B1*2");
    });

    test("should handle formulas with different workbook names", () => {
      const formula = "[OtherWorkbook]Sheet1!A1+[AnotherWorkbook]Sheet1!B1";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("[OtherWorkbook]Sheet1!A1+[AnotherWorkbook]Sheet1!B1");
    });

    test("should return original formula if parsing fails", () => {
      const formula = "INVALID(SYNTAX";
      const result = renameWorkbookInFormula({ 
        formula, 
        oldWorkbookName: "MyWorkbook", 
        newWorkbookName: "NewWorkbook" 
      });
      expect(result).toBe("INVALID(SYNTAX");
    });
  });

  describe("formulaReferencesWorkbook", () => {
    test("should return true for formula with workbook reference", () => {
      const formula = "[MyWorkbook]Sheet1!A1+10";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(true);
    });

    test("should return true for workbook sheet alias", () => {
      const formula = "[MyWorkbook]Sheet1";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(true);
    });

    test("should return false for formula without workbook reference", () => {
      const formula = "Sheet1!A1+B1*2";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(false);
    });

    test("should return false for formula with different workbook", () => {
      const formula = "[OtherWorkbook]Sheet1!A1+10";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(false);
    });

    test("should return true for formula with multiple references to same workbook", () => {
      const formula = "[MyWorkbook]Sheet1!A1+[MyWorkbook]Sheet2!B2";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(true);
    });

    test("should return true for formula with range reference", () => {
      const formula = "SUM([MyWorkbook]Sheet1!A1:B10)";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(true);
    });

    test("should return true for formula with table reference", () => {
      const formula = "[MyWorkbook]Sheet1!Table1[Column1]";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(true);
    });

    test("should return false for invalid formula", () => {
      const formula = "INVALID(SYNTAX";
      expect(formulaReferencesWorkbook(formula, "MyWorkbook")).toBe(false);
    });
  });

  describe("getReferencedWorkbookNames", () => {
    test("should return single workbook name", () => {
      const formula = "[MyWorkbook]Sheet1!A1+10";
      const result = getReferencedWorkbookNames(formula);
      expect(result).toEqual(["MyWorkbook"]);
    });

    test("should return workbook name from sheet alias", () => {
      const formula = "[MyWorkbook]Sheet1";
      const result = getReferencedWorkbookNames(formula);
      expect(result).toEqual(["MyWorkbook"]);
    });

    test("should return multiple unique workbook names", () => {
      const formula = "[MyWorkbook]Sheet1!A1+[OtherWorkbook]Sheet1!B1+[ThirdWorkbook]Sheet1!C1";
      const result = getReferencedWorkbookNames(formula);
      expect(result.sort()).toEqual(["MyWorkbook", "OtherWorkbook", "ThirdWorkbook"]);
    });

    test("should return unique workbook names (no duplicates)", () => {
      const formula = "[MyWorkbook]Sheet1!A1+[MyWorkbook]Sheet2!B1+[OtherWorkbook]Sheet1!C1";
      const result = getReferencedWorkbookNames(formula);
      expect(result.sort()).toEqual(["MyWorkbook", "OtherWorkbook"]);
    });

    test("should return empty array for formula with no workbook references", () => {
      const formula = "Sheet1!A1+B1*2";
      const result = getReferencedWorkbookNames(formula);
      expect(result).toEqual([]);
    });

    test("should handle complex nested formulas", () => {
      const formula = "IF([MyWorkbook]Sheet1!A1>0,SUM([OtherWorkbook]Sheet1!B1:B10),[ThirdWorkbook]Sheet1!C1)";
      const result = getReferencedWorkbookNames(formula);
      expect(result.sort()).toEqual(["MyWorkbook", "OtherWorkbook", "ThirdWorkbook"]);
    });

    test("should handle table references", () => {
      const formula = "[MyWorkbook]Sheet1!Table1[Column1]+[OtherWorkbook]Sheet1!Table2[Column2]";
      const result = getReferencedWorkbookNames(formula);
      expect(result.sort()).toEqual(["MyWorkbook", "OtherWorkbook"]);
    });

    test("should handle 3D ranges", () => {
      const formula = "SUM([MyWorkbook]Sheet1:Sheet3!A1:C5)";
      const result = getReferencedWorkbookNames(formula);
      expect(result).toEqual(["MyWorkbook"]);
    });

    test("should return empty array for invalid formula", () => {
      const formula = "INVALID(SYNTAX";
      const result = getReferencedWorkbookNames(formula);
      expect(result).toEqual([]);
    });
  });
});
