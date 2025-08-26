import { test, expect, describe } from "bun:test";
import { 
  renameNamedExpressionInFormula, 
  formulaReferencesNamedExpression, 
  getReferencedNamedExpressionNames 
} from "../../../src/core/named-expression-renamer";

describe("Named Expression Renamer", () => {
  describe("renameNamedExpressionInFormula", () => {
    test("should rename simple named expression reference", () => {
      const formula = "TAX_RATE * 100";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("SALES_TAX*100");
    });

    test("should rename multiple references to same named expression", () => {
      const formula = "TAX_RATE + TAX_RATE * 0.1";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("SALES_TAX+SALES_TAX*0.1");
    });

    test("should only rename matching named expression names", () => {
      const formula = "TAX_RATE + DISCOUNT_RATE";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("SALES_TAX+DISCOUNT_RATE");
    });

    test("should handle complex formulas with nested references", () => {
      const formula = "IF(PRICE > THRESHOLD, PRICE * TAX_RATE, PRICE)";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("IF(PRICE>THRESHOLD,PRICE*SALES_TAX,PRICE)");
    });

    test("should handle formulas with no named expression references", () => {
      const formula = "A1 + B1 * 2";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("A1+B1*2");
    });

    test("should handle formulas with different named expressions", () => {
      const formula = "DISCOUNT_RATE * 100";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("DISCOUNT_RATE*100");
    });

    test("should handle named expressions in function calls", () => {
      const formula = "SUM(A1:A10) * TAX_RATE";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("SUM(A1:A10)*SALES_TAX");
    });

    test("should handle named expressions with sheet scope", () => {
      const formula = "Sheet1!TAX_RATE + Sheet2!DISCOUNT_RATE";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("Sheet1!SALES_TAX+Sheet2!DISCOUNT_RATE");
    });

    test("should handle mixed cell references and named expressions", () => {
      const formula = "A1 * TAX_RATE + B1 * DISCOUNT_RATE";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("A1*SALES_TAX+B1*DISCOUNT_RATE");
    });

    test("should handle named expressions in array formulas", () => {
      const formula = "{TAX_RATE; DISCOUNT_RATE; TAX_RATE}";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("{SALES_TAX;DISCOUNT_RATE;SALES_TAX}");
    });

    test("should return original formula if parsing fails", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = renameNamedExpressionInFormula(invalidFormula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe(invalidFormula);
    });

    test("should handle case-sensitive named expression names", () => {
      const formula = "tax_rate + TAX_RATE";
      const result = renameNamedExpressionInFormula(formula, "TAX_RATE", "SALES_TAX");
      
      expect(result).toBe("tax_rate+SALES_TAX");
    });
  });

  describe("formulaReferencesNamedExpression", () => {
    test("should return true for formula with named expression reference", () => {
      const formula = "TAX_RATE * 100";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(true);
    });

    test("should return false for formula without named expression reference", () => {
      const formula = "A1 + B1";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(false);
    });

    test("should return false for formula with different named expression", () => {
      const formula = "DISCOUNT_RATE * 100";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(false);
    });

    test("should return true for formula with multiple references to same named expression", () => {
      const formula = "TAX_RATE + TAX_RATE * 0.1";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(true);
    });

    test("should return true for formula with sheet-scoped named expression", () => {
      const formula = "Sheet1!TAX_RATE * 1.1";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(true);
    });

    test("should return true for named expression in function call", () => {
      const formula = "SUM(A1:A10, TAX_RATE)";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(true);
    });

    test("should return true for named expression in array", () => {
      const formula = "{TAX_RATE; 0.08; 0.1}";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(true);
    });

    test("should return false for invalid formula", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = formulaReferencesNamedExpression(invalidFormula, "TAX_RATE");
      
      expect(result).toBe(false);
    });

    test("should be case-sensitive", () => {
      const formula = "tax_rate * 100";
      const result = formulaReferencesNamedExpression(formula, "TAX_RATE");
      
      expect(result).toBe(false);
    });
  });

  describe("getReferencedNamedExpressionNames", () => {
    test("should return single named expression name", () => {
      const formula = "TAX_RATE * 100";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result).toEqual(["TAX_RATE"]);
    });

    test("should return multiple unique named expression names", () => {
      const formula = "TAX_RATE + DISCOUNT_RATE";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "TAX_RATE"]);
    });

    test("should return unique named expression names (no duplicates)", () => {
      const formula = "TAX_RATE + TAX_RATE * DISCOUNT_RATE";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "TAX_RATE"]);
    });

    test("should return empty array for formula with no named expression references", () => {
      const formula = "A1 + B1 * 2";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result).toEqual([]);
    });

    test("should handle complex nested formulas", () => {
      const formula = "IF(PRICE > THRESHOLD, PRICE * TAX_RATE, DISCOUNT_RATE)";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "PRICE", "TAX_RATE", "THRESHOLD"]);
    });

    test("should return empty array for invalid formula", () => {
      const invalidFormula = "INVALID(SYNTAX";
      const result = getReferencedNamedExpressionNames(invalidFormula);
      
      expect(result).toEqual([]);
    });

    test("should handle sheet-scoped named expressions", () => {
      const formula = "Sheet1!TAX_RATE + Sheet2!DISCOUNT_RATE";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "TAX_RATE"]);
    });

    test("should handle named expressions in arrays", () => {
      const formula = "{TAX_RATE; DISCOUNT_RATE; PRICE}";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "PRICE", "TAX_RATE"]);
    });

    test("should handle named expressions in function calls", () => {
      const formula = "SUM(TAX_RATE, DISCOUNT_RATE) + AVERAGE(PRICE, THRESHOLD)";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "PRICE", "TAX_RATE", "THRESHOLD"]);
    });

    test("should handle mixed cell references and named expressions", () => {
      const formula = "A1 * TAX_RATE + B1:B10 + DISCOUNT_RATE";
      const result = getReferencedNamedExpressionNames(formula);
      
      expect(result.sort()).toEqual(["DISCOUNT_RATE", "TAX_RATE"]);
    });
  });
});
