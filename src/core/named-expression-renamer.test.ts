import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "./engine";
import {
  formulaReferencesNamedExpression,
  getReferencedNamedExpressionNames,
  renameNamedExpressionInFormula,
} from "./named-expression-renamer";
import { parseCellReference } from "./utils";

// General renaming behavior is covered in
// tests/unit/core/named-expression-renamer.test.ts; these cases cover LET.
describe("named expression renaming with LET names", () => {
  describe("renameNamedExpressionInFormula", () => {
    test("renames defined-name references inside LET values and calculations", () => {
      expect(
        renameNamedExpressionInFormula("LET(x, Rate, x * Rate)", "Rate", "Tax")
      ).toBe("LET(x,Tax,x*Tax)");
    });

    test("leaves LET names spelled like the renamed name alone", () => {
      expect(
        renameNamedExpressionInFormula("LET(y, 5, y + 1)", "y", "z")
      ).toBe("LET(y,5,y+1)");
      // LET names match without regard to case, so `Y` is the local too.
      expect(
        renameNamedExpressionInFormula("LET(y, 5, Y + 1)", "Y", "z")
      ).toBe("LET(y,5,Y+1)");
    });

    test("renames the defined name where no LET name hides it", () => {
      expect(
        renameNamedExpressionInFormula("LET(y, y * 2, y + 1) + y", "y", "z")
      ).toBe("LET(y,z*2,y+1)+z");
      expect(
        renameNamedExpressionInFormula("LET(a, y, LET(y, 2, a + y))", "y", "z")
      ).toBe("LET(a,z,LET(y,2,a+y))");
    });

    test("renames a LET name that would capture the renamed reference", () => {
      expect(
        renameNamedExpressionInFormula("LET(x, 1, x + y)", "y", "x")
      ).toBe("LET(x_1,1,x_1+x)");
      // The capture is case-insensitive like LET names.
      expect(
        renameNamedExpressionInFormula("LET(Total, 1, Total + y)", "y", "TOTAL")
      ).toBe("LET(Total_1,1,Total_1+TOTAL)");
    });

    test("renames every LET name that would capture the reference", () => {
      // Renaming only the inner x would let the outer x capture the reference.
      expect(
        renameNamedExpressionInFormula(
          "LET(x, 1, LET(x, 2, x + y) + x)",
          "y",
          "x"
        )
      ).toBe("LET(x_2,1,LET(x_1,2,x_1+x)+x_2)");
    });

    test("only renames LET names in scope at the reference", () => {
      expect(
        renameNamedExpressionInFormula("LET(x, y, x * 2)", "y", "x")
      ).toBe("LET(x,x,x*2)");
      expect(
        renameNamedExpressionInFormula("LET(x, 1, x) + y", "y", "x")
      ).toBe("LET(x,1,x)+x");
    });

    test("picks a fresh LET name that the formula does not use", () => {
      expect(
        renameNamedExpressionInFormula(
          "LET(x, 1, x_1, 2, x + x_1 + y + X_2)",
          "y",
          "x"
        )
      ).toBe("LET(x_3,1,x_1,2,x_3+x_1+x+X_2)");
    });

    test("the renamed formula evaluates like the original", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook({ workbookName: "Book" });
      engine.addSheet({ workbookName: "Book", sheetName: "Sheet1" });
      const cell = (ref: string) => ({
        workbookName: "Book",
        sheetName: "Sheet1",
        ...parseCellReference(ref),
      });
      engine.addNamedExpression({ expressionName: "y", expression: "100" });
      engine.setCellContent(cell("A1"), "=LET(x, 1, LET(x, 2, x + y) + x)");
      engine.setCellContent(cell("A2"), "=LET(y, 5, y + 1)");
      engine.setCellContent(cell("A3"), "=LET(a, y, LET(y, 2, a + y))");
      const read = () => ["A1", "A2", "A3"].map((ref) => engine.getCellValue(cell(ref)));
      expect(read()).toEqual([103, 6, 102]);

      engine.renameNamedExpression({ expressionName: "y", newName: "x" });

      const content = engine.getSheetSerialized({
        workbookName: "Book",
        sheetName: "Sheet1",
      });
      expect(content.get("A1")).toBe("=LET(x_2, 1, LET(x_1, 2, x_1+x)+x_2)");
      expect(content.get("A2")).toBe("=LET(y, 5, y+1)");
      expect(content.get("A3")).toBe("=LET(a, x, LET(y, 2, a+y))");
      expect(read()).toEqual([103, 6, 102]);
    });

    test("named expression formulas are renamed the same way", () => {
      const engine = FormulaEngine.buildEmpty();
      engine.addWorkbook({ workbookName: "Book" });
      engine.addSheet({ workbookName: "Book", sheetName: "Sheet1" });
      const a1 = { workbookName: "Book", sheetName: "Sheet1", colIndex: 0, rowIndex: 0 };
      engine.addNamedExpression({ expressionName: "y", expression: "100" });
      engine.addNamedExpression({
        expressionName: "Total",
        expression: "LET(x, 1, x + y)",
      });
      engine.setCellContent(a1, "=Total");
      expect(engine.getCellValue(a1)).toBe(101);

      engine.renameNamedExpression({ expressionName: "y", newName: "x" });

      expect(
        engine._namedExpressionManager.getNamedExpression({
          name: "Total",
          scope: { type: "global" },
        })?.expression
      ).toBe("LET(x_1,1,x_1+x)");
      expect(engine.getCellValue(a1)).toBe(101);
    });
  });

  describe("finding references", () => {
    test("LET names are not references to defined names", () => {
      expect(getReferencedNamedExpressionNames("LET(x, 1, x + 1)")).toEqual([]);
      expect(
        getReferencedNamedExpressionNames("LET(x, y, LET(y, 2, x + y + z))").sort()
      ).toEqual(["y", "z"]);
      expect(
        getReferencedNamedExpressionNames("LET(x, 1, Sheet1!x + x)")
      ).toEqual(["x"]);
    });

    test("formulaReferencesNamedExpression ignores LET names", () => {
      expect(formulaReferencesNamedExpression("LET(y, 5, y + 1)", "y")).toBe(false);
      expect(formulaReferencesNamedExpression("LET(y, y, y + 1)", "y")).toBe(true);
      expect(formulaReferencesNamedExpression("LET(x, 1, x) + y", "y")).toBe(true);
    });
  });
});
