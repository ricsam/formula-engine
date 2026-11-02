import { describe, expect, test } from "bun:test";
import { FormulaError } from "../../../core/types";
import type { FunctionNode } from "../../../parser/ast";
import { astToString } from "../../../parser/formatter";
import { parseFormula } from "../../../parser/parser";
import {
  expandLet,
  LetExpansionError,
  MAX_LET_EXPANSION_NODES,
} from "./let-expansion";

function letCall(formula: string): FunctionNode {
  const ast = parseFormula(formula);
  if (ast.type !== "function" || ast.name !== "LET") {
    throw new Error(`${formula} is not a LET call`);
  }
  return ast;
}

/** The expanded formula text, checked to be a stable, re-parseable cache key. */
function expand(formula: string): string {
  const text = astToString(expandLet(letCall(formula)));
  expect(astToString(parseFormula(text))).toBe(text);
  return text;
}

function expansionError(formula: string): LetExpansionError {
  try {
    expandLet(letCall(formula));
  } catch (error) {
    if (error instanceof LetExpansionError) return error;
    throw error;
  }
  throw new Error(`${formula} expanded without an error`);
}

describe("expandLet", () => {
  test("replaces names with the expressions bound to them", () => {
    expect(expand("LET(x, A1, x*2)")).toBe("A1*2");
    expect(expand("LET(r, A1:B3, SUM(r, 1))")).toBe("SUM(A1:B3,1)");
    expect(expand("LET(a, 1, b, a+1, c, b*a, c)")).toBe("(1+1)*1");
    expect(expand("LET(x, 5, {x, x+1})")).toBe("{5,5+1}");
  });

  test("keeps operator precedence of inserted expressions", () => {
    expect(expand("LET(x, A1+1, x*2)")).toBe("(A1+1)*2");
    expect(expand("LET(x, 2^3, -x)")).toBe("-(2^3)");
    expect(expand("LET(x, -A1, x^2)")).toBe("-A1^2");
    expect(expand("LET(x, A1-B1, 10-x)")).toBe("10-(A1-B1)");
    expect(expand("LET(x, A1+1, x%)")).toBe("(A1+1)%");
  });

  test("matches names without regard to case", () => {
    expect(expand("LET(Total, A1, TOTAL+total)")).toBe("A1+A1");
  });

  test("drops unused names", () => {
    expect(expand("LET(unused, CEILING(5), x, 2, x)")).toBe("2");
  });

  test("leaves defined names, qualified names and structured references alone", () => {
    expect(expand("LET(x, 1, TestSheet!x + x + y)")).toBe("TestSheet!x+1+y");
    expect(expand("LET(t, Table1[Col], SUM(t, Table1[Col]))")).toBe(
      "SUM(Table1[Col],Table1[Col])"
    );
  });

  test("names can shadow function names without affecting calls", () => {
    expect(expand("LET(SUM, 4, SUM(SUM))")).toBe("SUM(4)");
  });

  test("nested LETs shadow outer names", () => {
    expect(expand("LET(x, 1, LET(x, x+1, x*10))")).toBe("(1+1)*10");
    expect(expand("LET(x, 5, LET(x, 7, x*2) + x)")).toBe("7*2+5");
  });

  test("is hygienic: inserted expressions are never captured", () => {
    // x is bound to the defined name y; the inner y must not capture it.
    expect(expand("LET(x, y, LET(y, 2, x+y))")).toBe("y+2");
    expect(expand("LET(x, y, y, 1, x+y)")).toBe("y+1");
    expect(expand("LET(a, b, b, a, b)")).toBe("b");
  });

  test("an omitted value is 0", () => {
    expect(expand("LET(x, , x+1)")).toBe("0+1");
  });

  test("shares bound expressions instead of copying them", () => {
    const pairs = 40;
    const parts = ["n_0", "A1"];
    for (let index = 1; index < pairs; index++) {
      parts.push(`n_${index}`, `n_${index - 1}+n_${index - 1}`);
    }
    // The last name expands to 2^40 nodes but is never used. Expansion shares
    // bound expressions, so it stays proportional to the written formula; an
    // expansion that copied them would not finish.
    expect(expand(`LET(${parts.join(", ")}, n_2)`)).toBe("A1+A1+(A1+A1)");
  });

  describe("errors", () => {
    test.each([
      ["LET()", "but got 0"],
      ["LET(x)", "but got 1"],
      ["LET(x, 1)", "but got 2"],
      ["LET(x, 1, y, 2)", "but got 4"],
    ])("%s has the wrong argument count", (formula, detail) => {
      const error = expansionError(formula);
      expect(error.formulaError).toBe(FormulaError.VALUE);
      expect(error.message).toContain(detail);
    });

    test("a missing calculation is #VALUE!", () => {
      const error = expansionError("LET(x, 1, )");
      expect(error.formulaError).toBe(FormulaError.VALUE);
      expect(error.message).toContain("final calculation");
    });

    test.each([
      ["LET(x1, 1, 2)", "LET name 1 (X1) is a cell reference, not a name"],
      ["LET(a, 1, b2, 2, 3)", "LET name 2 (B2) is a cell reference, not a name"],
      [
        "LET(Sheet1!x, 1, 2)",
        "LET name 1 (Sheet1!x) must not be qualified by a sheet or workbook",
      ],
      ["LET(1, 1, 2)", "LET name 1 (1) is not a valid name"],
      ["LET(TRUE, 1, 2)", "LET name 1 (TRUE) is not a valid name"],
      ["LET(, 1, 2)", "LET name 1 is missing"],
      ["LET(x, 1, X, 2, 3)", "LET declares the name X more than once"],
    ])("%s is #NAME?", (formula, message) => {
      const error = expansionError(formula);
      expect(error.formulaError).toBe(FormulaError.NAME);
      expect(error.message).toBe(message);
    });

    test("a nested LET is validated with its enclosing LET", () => {
      expect(expansionError("LET(a, 1, LET(b, 2))").formulaError).toBe(
        FormulaError.VALUE
      );
    });

    test("an expansion past the size limit is #CALC!", () => {
      const pairs = Math.ceil(Math.log2(MAX_LET_EXPANSION_NODES + 1)) + 1;
      const parts = ["n_0", "A1"];
      for (let index = 1; index < pairs; index++) {
        parts.push(`n_${index}`, `n_${index - 1}+n_${index - 1}`);
      }
      const error = expansionError(`LET(${parts.join(", ")}, n_${pairs - 1})`);
      expect(error.formulaError).toBe(FormulaError.CALC);
      expect(error.message).toContain(
        MAX_LET_EXPANSION_NODES.toLocaleString("en-US")
      );
    });
  });
});
