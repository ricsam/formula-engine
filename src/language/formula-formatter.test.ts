import { describe, expect, test } from "bun:test";
import { astToString } from "../parser/formatter";
import { ParseError, parseFormula } from "../parser/parser";
import {
  formatFormula,
  tryFormatFormula,
  type FormulaFormatOptions,
} from "./formula-formatter";

function formulaBody(formula: string): string {
  return formula.startsWith("=") ? formula.slice(1) : formula;
}

function expectSameExpression(left: string, right: string): void {
  expect(astToString(parseFormula(formulaBody(left)))).toBe(
    astToString(parseFormula(formulaBody(right)))
  );
}

describe("public formula formatter", () => {
  test("compacts formulas to one readable line by default", () => {
    expect(formatFormula("= IF( A1 > 0,\n SUM(A1:A10, B1:B10), 0 )")).toBe(
      "=IF(A1>0, SUM(A1:A10, B1:B10), 0)"
    );
  });

  test("formats formula bodies without adding a formula marker", () => {
    expect(formatFormula("SUM(A1,B1)", { style: "compact" })).toBe(
      "SUM(A1, B1)"
    );
  });

  test("pretty-prints nested calls and binary operators", () => {
    expect(
      formatFormula("=IF(A1>0,SUM(A1:A10,B1:B10),0)", {
        style: "pretty",
      })
    ).toBe(`=IF(
  A1 > 0,
  SUM(
    A1:A10,
    B1:B10
  ),
  0
)`);
  });

  test("supports a caller-defined indentation unit", () => {
    expect(
      formatFormula("SUM(A1,MAX(B1,C1))", {
        style: "pretty",
        indent: "\t",
      })
    ).toBe("SUM(\n\tA1,\n\tMAX(\n\t\tB1,\n\t\tC1\n\t)\n)");
  });

  test("formats arrays compactly and across rows in pretty mode", () => {
    expect(formatFormula("={1,2;3,4}")).toBe("={1, 2; 3, 4}");
    expect(formatFormula("={1,2;3,4}", { style: "pretty" })).toBe(
      "={\n  1, 2;\n  3, 4\n}"
    );
  });

  test("keeps parentheses required by precedence and unary operators", () => {
    const cases = [
      "=(A1+B1)*C1",
      "=A1-(B1-C1)",
      "=(A1+B1)%",
      "=-(A1+B1)",
      "=A1^(B1^C1)",
    ];

    for (const formula of cases) {
      const compact = formatFormula(formula, { style: "compact" });
      const pretty = formatFormula(formula, { style: "pretty" });
      expectSameExpression(formula, compact);
      expectSameExpression(formula, pretty);
    }
  });

  test("preserves strings and quoted sheet names while normalizing layout", () => {
    const input = `=IF('John''s Sheet'!A1="a,b","x, y","z")`;
    const compact = formatFormula(input, { style: "compact" });

    expect(compact).toBe(`=IF('John''s Sheet'!A1="a,b", "x, y", "z")`);
    expectSameExpression(input, compact);
  });

  test("compact and pretty formatting are stable and semantics-preserving", () => {
    const formulas = [
      "=SUM([External]Data!A1:D,Table1[[#Data],[Revenue]:[Profit]])",
      '=IFERROR(INDEX(Table1[Value],MATCH(A1,Table1[Key],0)),"missing")',
      "=Sheet1:Sheet3!A1+B1*C1",
      "={SUM(A1,A2),MAX(B1,B2);TRUE,#N/A}",
    ];

    for (const formula of formulas) {
      const compact = formatFormula(formula, { style: "compact" });
      const pretty = formatFormula(compact, { style: "pretty" });

      expect(formatFormula(compact, { style: "compact" })).toBe(compact);
      expect(formatFormula(pretty, { style: "pretty" })).toBe(pretty);
      expect(formatFormula(pretty, { style: "compact" })).toBe(compact);
      expectSameExpression(formula, compact);
      expectSameExpression(formula, pretty);
    }
  });

  test("returns invalid formulas verbatim from the safe API", () => {
    const invalidFormulas = [
      "=SUM(A1,,B1)",
      "=A1+",
      '="unterminated',
      '="',
      '="escaped quote""',
      "~",
      "=",
    ];

    for (const formula of invalidFormulas) {
      const result = tryFormatFormula(formula, { style: "pretty" });
      expect(result.ok).toBe(false);
      expect(result.formula).toBe(formula);
      expect(result.changed).toBe(false);
      if (!result.ok) {
        expect(result.error.message.length).toBeGreaterThan(0);
      }
    }
  });

  test("reports parser positions against the original leading-equals input", () => {
    const result = tryFormatFormula("=SUM(A1,,B1)");
    expect(result).toMatchObject({
      ok: false,
      formula: "=SUM(A1,,B1)",
      changed: false,
      error: {
        name: "ParseError",
        position: { start: 8, end: 9 },
      },
    });
  });

  test("the strict API refuses invalid syntax", () => {
    expect(() => formatFormula("=SUM(A1,,B1)")).toThrow(ParseError);
  });

  test("reports whether valid formatting changed the text", () => {
    expect(tryFormatFormula("=SUM(A1, B1)")).toEqual({
      ok: true,
      formula: "=SUM(A1, B1)",
      changed: false,
    });
    expect(tryFormatFormula("=SUM( A1,B1 )")).toEqual({
      ok: true,
      formula: "=SUM(A1, B1)",
      changed: true,
    });
  });

  test("safe formatting also contains invalid option errors", () => {
    const options = { style: "wide" } as unknown as FormulaFormatOptions;
    const result = tryFormatFormula("=A1", options);
    expect(result).toMatchObject({
      ok: false,
      formula: "=A1",
      changed: false,
      error: { name: "TypeError" },
    });
  });
});
