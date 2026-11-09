import { describe, expect, test } from "bun:test";
import type { ASTNode } from "../parser/ast";
import { ParseError, parseFormula } from "../parser/parser";
import {
  formatFormula,
  tryFormatFormula,
  type FormulaFormatOptions,
} from "./formula-formatter";

function formulaBody(formula: string): string {
  return formula.startsWith("=") ? formula.slice(1) : formula;
}

/** The AST without the source positions of its nodes, for comparing trees. */
function withoutPositions(ast: ASTNode): unknown {
  return JSON.parse(
    JSON.stringify(ast, (key, value) => (key === "position" ? undefined : value))
  );
}

/**
 * Compares the parsed trees rather than their `astToString` text, so a
 * rendering that loses part of a node cannot make two formulas look equal.
 */
function expectSameExpression(left: string, right: string): void {
  expect(withoutPositions(parseFormula(formulaBody(left)))).toEqual(
    withoutPositions(parseFormula(formulaBody(right)))
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
      '=IF(OR([@pep_concentration]="np",[@pep_concentration]="null",[@pep_concentration]="MAP"),[@gate],TEXTJOIN(", ",TRUE,[@[pep_concentration]:[gate]]))',
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

  test("keeps every column of a current-row column span", () => {
    // The engine stores this output, so a span written as its first column
    // would make TEXTJOIN read one cell instead of the row's span.
    expect(
      formatFormula('=TEXTJOIN(", ",TRUE,[@[pep_concentration]:[gate]])')
    ).toBe('=TEXTJOIN(", ", TRUE, [@[pep_concentration]:[gate]])');
    expect(formatFormula("=SUM( [@[Net Sales]:[Gross Profit]] )")).toBe(
      "=SUM([@[Net Sales]:[Gross Profit]])"
    );
  });

  test("returns invalid formulas verbatim from the safe API", () => {
    const invalidFormulas = [
      "=SUM(A1;B1)",
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
    const result = tryFormatFormula("=SUM(A1;B1)");
    expect(result).toMatchObject({
      ok: false,
      formula: "=SUM(A1;B1)",
      changed: false,
      error: {
        name: "ParseError",
        position: { start: 7, end: 8 },
      },
    });
  });

  test("the strict API refuses invalid syntax", () => {
    expect(() => formatFormula("=SUM(A1;B1)")).toThrow(ParseError);
  });

  test("prints omitted arguments as nothing", () => {
    expect(formatFormula("=SEQUENCE(,3)")).toBe("=SEQUENCE(, 3)");
    expect(formatFormula("=IF(A1,1,)")).toBe("=IF(A1, 1,)");
    expect(formatFormula("=IF( A1 , , B1 )")).toBe("=IF(A1,, B1)");
    expect(formatFormula("=IF(,,)")).toBe("=IF(,,)");

    expect(formatFormula("=SEQUENCE(,3)", { style: "pretty" })).toBe(
      "=SEQUENCE(\n  ,\n  3\n)"
    );
    // An omitted last argument is marked by the preceding comma alone.
    expect(formatFormula("=IF(A1,1,)", { style: "pretty" })).toBe(
      "=IF(\n  A1,\n  1,\n)"
    );
    expect(formatFormula("=IF(A1,,B1)", { style: "pretty" })).toBe(
      "=IF(\n  A1,\n  ,\n  B1\n)"
    );
    expect(formatFormula("=SUM(,)", { style: "pretty" })).toBe(
      "=SUM(\n  ,\n)"
    );
    expect(formatFormula("={}", { style: "pretty" })).toBe("={}");
  });

  test("omitted arguments round-trip in both styles", () => {
    const formulas = [
      "=SEQUENCE(,3)",
      "=IF(A1,,B1)",
      "=IF(A1,1,)",
      "=IF(,,)",
      "=SUM(,)",
      "=IF(A1, SEQUENCE(, 2), LET(x, , y, 1, x + y))",
      "={}",
    ];

    for (const formula of formulas) {
      const compact = formatFormula(formula, { style: "compact" });
      const pretty = formatFormula(compact, { style: "pretty" });

      expect(compact).not.toContain(", ,");
      expect(compact).not.toContain(", )");
      expect(pretty.split("\n").some((line) => line.trim() === "")).toBe(false);
      expect(formatFormula(compact, { style: "compact" })).toBe(compact);
      expect(formatFormula(pretty, { style: "pretty" })).toBe(pretty);
      expect(formatFormula(pretty, { style: "compact" })).toBe(compact);
      expectSameExpression(formula, compact);
      expectSameExpression(formula, pretty);
    }
  });

  test("formats LET calls in both styles", () => {
    expect(formatFormula("=LET(x,A1,y,x*2,x+y)")).toBe(
      "=LET(x, A1, y, x*2, x+y)"
    );
    expect(formatFormula("=LET(x,A1,y,x*2,x+y)", { style: "pretty" })).toBe(`=LET(
  x,
  A1,
  y,
  x * 2,
  x + y
)`);
  });

  test("LET formulas round-trip in both styles", () => {
    const formulas = [
      "=LET(x, A1, x*2)",
      "=LET(total, SUM(A1:A10), count, COUNT(A1:A10), IF(count > 0, total / count, 0))",
      "=LET(x, 1, LET(x, x+1, y, -x^2, {x, y; x*y, 0}))",
      "=LET(a, , b, Sheet2!Rate, a + b)",
      "=LET(a, UNIQUE(TOCOL(A2:A1000, 1)), b, UNIQUE(TOCOL(B2:B1000, 1)), HSTACK(TOCOL(IF(SEQUENCE(,ROWS(b)), a)), TOCOL(IF(SEQUENCE(ROWS(a)), TOROW(b)))))",
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
