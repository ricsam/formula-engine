import { describe, expect, test } from "bun:test";

import { FormulaEngine } from "../core/engine";
import {
  analyzeFormula,
  findFormulaReferenceAt,
  type FormulaAnalysis,
  type FormulaTokenKind,
  type FormulaReferenceTarget,
} from "./formula-analysis";

const origin = {
  workbookName: "Book",
  sheetName: "Sheet1",
  colIndex: 3,
  rowIndex: 9,
};

function onlyTarget(
  analysis: ReturnType<typeof analyzeFormula>
): FormulaReferenceTarget {
  const resolution = analysis.references[0]?.resolution;
  expect(resolution?.status).toBe("resolved");
  if (resolution?.status !== "resolved") {
    throw new Error("Expected a resolved reference");
  }
  expect(resolution.targets).toHaveLength(1);
  return resolution.targets[0]!;
}

function buildEngine(): FormulaEngine {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName: "Book" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet1" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet2" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet3" });
  return engine;
}

/** Each identifier-like token as `text:kind`, in order. */
function nameTokens(analysis: FormulaAnalysis): string[] {
  const kinds: FormulaTokenKind[] = ["variable", "named-expression", "cell-reference"];
  return analysis.tokens
    .filter((token) => kinds.includes(token.kind))
    .map(
      (token) =>
        `${analysis.formula.slice(token.span.start, token.span.end)}:${token.kind}`
    );
}

function diagnosticCodes(analysis: FormulaAnalysis): string[] {
  return analysis.diagnostics.map((diagnostic) => diagnostic.code);
}

describe("analyzeFormula", () => {
  test("returns semantic tokens and original-string spans including leading =", () => {
    const formula = "=SUM(A2, Sheet2!$B$3:C4)";
    const analysis = analyzeFormula({ formula, origin });

    expect(analysis.formula).toBe(formula);
    expect(analysis.tokens[0]).toEqual({
      span: { start: 0, end: 1 },
      kind: "formula-marker",
      modifiers: [],
    });
    expect(
      analysis.tokens.find((token) => token.kind === "function")?.span
    ).toEqual({ start: 1, end: 4 });
    expect(
      analysis.references.map(({ kind, span }) => ({ kind, span }))
    ).toEqual([
      { kind: "cell", span: { start: 5, end: 7 } },
      { kind: "range", span: { start: 9, end: 23 } },
    ]);

    const first = analysis.references[0]!;
    expect(
      analysis.tokens
        .filter((token) => token.referenceId === first.id)
        .map((token) => formula.slice(token.span.start, token.span.end))
    ).toEqual(["A2"]);
    expect(findFormulaReferenceAt(analysis, 5)).toBe(first);
    // A caret immediately after a reference retains the left reference.
    expect(findFormulaReferenceAt(analysis, 7)).toBe(first);
  });

  test("uses the same offsets for formulas without a leading =", () => {
    const analysis = analyzeFormula({ formula: "A2+1", origin });
    expect(
      analysis.tokens.some((token) => token.kind === "formula-marker")
    ).toBe(false);
    expect(analysis.references[0]?.span).toEqual({ start: 0, end: 2 });
  });

  test("uses UTF-16 offsets expected by browser editors", () => {
    const formula = '="😀"+A2';
    const analysis = analyzeFormula({ formula, origin });
    const span = analysis.references[0]?.span;
    expect(span).toEqual({
      start: formula.indexOf("A2"),
      end: formula.indexOf("A2") + 2,
    });
  });

  test("never throws for incomplete input and preserves references parsed so far", () => {
    for (const formula of ["=", "=SUM(", "=SUM(A2,", "='unterminated", "=@"]) {
      expect(() => analyzeFormula({ formula, origin })).not.toThrow();
      expect(
        analyzeFormula({ formula, origin }).diagnostics.length
      ).toBeGreaterThan(0);
    }

    const partial = analyzeFormula({ formula: "=SUM(A2,", origin });
    expect(partial.references[0]?.span).toEqual({ start: 5, end: 7 });
    expect(partial.references[0]?.resolution.status).toBe("resolved");
  });

  test("reports missing origin instead of throwing", () => {
    const analysis = analyzeFormula({ formula: "=A2" });
    expect(analysis.references[0]?.resolution).toEqual({
      status: "unresolved",
      reason: "missing-origin",
    });
    expect(analysis.diagnostics).toContainEqual({
      span: { start: 1, end: 3 },
      severity: "warning",
      code: "reference.missing-origin",
      message: "An origin cell is required to resolve this reference.",
    });
  });

  test("normalizes reversed finite and open-ended ranges", () => {
    expect(onlyTarget(analyzeFormula({ formula: "=B3:A1", origin }))).toEqual({
      type: "range",
      address: {
        workbookName: "Book",
        sheetName: "Sheet1",
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "number", value: 2 },
          },
        },
      },
    });
    expect(onlyTarget(analyzeFormula({ formula: "=D5:B", origin }))).toEqual({
      type: "range",
      address: {
        workbookName: "Book",
        sheetName: "Sheet1",
        range: {
          start: { col: 1, row: 4 },
          end: {
            col: { type: "number", value: 3 },
            row: { type: "infinity", sign: "positive" },
          },
        },
      },
    });
  });
});

describe("FormulaEngine.analyzeFormula", () => {
  test("validates workbook and sheet scope and resolves cell/range targets", () => {
    const engine = buildEngine();
    const range = engine.analyzeFormula({
      formula: "=Sheet2!B3:C4",
      origin,
    });
    expect(onlyTarget(range)).toEqual({
      type: "range",
      address: {
        workbookName: "Book",
        sheetName: "Sheet2",
        range: {
          start: { col: 1, row: 2 },
          end: {
            col: { type: "number", value: 2 },
            row: { type: "number", value: 3 },
          },
        },
      },
    });

    expect(
      engine.analyzeFormula({ formula: "=Missing!A1", origin }).references[0]
        ?.resolution
    ).toEqual({ status: "unresolved", reason: "missing-sheet" });
    expect(
      engine.analyzeFormula({
        formula: "=[Missing]Sheet1!A1",
        origin,
      }).references[0]?.resolution
    ).toEqual({ status: "unresolved", reason: "missing-workbook" });
  });

  test("expands a 3D reference into directly highlightable sheet targets", () => {
    const engine = buildEngine();
    const analysis = engine.analyzeFormula({
      formula: "=Sheet1:Sheet3!A2",
      origin,
    });
    const resolution = analysis.references[0]?.resolution;
    expect(resolution?.status).toBe("resolved");
    if (resolution?.status !== "resolved") return;
    expect(
      resolution.targets.map((target) =>
        target.type === "cell" ? target.address.sheetName : undefined
      )
    ).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
  });

  test("resolves named expressions using sheet, workbook, then global scope", () => {
    const engine = buildEngine();
    engine.addNamedExpression({ expressionName: "Rate", expression: "1" });
    engine.addNamedExpression({
      expressionName: "Rate",
      expression: "2",
      workbookName: "Book",
    });
    engine.addNamedExpression({
      expressionName: "Rate",
      expression: "3",
      workbookName: "Book",
      sheetName: "Sheet1",
    });

    expect(
      onlyTarget(engine.analyzeFormula({ formula: "=Rate", origin }))
    ).toEqual({
      type: "named-expression",
      name: "Rate",
      scope: {
        type: "sheet",
        workbookName: "Book",
        sheetName: "Sheet1",
      },
    });
    expect(
      onlyTarget(engine.analyzeFormula({ formula: "=Sheet2!Rate", origin }))
    ).toEqual({
      type: "named-expression",
      name: "Rate",
      scope: { type: "workbook", workbookName: "Book" },
    });
  });

  test("resolves table columns and current-row references to display ranges", () => {
    const engine = buildEngine();
    engine.setCellContent(
      { workbookName: "Book", sheetName: "Sheet1", colIndex: 0, rowIndex: 0 },
      "Amount"
    );
    engine.setCellContent(
      { workbookName: "Book", sheetName: "Sheet1", colIndex: 1, rowIndex: 0 },
      "Count"
    );
    engine.addTable({
      tableName: "Sales",
      workbookName: "Book",
      sheetName: "Sheet1",
      start: "A1",
      numRows: { type: "number", value: 4 },
      numCols: 2,
    });

    expect(
      onlyTarget(engine.analyzeFormula({ formula: "=Sales[Amount]", origin }))
    ).toEqual({
      type: "table",
      workbookName: "Book",
      sheetName: "Sheet1",
      tableName: "Sales",
      range: {
        start: { col: 0, row: 1 },
        end: {
          col: { type: "number", value: 0 },
          row: { type: "number", value: 4 },
        },
      },
    });

    const currentOrigin = { ...origin, colIndex: 0, rowIndex: 2 };
    expect(
      onlyTarget(
        engine.analyzeFormula({ formula: "=[@Amount]", origin: currentOrigin })
      )
    ).toEqual({
      type: "table",
      workbookName: "Book",
      sheetName: "Sheet1",
      tableName: "Sales",
      range: {
        start: { col: 0, row: 2 },
        end: {
          col: { type: "number", value: 0 },
          row: { type: "number", value: 2 },
        },
      },
    });

    expect(
      engine.analyzeFormula({ formula: "=Sales[Missing]", origin })
        .references[0]?.resolution
    ).toEqual({ status: "unresolved", reason: "missing-column" });
  });

  test("treats function-like and lexical-special text as structured column names", () => {
    const engine = buildEngine();
    engine.setCellContent(
      { workbookName: "Book", sheetName: "Sheet1", colIndex: 0, rowIndex: 0 },
      "Cell number needed (M) with 20% excess"
    );
    engine.setCellContent(
      { workbookName: "Book", sheetName: "Sheet1", colIndex: 1, rowIndex: 0 },
      "Question?"
    );
    engine.addTable({
      tableName: "eAPCHarvest",
      workbookName: "Book",
      sheetName: "Sheet1",
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 2,
    });

    for (const column of [
      "Cell number needed (M) with 20% excess",
      "Question?",
    ]) {
      const formula = `=eAPCHarvest[${column}]`;
      const analysis = engine.analyzeFormula({ formula, origin });

      expect(analysis.references[0]?.resolution.status).toBe("resolved");
      expect(
        analysis.diagnostics.filter(
          (diagnostic) =>
            diagnostic.code === "function.unknown" ||
            diagnostic.code === "syntax.invalid-token"
        )
      ).toEqual([]);
      expect(
        analysis.tokens
          .filter((token) => token.kind === "table-column")
          .map((token) => formula.slice(token.span.start, token.span.end))
      ).not.toHaveLength(0);
    }
  });

  test("still diagnoses genuine unknown functions and malformed syntax", () => {
    const engine = buildEngine();
    const unknownFunction = engine.analyzeFormula({
      formula: "=NEEDED(A1)",
      origin,
    });
    const malformedReference = engine.analyzeFormula({
      formula: "=eAPCHarvest[Cell number needed (M)",
      origin,
    });

    expect(unknownFunction.diagnostics).toContainEqual(
      expect.objectContaining({ code: "function.unknown" })
    );
    expect(malformedReference.diagnostics).toContainEqual(
      expect.objectContaining({ severity: "error" })
    );
    expect(malformedReference.diagnostics).not.toContainEqual(
      expect.objectContaining({ code: "function.unknown" })
    );
  });

  test("does not classify workbook qualifiers as functions", () => {
    const engine = buildEngine();
    engine.addWorkbook({ workbookName: "Book needed (M)" });
    engine.addSheet({
      workbookName: "Book needed (M)",
      sheetName: "Sheet1",
    });

    const analysis = engine.analyzeFormula({
      formula: "=[Book needed (M)]Sheet1!A1",
      origin,
    });

    expect(analysis.references[0]?.resolution.status).toBe("resolved");
    expect(analysis.diagnostics).not.toContainEqual(
      expect.objectContaining({ code: "function.unknown" })
    );
  });

  describe("tables in other workbooks", () => {
    function buildBatchEngine(): FormulaEngine {
      const engine = buildEngine();
      engine.addWorkbook({ workbookName: "Batch 1" });
      engine.addSheet({ workbookName: "Batch 1", sheetName: "Data" });
      engine.setCellContent(
        { workbookName: "Batch 1", sheetName: "Data", colIndex: 0, rowIndex: 0 },
        "Sample"
      );
      engine.setCellContent(
        { workbookName: "Batch 1", sheetName: "Data", colIndex: 1, rowIndex: 0 },
        "Value"
      );
      engine.addTable({
        tableName: "Results",
        workbookName: "Batch 1",
        sheetName: "Data",
        start: "A1",
        numRows: { type: "number", value: 3 },
        numCols: 2,
      });
      return engine;
    }

    /** Each token of the reference as `text:kind`, with its modifiers. */
    function referenceTokens(analysis: FormulaAnalysis): string[] {
      const reference = analysis.references[0]!;
      return analysis.tokens
        .filter((token) => token.referenceId === reference.id)
        .map(
          (token) =>
            `${analysis.formula.slice(token.span.start, token.span.end)}:${
              token.kind
            }:${token.modifiers.join("+")}`
        );
    }

    test("resolves [Workbook]Table[Column] against the other workbook", () => {
      const engine = buildBatchEngine();
      const formula = "=SUM([Batch 1]Results[Value])";
      const analysis = engine.analyzeFormula({ formula, origin });

      expect(onlyTarget(analysis)).toEqual({
        type: "table",
        workbookName: "Batch 1",
        sheetName: "Data",
        tableName: "Results",
        range: {
          start: { col: 1, row: 1 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "number", value: 3 },
          },
        },
      });
      expect(analysis.references[0]).toMatchObject({
        kind: "structured-reference",
        span: { start: 5, end: formula.length - 1 },
      });
      expect(referenceTokens(analysis)).toEqual([
        "[:delimiter:external",
        "Batch:workbook-name:external",
        "1:workbook-name:external",
        "]:delimiter:external",
        "Results:table-name:external",
        "[:delimiter:external",
        "Value:table-column:external",
        "]:delimiter:external",
      ]);
      expect(analysis.diagnostics).toEqual([]);
    });

    test("resolves current-row references at the origin row", () => {
      const engine = buildBatchEngine();
      const analysis = engine.analyzeFormula({
        formula: "=[Batch 1]Results[@[Sample]:[Value]]",
        origin: { ...origin, rowIndex: 2 },
      });

      expect(onlyTarget(analysis)).toEqual({
        type: "table",
        workbookName: "Batch 1",
        sheetName: "Data",
        tableName: "Results",
        range: {
          start: { col: 0, row: 2 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "number", value: 2 },
          },
        },
      });
    });

    test("reports missing workbooks, tables and columns", () => {
      const engine = buildBatchEngine();
      const resolution = (formula: string) =>
        engine.analyzeFormula({ formula, origin }).references[0]?.resolution;

      expect(resolution("=[Missing]Results[Value]")).toEqual({
        status: "unresolved",
        reason: "missing-workbook",
      });
      expect(resolution("=[Batch 1]Missing[Value]")).toEqual({
        status: "unresolved",
        reason: "missing-table",
      });
      expect(resolution("=[Batch 1]Results[Missing]")).toEqual({
        status: "unresolved",
        reason: "missing-column",
      });
      // Tables are workbook resources: Book has no Results table.
      expect(resolution("=Results[Value]")).toEqual({
        status: "unresolved",
        reason: "missing-table",
      });
    });

    test("diagnoses [Workbook]!Table[Column] without claiming a bare column", () => {
      const engine = buildBatchEngine();
      const analysis = engine.analyzeFormula({
        formula: "=[Batch 1]!Results[Value]",
        origin,
      });

      expect(analysis.diagnostics).toContainEqual(
        expect.objectContaining({
          code: "syntax.parse-error",
          message:
            "Cross-workbook table references use [Batch 1]Results[...] instead of [Batch 1]!Results[...]",
        })
      );
      expect(
        analysis.tokens
          .filter((token) => token.kind === "table-column")
          .map((token) => analysis.formula.slice(token.span.start, token.span.end))
      ).toEqual(["Value"]);
    });
  });
});

describe("LET names in formula analysis", () => {
  test("marks declarations and uses as variables without references or warnings", () => {
    const engine = buildEngine();
    const analysis = engine.analyzeFormula({
      formula: "=LET(total, A1, total*2)",
      origin,
    });

    expect(nameTokens(analysis)).toEqual([
      "total:variable",
      "A1:cell-reference",
      "total:variable",
    ]);
    expect(analysis.references.map((reference) => reference.kind)).toEqual([
      "cell",
    ]);
    expect(analysis.diagnostics).toEqual([]);
    const variable = analysis.tokens.find((token) => token.kind === "variable");
    expect(variable?.referenceId).toBeUndefined();
    expect(variable?.modifiers).toEqual([]);
  });

  test("LET names hide defined names only inside their scope", () => {
    const engine = buildEngine();
    engine.addNamedExpression({ expressionName: "Rate", expression: "0.5" });
    const analysis = engine.analyzeFormula({
      formula: "=LET(Rate, 2, RATE*10) + Rate + Sheet2!Rate",
      origin,
    });

    expect(nameTokens(analysis)).toEqual([
      "Rate:variable",
      "RATE:variable",
      "Rate:named-expression",
      "Rate:named-expression",
    ]);
    expect(
      analysis.references.map((reference) => [
        analysis.formula.slice(reference.span.start, reference.span.end),
        reference.resolution.status,
      ])
    ).toEqual([
      ["Rate", "resolved"],
      ["Sheet2!Rate", "resolved"],
    ]);
    expect(analysis.diagnostics).toEqual([]);
  });

  test("follows LET scoping: values see earlier names, not their own", () => {
    const analysis = analyzeFormula({
      formula: "=LET(a, a, b, a+1, LET(c, b, a+b+c)) + c",
      origin,
    });

    expect(nameTokens(analysis)).toEqual([
      "a:variable",
      "a:named-expression",
      "b:variable",
      "a:variable",
      "c:variable",
      "b:variable",
      "a:variable",
      "b:variable",
      "c:variable",
      "c:named-expression",
    ]);
    // Only the names outside any scope are unresolved defined names.
    expect(
      analysis.diagnostics.map((diagnostic) =>
        analysis.formula.slice(diagnostic.span.start, diagnostic.span.end)
      )
    ).toEqual(["a", "c"]);
    expect(diagnosticCodes(analysis)).toEqual([
      "reference.missing-name",
      "reference.missing-name",
    ]);
  });

  test("degrades gracefully while a LET is typed", () => {
    const partial = (formula: string) => {
      const analysis = analyzeFormula({ formula, origin });
      expect(diagnosticCodes(analysis)).not.toContain("reference.missing-name");
      return nameTokens(analysis);
    };

    expect(partial("=LET(")).toEqual([]);
    expect(partial("=LET(total")).toEqual(["total:variable"]);
    expect(partial("=LET(total,")).toEqual(["total:variable"]);
    expect(partial("=LET(x, 1, x")).toEqual(["x:variable", "x:variable"]);
    expect(partial("=LET(x, 1, x*")).toEqual(["x:variable", "x:variable"]);
    expect(partial("=LET(x, 1, y")).toEqual(["x:variable", "y:variable"]);
    expect(partial("=LET(x, 1, y, x+")).toEqual([
      "x:variable",
      "y:variable",
      "x:variable",
    ]);
    expect(partial("=SUM(LET(x, A2, LET(y, x")).toEqual([
      "x:variable",
      "A2:cell-reference",
      "y:variable",
      "x:variable",
    ]);
  });

  test("a trailing name in an unfinished LET can still be a defined name", () => {
    const engine = buildEngine();
    engine.addNamedExpression({ expressionName: "Rate", expression: "0.5" });
    const analysis = engine.analyzeFormula({ formula: "=LET(x, 1, Rate", origin });

    expect(nameTokens(analysis)).toEqual([
      "x:variable",
      "Rate:named-expression",
    ]);
    expect(analysis.references[0]?.resolution.status).toBe("resolved");
  });

  test("reports invalid and repeated LET names", () => {
    const invalid = analyzeFormula({ formula: "=LET(x1, 5, x1)", origin });
    expect(nameTokens(invalid)).toEqual([
      "x1:cell-reference",
      "x1:cell-reference",
    ]);
    expect(invalid.diagnostics).toEqual([
      {
        span: { start: 5, end: 7 },
        severity: "warning",
        code: "let.invalid-name",
        message:
          "x1 cannot be a LET name because it is a cell reference. LET names are plain names such as total or x_1.",
      },
    ]);

    const qualified = analyzeFormula({ formula: "=LET(Sheet1!x, 5, 1)", origin });
    expect(diagnosticCodes(qualified)).toContain("let.invalid-name");

    const repeated = analyzeFormula({ formula: "=LET(x, 1, X, 2, x)", origin });
    expect(nameTokens(repeated)).toEqual([
      "x:variable",
      "X:variable",
      "x:variable",
    ]);
    expect(repeated.diagnostics).toEqual([
      {
        span: { start: 11, end: 12 },
        severity: "warning",
        code: "let.duplicate-name",
        message: "The name X is already declared in this LET.",
      },
    ]);
  });

  test("leaves table columns and other functions' arguments alone", () => {
    const analysis = analyzeFormula({
      formula: "=LET(x, 1, Sales[x] + SUM(x, y))",
      origin,
    });

    expect(nameTokens(analysis)).toEqual([
      "x:variable",
      "x:variable",
      "y:named-expression",
    ]);
    expect(
      analysis.tokens
        .filter((token) => token.kind === "table-column")
        .map((token) => analysis.formula.slice(token.span.start, token.span.end))
    ).toEqual(["x"]);
  });
});
