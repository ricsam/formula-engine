import { describe, expect, test } from "bun:test";

import { FormulaEngine } from "../core/engine";
import {
  analyzeFormula,
  findFormulaReferenceAt,
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
  engine.addWorkbook("Book");
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet1" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet2" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet3" });
  return engine;
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
});
