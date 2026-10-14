import { describe, expect, test, spyOn } from "bun:test";
import { FormulaEngine } from "../../../../src/core/engine";
import { StyleLookupIndex } from "../../../../src/core/managers/style-lookup-index";
import { StyleManager, type StyleDataChange } from "../../../../src/core/managers/style-manager";
import type { EvaluationManager } from "../../../../src/core/managers/evaluation-manager";
import type { CellAddress, ConditionalStyle, DirectCellStyle, RangeAddress } from "../../../../src/core/types";
import { isCellInRange } from "../../../../src/core/utils";

const cell: CellAddress = { workbookName: "book", sheetName: "sheet", rowIndex: 2, colIndex: 2 };
const area = (row = 0, col = 0, endRow = 9, endCol = 9): RangeAddress => ({
  workbookName: "book", sheetName: "sheet",
  range: { start: { row, col }, end: {
    row: endRow === Infinity ? { type: "infinity", sign: "positive" } : { type: "number", value: endRow },
    col: endCol === Infinity ? { type: "infinity", sign: "positive" } : { type: "number", value: endCol },
  } },
});
const direct = (): DirectCellStyle => ({ areas: [area()], style: { bold: true } });
const conditional = (): ConditionalStyle => ({ areas: [area()], condition: {
  type: "formula", formula: "A1", color: { l: 50, c: 30, h: 20 },
} });
const manager = () => new StyleManager({ evaluateFormula: () => false } as unknown as EvaluationManager);

describe("style lookup index", () => {
  test("matches a linear oracle with overlaps, open bounds, multiple sheets and area ordering", () => {
    const rules = Array.from({ length: 515 }, (_, i) => ({ areas: [
      area(i % 31, i % 17, i % 31 + 4, i % 17 + 3),
      area(i % 13, i % 7, Infinity, i % 7 + 1),
      { ...area(0, 0, Infinity, Infinity), sheetName: "other" },
    ] }));
    rules.push({ areas: [area(0, 0, Infinity, Infinity), area(0, 0, 5, Infinity), area()] });
    const index = new StyleLookupIndex(rules);
    for (const sheetName of ["sheet", "other", "missing"]) {
      for (let rowIndex = 0; rowIndex < 45; rowIndex += 3) {
        const address = { ...cell, sheetName, rowIndex, colIndex: rowIndex % 23 };
        const expected = rules.flatMap((rule, ruleIndex) => rule.areas.flatMap((a, areaIndex) =>
          a.sheetName === sheetName && isCellInRange(address, a.range) ? [[ruleIndex, areaIndex]] : []));
        expect(index.query(address).map(m => [m.ruleIndex, m.areaIndex])).toEqual(expected);
      }
    }
    expect(index.query({ ...cell, workbookName: "missing" })).toEqual([]);
  });

  test("warm lookups never reread source areas; repeated direct lookups skip the index", () => {
    let reads = 0;
    const rules = Array.from({ length: 515 }, (_, i) => ({
      get areas() { reads++; return [area(i * 10, 0, i * 10 + 2, 4), area(i * 10, 8, i * 10 + 2, 9), { ...area(), sheetName: "other" }]; },
      style: { fontSize: i + 10 },
    }));
    const index = new StyleLookupIndex(rules);
    reads = 0;
    for (let i = 0; i < 100; i++) expect(index.query(cell)).toHaveLength(1);
    expect(reads).toBe(0);
    const m = manager();
    m.resetStyles([], rules);
    m.getCellStyle(cell);
    const query = spyOn(StyleLookupIndex.prototype, "query");
    try {
      for (let i = 0; i < 100; i++) expect(m.getCellStyle(cell)).toEqual({ fontSize: 10 });
      // Only the conditional candidate index is queried, not the direct index.
      expect(query).toHaveBeenCalledTimes(100);
    } finally { query.mockRestore(); }
  });

  test("spatial pruning avoids visiting unrelated rectangles on the same sheet", () => {
    const index = new StyleLookupIndex(Array.from({ length: 1548 }, (_, i) => ({ areas: [area(i * 10, 0, i * 10 + 2, 5)] })));
    type TestNode = { entries?: { minRow: number }[]; left?: TestNode; right?: TestNode };
    const sheets = (index as unknown as { sheets: Map<string, Map<string, TestNode>> }).sheets;
    let visited = 0;
    const instrument = (node?: TestNode): void => {
      if (!node) return;
      for (const entry of node.entries ?? []) {
        const value = entry.minRow;
        Object.defineProperty(entry, "minRow", { get: () => { visited++; return value; } });
      }
      instrument(node.left);
      instrument(node.right);
    };
    instrument(sheets.get("book")?.get("sheet"));
    expect(index.query(cell)).toHaveLength(1);
    expect(visited).toBeLessThanOrEqual(8);
  });

  test("composes direct rules once and evaluates conditional conditions every time", () => {
    let truth = false;
    let calls = 0;
    const m = new StyleManager({ evaluateFormula: () => { calls++; return truth; } } as unknown as EvaluationManager);
    m.addCellStyle({ areas: [area(), area()], style: { bold: true, color: "red" } });
    m.addCellStyle({ areas: [area()], style: { bold: false, italic: true } });
    m.addConditionalStyle(conditional());
    expect(m.getCellStyle(cell)).toEqual({ bold: false, color: "red", italic: true });
    truth = true;
    expect(m.getCellStyle(cell)?.backgroundColor).toBeDefined();
    truth = false;
    const result = m.getCellStyle(cell)!;
    result.bold = true;
    expect(m.getCellStyle(cell)?.bold).toBe(false);
    expect(calls).toBe(4);
  });

  test("invalidates cached hits and misses across mutation, batch, reset and history paths", () => {
    const m = manager();
    expect(m.getCellStyle(cell)).toBeUndefined();
    m.addCellStyle(direct());
    expect(m.getCellStyle(cell)?.bold).toBe(true);
    m.removeCellStyle("book", 0);
    expect(m.getCellStyle(cell)).toBeUndefined();
    m.batchMutations(() => {
      m.addCellStyle(direct());
      expect(m.getCellStyle(cell)?.bold).toBe(true);
      m.clearCellStylesInRange(area(2, 2, 2, 2));
      expect(m.getCellStyle(cell)).toBeUndefined();
    });
    m.resetStyles([], [direct()]);
    expect(m.getCellStyle(cell)?.bold).toBe(true);
    m.updateSheetName("book", "sheet", "new");
    expect(m.getCellStyle(cell)).toBeUndefined();
    expect(m.getCellStyle({ ...cell, sheetName: "new" })?.bold).toBe(true);
    m.updateWorkbookName("book", "newBook");
    expect(m.getCellStyle({ ...cell, sheetName: "new" })).toBeUndefined();
    expect(m.getCellStyle({ ...cell, workbookName: "newBook", sheetName: "new" })?.bold).toBe(true);
    m.removeWorkbookStyles("newBook");
    expect(m.getCellStyle({ ...cell, workbookName: "newBook", sheetName: "new" })).toBeUndefined();
    m.restoreFromSnapshot({ conditionalStyles: [], cellStyles: [direct()], cellDataTypes: [] });
    expect(m.getCellStyle(cell)?.bold).toBe(true);
    m.removeSheetStyles("book", "sheet");
    expect(m.getCellStyle(cell)).toBeUndefined();
    const changes: StyleDataChange[] = [{ kind: "cell-style", after: { index: 0, value: direct() } }];
    m.applyHistoryChanges(changes, "redo");
    expect(m.getCellStyle(cell)?.bold).toBe(true);
    m.applyHistoryChanges(changes, "undo");
    expect(m.getCellStyle(cell)).toBeUndefined();
    m.addCellStyle(direct());
    m.getCellStyle(cell);
    m.clearCellStyles(area());
    expect(m.getCellStyle(cell)).toBeUndefined();
  });

  test("real engine conditional dependencies stay fresh across cross-sheet edits, history and sheet rename", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: "book" });
    engine.addSheet({ workbookName: "book", sheetName: "sheet" });
    engine.addSheet({ workbookName: "book", sheetName: "inputs" });
    const input = { ...cell, sheetName: "inputs", rowIndex: 0, colIndex: 0 };
    engine.setCellContent(input, 0);
    engine.addCellStyle(direct());
    engine.addConditionalStyle({ areas: [area()], condition: {
      type: "formula", formula: "inputs!A1 > 0", color: { l: 50, c: 30, h: 20 },
    } });
    const check = (address: CellAddress, active: boolean): void => {
      // Repeat each read to exercise warm direct caches and conditional indexes.
      for (let i = 0; i < 3; i++) {
        const style = engine.getCellStyle(address);
        expect(style?.bold).toBe(true);
        expect(style?.backgroundColor !== undefined).toBe(active);
      }
    };
    check(cell, false);
    engine.setCellContent(input, 1);
    check(cell, true);
    expect(engine.undo()).toBe(true);
    check(cell, false);
    expect(engine.redo()).toBe(true);
    check(cell, true);

    engine.renameSheet({ workbookName: "book", sheetName: "sheet", newSheetName: "renamed" });
    const renamed = { ...cell, sheetName: "renamed" };
    check(renamed, true);
    expect(engine.getCellStyle(cell)).toBeUndefined();
    engine.setCellContent(input, 0);
    check(renamed, false);
    expect(engine.undo()).toBe(true);
    check(renamed, true);
    expect(engine.undo()).toBe(true);
    check(cell, true);
    expect(engine.redo()).toBe(true);
    check(renamed, true);
  });

  test("virtual formulas invalidate only affected dependency chains, including empty inputs", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: "book" });
    engine.addSheet({ workbookName: "book", sheetName: "sheet" });
    engine.addSheet({ workbookName: "book", sheetName: "inputs" });
    const input = { ...cell, sheetName: "inputs", rowIndex: 0, colIndex: 0 };
    const dependent = { ...input, colIndex: 1 };
    engine.setCellContent(dependent, "=SUM(A1)+1");
    const evaluator = (engine as unknown as { evaluationManager: EvaluationManager }).evaluationManager;
    expect(engine.evaluateFormula("=inputs!B1+1", cell)).toBe(2);
    const evaluate = spyOn(evaluator, "evaluateCellNode");
    try {
      engine.setCellContent({ ...input, rowIndex: 20 }, 100);
      expect(engine.evaluateFormula("=inputs!B1+1", cell)).toBe(2);
      expect(evaluate).toHaveBeenCalledTimes(0);
      engine.setCellContent(input, 5);
      expect(engine.evaluateFormula("=inputs!B1+1", cell)).toBe(7);
      expect(evaluate.mock.calls.length).toBeGreaterThan(0);
      expect(engine.undo()).toBe(true);
      expect(engine.evaluateFormula("=inputs!B1+1", cell)).toBe(2);
      expect(engine.redo()).toBe(true);
      expect(engine.evaluateFormula("=inputs!B1+1", cell)).toBe(7);
    } finally { evaluate.mockRestore(); }
  });

  test("gradient values and bounds are reevaluated after warming the direct cache", () => {
    let value = 2;
    let max = 10;
    const m = new StyleManager({
      getCellEvaluationResult: () => ({ type: "value", result: { type: "number", value } }),
      evaluateFormula: (formula: string) => formula.startsWith("=MIN") ? 0 : max,
    } as unknown as EvaluationManager);
    m.addCellStyle(direct());
    m.addConditionalStyle({ areas: [area()], condition: {
      type: "gradient",
      min: { type: "lowest_value", color: { l: 10, c: 0, h: 0 } },
      max: { type: "highest_value", color: { l: 90, c: 0, h: 0 } },
    } });
    const original = m.getCellStyle(cell)!.backgroundColor!;
    value = 8;
    expect(m.getCellStyle(cell)?.backgroundColor).not.toBe(original);
    max = 40;
    expect(m.getCellStyle(cell)?.backgroundColor).toBe(original);
  });

  test("direct cache stays bounded during long scrolls and detaches nested output", () => {
    const m = manager();
    m.addCellStyle({ areas: [area(0, 0, Infinity, Infinity)], style: { borderSides: { top: true } } });
    m.getCellStyle(cell)!.borderSides!.top = false;
    expect(m.getCellStyle(cell)?.borderSides?.top).toBe(true);
    for (let rowIndex = 0; rowIndex < 10000; rowIndex++) m.getCellStyle({ ...cell, rowIndex });
    expect((m as unknown as { directStyleCache: Map<string, unknown> }).directStyleCache.size).toBe(8192);
    expect(m.getCellStyle(cell)?.borderSides?.top).toBe(true);
  });

  test("conditional and data-type indexes invalidate independently", () => {
    const m = new StyleManager({ evaluateFormula: () => true } as unknown as EvaluationManager);
    expect(m.getCellStyle(cell)).toBeUndefined();
    m.addConditionalStyle(conditional());
    expect(m.getCellStyle(cell)?.backgroundColor).toBeDefined();
    m.removeConditionalStyle("book", 0);
    expect(m.getCellStyle(cell)).toBeUndefined();
    m.addConditionalStyle(conditional());
    m.getCellStyle(cell);
    m.clearCellStyles(area());
    expect(m.getCellStyle(cell)).toBeUndefined();
    expect(m.getCellDataType(cell)).toBe("general");
    m.addCellDataType({ areas: [area()], dataType: "text" });
    expect(m.getCellDataType(cell)).toBe("text");
    m.addCellDataType({ areas: [area()], dataType: "number" });
    expect(m.getCellDataType(cell)).toBe("number");
    m.clearCellDataTypesInRange(area());
    expect(m.getCellDataType(cell)).toBe("general");
    const changes: StyleDataChange[] = [{ kind: "cell-data-type", after: { index: 0, value: { areas: [area()], dataType: "boolean" } } }];
    m.applyHistoryChanges(changes, "redo");
    expect(m.getCellDataType(cell)).toBe("boolean");
    m.applyHistoryChanges(changes, "undo");
    expect(m.getCellDataType(cell)).toBe("general");
  });
});
