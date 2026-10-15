import { describe, expect, spyOn, test } from "bun:test";
import { functions } from "../../../src/functions/function-registry";
import { FormulaEngine } from "../../../src/core/engine";
import { deserialize } from "../../../src/core/map-serializer";
import { cellAddressToKey, parseCellReference } from "../../../src/core/utils";
import type { SerializedCellValue } from "../../../src/core/types";

const scope = { workbookName: "Book", sheetName: "Sheet" };
const cell = (ref: string) => ({ ...scope, ...parseCellReference(ref) });
function build(entries: [string, SerializedCellValue][] = []) {
  const engine = FormulaEngine.buildEmpty({ undoRedo: { maxEntries: 17 } });
  engine.addWorkbook({ workbookName: scope.workbookName });
  engine.addSheet(scope);
  engine.setSheetContent(scope, new Map(entries));
  return engine;
}
const node = (engine: FormulaEngine, ref: string) =>
  engine._dependencyManager.getCellValueOrEmptyCellNode(cellAddressToKey(cell(ref)));

describe("FormulaEngine.clearCache", () => {
  test("invalidates lazily, retains node identity, clears orders/SCCs and notifies after invalidation", () => {
    const engine = build([["A1", 2], ["B1", "=A1+3"]]);
    expect(engine.getCellValue(cell("B1"))).toBe(5);
    const original = node(engine, "B1");
    expect(original.resolved).toBe(true);
    let updates = 0;
    const unsubscribe = engine.onUpdate(() => {
      updates++;
      expect(original.resolved).toBe(false);
      const snapshot = deserialize(engine.serializeEngine()) as any;
      expect(snapshot.managers.cache).toEqual({ evaluationOrders: [], sccs: [] });
      expect(engine.getCellValue(cell("B1"))).toBe(5);
    });
    expect(engine.clearCache()).toBeUndefined();
    expect(updates).toBe(1);
    expect(node(engine, "B1")).toBe(original);
    unsubscribe();
    engine.clearCache();
    engine.clearCache();
    expect(updates).toBe(1);
    expect(original.resolved).toBe(false);
    engine.setCellContent(cell("A1"), 7);
    expect(engine.getCellValue(cell("B1"))).toBe(10);
    engine.setCellContent(cell("A1"), 9);
    expect(engine.getCellValue(cell("B1"))).toBe(12);
  });

  test("recomputes cached formulas and virtual roots without invalidating another engine", () => {
    const evaluate = spyOn(functions.SUM!, "evaluate").mockReturnValue({ type: "value", result: { type: "number", value: 0.25 } });
    try {
      const engine = build([["A1", "=SUM(1)"], ["B1", "=A1*2"]]);
      const other = build([["A1", "=SUM(1)"]]);
      expect(engine.getCellValue(cell("B1"))).toBe(0.5);
      expect(engine.evaluateFormula("=SUM(1)+1", cell("C1"))).toBe(1.25);
      expect(other.getCellValue(cell("A1"))).toBe(0.25);
      evaluate.mockReturnValue({ type: "value", result: { type: "number", value: 0.75 } });
      expect(engine.getCellValue(cell("B1"))).toBe(0.5);
      engine.clearCache();
      expect(engine.getCellValue(cell("B1"))).toBe(1.5);
      expect(engine.evaluateFormula("=SUM(1)+1", cell("C1"))).toBe(1.75);
      expect(other.getCellValue(cell("A1"))).toBe(0.25);
    } finally {
      evaluate.mockRestore();
    }
  });

  test("preserves workbook objects, ancillary data, configuration and both history stacks", () => {
    const engine = build([["A1", "Amount"], ["A2", 4], ["C1", "=SUM(T[Amount])+X"]]);
    engine.addTable({ ...scope, tableName: "T", start: "A1", numRows: { type: "number", value: 1 }, numCols: 1 });
    engine.addNamedExpression({ expressionName: "X", expression: "2" });
    const area = { ...scope, range: { start: { col: 0, row: 0 }, end: { col: { type: "number" as const, value: 0 }, row: { type: "number" as const, value: 1 } } } };
    engine.addCellStyle({ areas: [area], style: { bold: true } });
    engine.addRangeMetadata({ id: "meta", areas: [area], metadata: { note: "keep" } });
    engine.setWorkbookMetadata("Book", { note: "workbook" });
    engine.setSheetMetadata(scope, { note: "sheet" });
    engine.setCellMetadata(cell("A2"), { note: "cell" });
    engine.createRef(area);
    engine.clearUndoRedoHistory();
    engine.setCellContent(cell("A2"), 8);
    engine.setCellContent(cell("A2"), 12);
    engine.undo();
    expect(engine.getCellValue(cell("C1"))).toBe(10);
    const state = engine.getState();
    const serializedState = structuredClone(state);
    const workbook = engine.getWorkbooks().get("Book")!;
    const history = engine.getUndoRedoState();
    engine.clearCache();
    expect(engine.getWorkbooks().get("Book")).toBe(workbook);
    expect(engine.getState()).toEqual(serializedState);
    expect(engine.getState().workbooks).toBe(state.workbooks);
    expect(engine.getUndoRedoState()).toEqual(history);
    expect(engine.getCellValue(cell("C1"))).toBe(10);
    expect(engine.redo()).toBe(true);
    expect(engine.getCellValue(cell("C1"))).toBe(14);
    expect(engine.undo()).toBe(true);
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(cell("C1"))).toBe(6);
    engine.updateNamedExpression({ expressionName: "X", expression: "10" });
    expect(engine.getCellValue(cell("C1"))).toBe(14);
  });

  test("rebuilds spills from a direct target read and handles shape/blocker changes", () => {
    const engine = build([["D1", 2], ["A1", "=SEQUENCE(D1,2)"], ["F1", "=SUM(A1:B3)"]]);
    expect(engine.getCellValue(cell("B2"))).toBe(4);
    expect(engine.getCellValue(cell("F1"))).toBe(10);
    for (let i = 0; i < 3; i++) {
      engine.clearCache();
      expect(engine.getCellValue(cell("B2"))).toBe(4);
      expect(engine.getCellValue(cell("F1"))).toBe(10);
    }
    engine.setCellContent(cell("D1"), 3);
    engine.clearCache();
    expect(engine.getCellValue(cell("B3"))).toBe(6);
    expect(engine.getCellValue(cell("F1"))).toBe(21);
    engine.setCellContent(cell("B2"), 99);
    engine.clearCache();
    expect(engine.getCellValue(cell("A1"))).toBe("#SPILL!");
    engine.setCellContent(cell("B2"), undefined);
    engine.clearCache();
    expect(engine.getCellValue(cell("B3"))).toBe(6);
  });

  test("refreshes criteria ranges, virtual formulas, cycles and restored warm snapshots", () => {
    const engine = build([["A1", 10], ["A2", 30], ["B1", "yes"], ["B2", "yes"], ["C1", '=AVERAGEIFS(A1:A2,B1:B2,"yes")'], ["D1", "=D2"], ["D2", "=D1"]]);
    expect(engine.getCellValue(cell("C1"))).toBe(20);
    expect(engine.evaluateFormula("=SUM(A1:A2)", cell("E1"))).toBe(40);
    const cycle = engine.getCellValue(cell("D1"))!;
    const restored = FormulaEngine.buildEmpty();
    restored.resetToSerializedEngine(engine.serializeEngine());
    for (const target of [engine, restored]) {
      target.clearCache();
      expect(target.getCellValue(cell("C1"))).toBe(20);
      expect(target.evaluateFormula("=SUM(A1:A2)", cell("E1"))).toBe(40);
      expect(target.getCellValue(cell("D1"))).toBe(cycle);
      target.setCellContent(cell("A2"), 50);
      expect(target.getCellValue(cell("C1"))).toBe(30);
      expect(target.evaluateFormula("=SUM(A1:A2)", cell("E1"))).toBe(60);
      target.setCellContent(cell("D2"), 5);
      expect(target.getCellValue(cell("D1"))).toBe(5);
    }
  });

  test("empty engines notify; transactions coalesce notifications without history entries", () => {
    const engine = FormulaEngine.buildEmpty();
    let updates = 0;
    engine.onUpdate(() => updates++);
    engine.clearCache();
    expect(updates).toBe(1);
    const history = engine.getUndoRedoState();
    engine.transact(() => { engine.clearCache(); engine.clearCache(); });
    expect(updates).toBe(2);
    expect(engine.getUndoRedoState()).toEqual(history);
  });
});
