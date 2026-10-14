import { describe, expect, spyOn, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { EvaluateAllCellsResult, SerializedCellValue, SpilledValuesEvaluationResult } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";
import { readCriteriaRange, registerCriteriaRangeReader } from "../../../src/evaluator/criteria-range-cache";
import type { EvaluationContext } from "../../../src/evaluator/evaluation-context";
import { FormulaEvaluator } from "../../../src/evaluator/formula-evaluator";
import { EvaluationManager } from "../../../src/core/managers/evaluation-manager";

const address = (ref: string, workbookName = "Book", sheetName = "Sheet") => ({
  workbookName, sheetName, ...parseCellReference(ref),
});
function build(workbookName = "Book", base = 10) {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName });
  engine.addSheet({ workbookName, sheetName: "Sheet" });
  engine.setSheetContent({ workbookName, sheetName: "Sheet" }, new Map<string, SerializedCellValue>([
    ["A1", base], ["A2", base * 2], ["A3", base * 3],
    ["B1", "yes"], ["B2", "no"], ["B3", "yes"],
    ["C1", '=AVERAGEIFS(A1:A3,B1:B3,"yes")'],
    ["C2", '=AVERAGEIFS(A1:A3,B1:B3,"no")'],
  ]));
  return engine;
}

describe("criteria raw range cache", () => {
  test("reuses coordinate indexes only for current trusted result identity", () => {
    const evaluator = {} as FormulaEvaluator;
    const context = { cellAddress: address("C1") } as EvaluationContext;
    let result: EvaluateAllCellsResult = { type: "values", values: [
      { relativePos: { x: 0, y: 0 }, result: { type: "value", result: { type: "number", value: 1 } } },
    ] };
    let reads = 0;
    let evaluations = 0;
    const range: SpilledValuesEvaluationResult = {
      type: "spilled-values",
      source: "test raw range",
      spillArea: () => ({ start: { col: 0, row: 0 }, end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } } }),
      evaluate: () => ({ type: "value", result: { type: "number", value: 1 } }),
      evaluateAllCells() { evaluations++; return result; },
    };
    registerCriteriaRangeReader(range, evaluator, () => { reads++; return result; });
    const read = () => readCriteriaRange(range, evaluator, context, "row-major", true);
    const first = read();
    for (let i = 0; i < 100; i++) expect(read().index).toBe(first.index!);
    expect(reads).toBe(101); // dependency registration runs on every hit
    expect(evaluations).toBe(0);
    result = { type: "values", values: [] };
    expect(read().index).not.toBe(first.index);
    expect(read().index?.size).toBe(0);
    // Metadata/object spreading cannot turn a transformed range into a raw reader.
    const transformed = { ...range };
    readCriteriaRange(transformed, evaluator, context, "row-major", true);
    expect(evaluations).toBe(1);
    // A result passed to a different evaluator must not use the old engine reader.
    readCriteriaRange(range, {} as FormulaEvaluator, context, "row-major", true);
    expect(evaluations).toBe(2);
  });

  test("sourceRange transforms, sparse ranges and formula source changes stay correct", () => {
    const engine = build();
    engine.setCellContent(address("C3"), '=AVERAGEIFS(A1:A3,ROW(A1:A3),">1")');
    expect(engine.getCellValue(address("C3"))).toBe(25);
    engine.setCellContent(address("C4"), '=AVERAGEIFS(A1:A3,B1:B3,"y*")');
    expect(engine.getCellValue(address("C4"))).toBe(20);
    engine.setCellContent(address("A3"), '=A1*5');
    expect(engine.getCellValue(address("C1"))).toBe(30);
    engine.setCellContent(address("A1"), 20);
    expect(engine.getCellValue(address("C1"))).toBe(60);
    engine.setCellContent(address("B3"), "");
    expect(engine.getCellValue(address("C1"))).toBe(20);
    engine.setCellContent(address("B2"), "yes");
    expect(engine.getCellValue(address("C1"))).toBe(20);
    expect(engine.getCellValue(address("C3"))).toBe(60);
  });

  test("mutations, undo/redo and evaluation reset replace indexed values", () => {
    const engine = build();
    const get = () => engine.getCellValue(address("C1"));
    expect(get()).toBe(20);
    engine.clearUndoRedoHistory();
    engine.setCellContent(address("B2"), "yes");
    engine.setCellContent(address("A2"), 80);
    expect(get()).toBe(40);
    expect(engine.undo()).toBe(true);
    expect(get()).toBe(20);
    expect(engine.undo()).toBe(true);
    expect(get()).toBe(20);
    expect(engine.redo()).toBe(true);
    expect(engine.redo()).toBe(true);
    expect(get()).toBe(40);
    engine._evaluationManager.clearEvaluationCache();
    expect(get()).toBe(40);
    engine.setCellContent(address("B1"), "no");
    expect(get()).toBe(55);
    const replacement = build("Book", 100);
    engine.resetToSerializedEngine(replacement.serializeEngine({ includeEvaluationState: false }));
    expect(get()).toBe(200);
    replacement.getCellValue(address("C1"));
    engine.resetToSerializedEngine(replacement.serializeEngine());
    expect(get()).toBe(200);
    engine.setCellContent(address("A1"), 500);
    expect(get()).toBe(400);
  });

  test("isolates workbooks, sheets and engines with identical source coordinates", () => {
    const engine = build();
    engine.addWorkbook({ workbookName: "Other" });
    engine.addSheet({ workbookName: "Other", sheetName: "Sheet" });
    engine.setSheetContent({ workbookName: "Other", sheetName: "Sheet" }, new Map<string, SerializedCellValue>([
      ["A1", 100], ["A2", 200], ["A3", 300], ["B1", "yes"], ["B2", "no"], ["B3", "yes"],
      ["C1", '=AVERAGEIFS(A1:A3,B1:B3,"yes")'],
    ]));
    engine.addSheet({ workbookName: "Book", sheetName: "Second" });
    engine.setSheetContent({ workbookName: "Book", sheetName: "Second" }, new Map<string, SerializedCellValue>([
      ["A1", 4], ["A2", 6], ["B1", "yes"], ["B2", "yes"],
      ["C1", '=AVERAGEIFS(A1:A3,B1:B3,"yes")'],
    ]));
    expect(engine.getCellValue(address("C1"))).toBe(20);
    expect(engine.getCellValue(address("C1", "Other"))).toBe(200);
    expect(engine.getCellValue(address("C1", "Book", "Second"))).toBe(5);
    const otherEngine = build("Book", 1000);
    expect(otherEngine.getCellValue(address("C1"))).toBe(2000);
    expect(engine.getCellValue(address("C1"))).toBe(20);
    engine.setCellContent(address("B3", "Other"), "no");
    expect(engine.getCellValue(address("C1", "Other"))).toBe(100);
    expect(engine.getCellValue(address("C1"))).toBe(20);
  });

  test("warm table snapshots retain trusted readers and transformed ROW values", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: "Book" });
    engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });
    const cells = new Map<string, SerializedCellValue>([["A1", "Amount"], ["B1", "Group"]]);
    for (let i = 2; i <= 101; i++) {
      cells.set(`A${i}`, i - 1);
      cells.set(`B${i}`, (i - 1) % 2);
      cells.set(`D${i}`, (i - 1) % 2);
      cells.set(`C${i}`, `=AVERAGEIFS(Data[Amount],Data[Group],D${i})`);
    }
    cells.set("E1", '=AVERAGEIFS(Data[Amount],ROW(Data[Amount]),">50")');
    engine.setSheetContent({ workbookName: "Book", sheetName: "Sheet" }, cells);
    engine.addTable({ workbookName: "Book", sheetName: "Sheet", tableName: "Data", start: "A1", numCols: 2, numRows: { type: "number", value: 100 } });
    expect(engine.getCellValue(address("C2"))).toBe(50);
    expect(engine.getCellValue(address("E1"))).toBe(75);
    const snapshot = engine.serializeEngine();
    let restoredAllCellsCalls = 0;
    const original = EvaluationManager.prototype.deserializeFunctionEvaluationResultSnapshot;
    const spy = spyOn(EvaluationManager.prototype, "deserializeFunctionEvaluationResultSnapshot").mockImplementation(function (this: EvaluationManager, value, resolve) {
      const result = original.call(this, value, resolve);
      if (result.type === "spilled-values" && value.type === "spilled-values" && value.spill.kind === "source-range") {
        const originalAll = result.evaluateAllCells;
        result.evaluateAllCells = function (options) {
          restoredAllCellsCalls++;
          return originalAll.call(this, options);
        };
      }
      return result;
    });
    try {
      engine.resetToSerializedEngine(snapshot);
      // Cold consumers reuse restored raw table ranges. Warm E1 then recomputes
      // against restored ROW's materialized values, not the raw Amount column.
      for (let i = 2; i <= 101; i++) expect(engine.getCellValue(address(`C${i}`))).toBe((i - 1) % 2 ? 50 : 51);
      engine.setCellContent(address("F1"), '=AVERAGEIFS(Data[Amount],ROW(Data[Amount]),">50")+1');
      expect(engine.getCellValue(address("F1"))).toBe(76);
      expect(restoredAllCellsCalls).toBe(0);
      // Re-serialization must also recognize restored raw readers.
      engine.resetToSerializedEngine(engine.serializeEngine());
      engine.setCellContent(address("D3"), 1);
      expect(engine.getCellValue(address("C3"))).toBe(50);
      expect(restoredAllCellsCalls).toBe(0);
      engine.setCellContent(address("B2"), 0);
      expect(engine.getCellValue(address("C2"))).toBe(51);
      expect(engine.undo()).toBe(true);
      expect(engine.getCellValue(address("C2"))).toBe(50);
    } finally { spy.mockRestore(); }
  });

  test("repeated AVERAGEIFS bypasses raw evaluateAllCells without losing consumer dependencies", () => {
    const original = FormulaEvaluator.prototype.evaluateRange;
    let allCellsCalls = 0;
    const spy = spyOn(FormulaEvaluator.prototype, "evaluateRange").mockImplementation(function (this: FormulaEvaluator, node, context) {
      const result = original.call(this, node, context);
      if (result.type === "spilled-values") {
        const originalAll = result.evaluateAllCells;
        result.evaluateAllCells = function (options) {
          allCellsCalls++;
          return originalAll.call(this, options);
        };
      }
      return result;
    });
    try {
      const engine = build();
      const cells = new Map<string, SerializedCellValue>();
      for (let i = 1; i <= 200; i++) {
        cells.set(`A${i}`, i);
        cells.set(`B${i}`, i % 2);
        cells.set(`D${i}`, i % 2);
        cells.set(`C${i}`, `=AVERAGEIFS(A1:A200,B1:B200,D${i})`);
      }
      engine.setSheetContent({ workbookName: "Book", sheetName: "Sheet" }, cells);
      for (let i = 1; i <= 200; i++) expect(engine.getCellValue(address(`C${i}`))).toBe(i % 2 ? 100 : 101);
      expect(allCellsCalls).toBe(0);
      engine.setCellContent(address("A1"), 201);
      for (let i = 1; i <= 200; i++) expect(engine.getCellValue(address(`C${i}`))).toBe(i % 2 ? 102 : 101);
      expect(allCellsCalls).toBe(0);
    } finally { spy.mockRestore(); }
  });
});
