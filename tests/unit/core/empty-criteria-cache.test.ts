import { expect, spyOn, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { DependencyManager } from "../../../src/core/managers/dependency-manager";
import { FormulaEvaluator } from "../../../src/evaluator/formula-evaluator";
import { parseCellReference } from "../../../src/core/utils";
import { FormulaError, type SerializedCellValue, type SpilledValuesEvaluationResult, type EvaluateAllCellsResult } from "../../../src/core/types";
import { registerCriteriaRangeReader } from "../../../src/evaluator/criteria-range-cache";
import type { EvaluationContext } from "../../../src/evaluator/evaluation-context";
import { processMultiCriteriaValues } from "../../../src/functions/criteria-utils";

const address = (ref: string, sheetName = "Sheet") => ({ workbookName: "Book", sheetName, ...parseCellReference(ref) });
function build(cells: [string, SerializedCellValue][]) {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName: "Book" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });
  engine.setSheetContent(address("A1"), new Map(cells));
  return engine;
}

test("dense empty criteria distinguishes missing, explicit/formula blanks, zero, false and errors", () => {
  const e = build([
    ["A1", 10], ["A2", 20], ["A3", 30], ["A4", 40], ["A5", 50], ["A6", 60], ["A7", 70],
    ["B2", ""], ["B3", '=""'], ["B4", 0], ["B5", false], ["B6", "0"], ["B7", "=NA()"],
    ["D1", '=COUNTIFS(B1:B7,"")'], ["D2", '=SUMIFS(A1:A7,B1:B7,"")'],
    ["D3", '=AVERAGEIFS(A1:A7,B1:B7,"")'],
    ["D4", '=COUNTIFS(B1:C7,"")'],
    ["D5", '=COUNTIFS(C1:C7,"",B1:B7,"<>")'],
    ["D7", '=SUMIFS(A1:A7,C1:C7,"",B1:B7,"0*")'],
    ["D8", '=COUNTIFS(C1:C7,"",B1:B7,">1")'],
    ["D9", '=SUMIFS(B7:B8,C7:C8,"")'],
  ]);
  expect(e.getCellValue(address("D1"))).toBe(3);
  expect(e.getCellValue(address("D2"))).toBe(60);
  expect(e.getCellValue(address("D3"))).toBe(20);
  expect(e.getCellValue(address("D4"))).toBe(10);
  expect(e.getCellValue(address("D5"))).toBe(4); // Mixed <> retains error-as-nonempty behavior.
  expect(e.getCellValue(address("D7"))).toBe(60);
  expect(e.getCellValue(address("D8"))).toBe(0);
  expect(e.getCellValue(address("D9"))).toBe(0); // SUMIFS skips matched errors.
});

test("empty criteria keeps dense missing values, shifted bounds and current short-range semantics", () => {
  const e = build([
    ["B3", 12], ["C4", 18],
    ["F1", '=COUNTIFS(B3:C5,"")'],
    ["F2", '=SUMIFS(B3:C5,E7:F9,"")'],
    ["F3", '=AVERAGEIFS(B3:C5,E7:F9,"")'],
    ["F4", '=COUNTIFS(B3:C5,"",H1:H2,"")'],
    ["F5", '=COUNTIFS(B:B,"")'],
  ]);
  expect(e.getCellValue(address("F1"))).toBe(4);
  expect(e.getCellValue(address("F2"))).toBe(30);
  expect(e.getCellValue(address("F3"))).toBe(15);
  expect(e.getCellValue(address("F4"))).toBe(4);
  expect(e.getCellValue(address("F5"))).toBe("#VALUE!");
});

test("raw empty-criteria consumers invalidate across edits, undo/redo, clearCache and restore", () => {
  const e = build([["A1", 10], ["A2", 20], ["A3", 30],
    ["D1", '=SUMIFS(A1:A3,B1:B3,"")'], ["D2", '=AVERAGEIFS(A1:A3,B1:B3,"")']]);
  const get = () => [e.getCellValue(address("D1")), e.getCellValue(address("D2"))];
  expect(get()).toEqual([60, 20]);
  e.clearUndoRedoHistory();
  e.setCellContent(address("B2"), "x"); expect(get()).toEqual([40, 20]);
  e.undo(); expect(get()).toEqual([60, 20]);
  e.redo(); expect(get()).toEqual([40, 20]);
  e.setCellContent(address("B2"), ""); expect(get()).toEqual([60, 20]);
  e.clearCache(); expect(get()).toEqual([60, 20]);
  e.resetToSerializedEngine(e.serializeEngine()); expect(get()).toEqual([60, 20]);
  e.setCellContent(address("A3"), 60); expect(get()).toEqual([90, 30]);
  e.setCellContent(address("B1"), '=IF(A1>0,"x","")'); expect(get()).toEqual([80, 40]);
  e.setCellContent(address("A1"), -10); expect(get()).toEqual([70, 70 / 3]);
  e.addSheet({ workbookName: "Book", sheetName: "Other" });
  e.setCellContent(address("A1", "Other"), '=SUMIFS(Sheet!A1:A3,Sheet!B1:B3,"")');
  expect(e.getCellValue(address("A1", "Other"))).toBe(70);
  e.setCellContent(address("B3"), "x");
  expect(e.getCellValue(address("A1", "Other"))).toBe(10);
});

test("transformed value/criteria spills keep callbacks, raw ranges include outside spill anchors", () => {
  const e = build([["A1", "=SEQUENCE(3)"],
    ["D1", '=SUMIFS(A2:A3,B2:B3,"")'],
    ["D2", '=SUMIFS(ROW(A1:A3),B1:B3,"")'],
    ["D3", '=SUMIFS(A1:A3,B1:B3,"",ROW(A1:A3),">1")']]);
  expect(e.getCellValue(address("D1"))).toBe(5);
  expect(e.getCellValue(address("D2"))).toBe(6);
  expect(e.getCellValue(address("D3"))).toBe(5);
  e.setCellContent(address("A1"), "=SEQUENCE(2)");
  expect(e.getCellValue(address("D1"))).toBe(2);
  e.clearCache(); expect(e.getCellValue(address("D1"))).toBe(2);
  e.resetToSerializedEngine(e.serializeEngine());
  expect(e.getCellValue(address("D2"))).toBe(6);
});

test("empty criteria shares range dependencies instead of Q*area cell callbacks", () => {
  const cells: [string, SerializedCellValue][] = [];
  for (let i = 1; i <= 100; i++) { cells.push([`A${i}`, i], [`B${i}`, i % 2 ? "" : "yes"]); }
  for (let i = 1; i <= 10; i++) cells.push([`D${i}`, `=SUMIFS(A1:A100,B1:B100,"",A1:A100,">${i}")`]);
  const e = build(cells);
  const spy = spyOn(DependencyManager.prototype, "getCellValueOrEmptyCellNode");
  try {
    for (const cold of [false, true]) {
      if (cold) e.clearCache();
      spy.mockClear();
      for (let i = 1; i <= 10; i++) {
        let sum = 0; for (let n = i + 1; n <= 100; n++) if (n % 2) sum += n;
        expect(e.getCellValue(address(`D${i}`))).toBe(sum);
      }
      expect(spy.mock.calls.length).toBeLessThan(30);
    }
    e.resetToSerializedEngine(e.serializeEngine());
    e.setCellContent(address("D11"), '=SUMIFS(A1:A100,B1:B100,"",A1:A100,">11")');
    spy.mockClear();
    expect(e.getCellValue(address("D11"))).toBe(2464);
    expect(spy.mock.calls.length).toBeLessThan(10);
    e.setCellContent(address("B99"), "yes");
    expect(e.getCellValue(address("D1"))).toBe(2400);
  } finally { spy.mockRestore(); }
});

test("bulk awaiting/error states are not silently converted to missing blanks", () => {
  const evaluator = {} as FormulaEvaluator;
  const context = { cellAddress: address("D1"), dependencyNode: {} } as EvaluationContext;
  const waiting: EvaluateAllCellsResult = { type: "awaiting-evaluation", waitingFor: context.dependencyNode, errAddress: context.dependencyNode };
  let current: EvaluateAllCellsResult = waiting;
  const range: SpilledValuesEvaluationResult = {
    type: "spilled-values", source: "test",
    spillArea: () => ({ start: { row: 0, col: 0 }, end: { row: { type: "number", value: 1 }, col: { type: "number", value: 0 } } }),
    evaluate: () => { throw new Error("must not read cells while waiting"); },
    evaluateAllCells: () => { throw new Error("must use trusted reader"); },
  };
  registerCriteriaRangeReader(range, evaluator, () => current, () => true);
  const run = () => processMultiCriteriaValues(evaluator, range, [{ rangeResult: range, parsedCriteria: { type: "exact", value: { type: "string", value: "" } } }], context, "row-major");
  expect(run()).toBe(waiting);
  current = { type: "error", err: FormulaError.REF, errAddress: context.dependencyNode };
  expect(run()).toBe(current);
});

// Copying sourceRange metadata must not confer trusted-reader identity.
test("unregistered raw-looking spills retain callback fallback", () => {
  const original = FormulaEvaluator.prototype.evaluateRange;
  let reads = 0;
  const spy = spyOn(FormulaEvaluator.prototype, "evaluateRange").mockImplementation(function(node, context) {
    const result = original.call(this, node, context);
    if (result.type !== "spilled-values") return result;
    return { ...result, evaluate(offset, ctx) { reads++; return result.evaluate(offset, ctx); } };
  });
  try {
    const e = build([["A1", 10], ["A2", 20], ["D1", '=SUMIFS(A1:A2,B1:B2,"")']]);
    expect(e.getCellValue(address("D1"))).toBe(30);
    expect(reads).toBeGreaterThan(0);
  } finally { spy.mockRestore(); }
});
