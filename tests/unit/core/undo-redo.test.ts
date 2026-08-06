import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type {
  DirectCellStyle,
  RangeAddress,
  SerializedCellValue,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";

function cell(ref: string) {
  return {
    workbookName,
    sheetName,
    ...parseCellReference(ref),
  };
}

function sheetCell(targetSheetName: string, ref: string) {
  return {
    workbookName,
    sheetName: targetSheetName,
    ...parseCellReference(ref),
  };
}

function range(
  targetSheetName: string,
  start: string,
  end = start
): RangeAddress {
  const startAddress = parseCellReference(start);
  const endAddress = parseCellReference(end);
  return {
    workbookName,
    sheetName: targetSheetName,
    range: {
      start: { col: startAddress.colIndex, row: startAddress.rowIndex },
      end: {
        col: { type: "number", value: endAddress.colIndex },
        row: { type: "number", value: endAddress.rowIndex },
      },
    },
  };
}

function buildUndoableEngine() {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  engine.clearUndoRedoHistory();
  return engine;
}

function content(engine: FormulaEngine, ref: string): SerializedCellValue | "" {
  return engine.getSheetSerialized({ workbookName, sheetName }).get(ref) ?? "";
}

describe("FormulaEngine undo/redo", () => {
  test("is enabled by default", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    engine.clearUndoRedoHistory();

    engine.setCellContent(cell("A1"), 10);

    expect(engine.getUndoRedoState()).toMatchObject({
      enabled: true,
      canUndo: true,
      canRedo: false,
      undoDepth: 1,
      redoDepth: 0,
    });
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(cell("A1"))).toBe("");
  });

  test("retains configurable entry and byte limits while always enabled", () => {
    const engine = FormulaEngine.buildEmpty({
      undoRedo: { maxEntries: 1, maxBytes: 1_024 },
    });
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    engine.clearUndoRedoHistory();

    engine.setCellContent(cell("A1"), 1);
    engine.setCellContent(cell("A2"), 2);

    expect(engine.getUndoRedoState()).toMatchObject({
      enabled: true,
      maxEntries: 1,
      maxBytes: 1_024,
      undoDepth: 1,
    });
    expect(engine.getUndoRedoState().undoBytes).toBeGreaterThan(0);
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(cell("A1"))).toBe(1);
    expect(engine.getCellValue(cell("A2"))).toBe("");
    expect(engine.undo()).toBe(false);
  });

  test("undoes and redoes cell edits while formulas recalculate", () => {
    const engine = buildUndoableEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["B1", "=A1*2"],
      ])
    );
    expect(engine.getCellValue(cell("B1"))).toBe(2);
    engine.clearUndoRedoHistory();

    engine.setCellContent(cell("A1"), 5);

    expect(engine.getCellValue(cell("B1"))).toBe(10);
    expect(engine.canUndo()).toBe(true);
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(cell("A1"))).toBe(1);
    expect(engine.getCellValue(cell("B1"))).toBe(2);

    expect(engine.redo()).toBe(true);
    expect(engine.getCellValue(cell("A1"))).toBe(5);
    expect(engine.getCellValue(cell("B1"))).toBe(10);
  });

  test("clears redo after undo followed by a new mutation", () => {
    const engine = buildUndoableEngine();

    engine.setCellContent(cell("A1"), 1);
    engine.setCellContent(cell("A2"), 2);
    expect(engine.undo()).toBe(true);
    expect(engine.canRedo()).toBe(true);

    engine.setCellContent(cell("B1"), 3);

    expect(engine.canRedo()).toBe(false);
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: true,
      canRedo: false,
      undoDepth: 2,
      redoDepth: 0,
    });
  });

  test("does not record semantic no-op mutations", () => {
    const engine = buildUndoableEngine();

    engine.setCellContent(cell("A1"), 1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);

    engine.setCellContent(cell("A1"), 1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);

    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(cell("A1"))).toBe("");
    expect(engine.canUndo()).toBe(false);
  });

  test("records nested move operations as one undo step", () => {
    const engine = buildUndoableEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["B1", "=A1"],
      ])
    );
    engine.clearUndoRedoHistory();

    engine.moveCell(cell("A1"), cell("C1"));

    expect(content(engine, "A1")).toBe("");
    expect(content(engine, "C1")).toBe(10);
    expect(engine.getCellValue(cell("B1"))).toBe(10);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);

    expect(engine.undo()).toBe(true);
    expect(content(engine, "A1")).toBe(10);
    expect(content(engine, "C1")).toBe("");
    expect(engine.getCellValue(cell("B1"))).toBe(10);
    expect(engine.canUndo()).toBe(false);
  });

  test("restores sheet and workbook renames with formulas and references", () => {
    const engine = buildUndoableEngine();
    engine.addSheet({ workbookName, sheetName: "Summary" });
    engine.setCellContent(cell("A1"), 7);
    engine.setCellContent(sheetCell("Summary", "A1"), "=Sheet1!A1");
    const refId = engine.createRef(range(sheetName, "A1"));
    engine.clearUndoRedoHistory();

    engine.renameSheet({
      workbookName,
      sheetName,
      newSheetName: "Data",
    });

    expect(engine.getCellValue(sheetCell("Summary", "A1"))).toBe(7);
    expect(engine.getRefAddress(refId)?.sheetName).toBe("Data");
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(sheetCell("Summary", "A1"))).toBe(7);
    expect(engine.getRefAddress(refId)?.sheetName).toBe(sheetName);
    expect(engine.redo()).toBe(true);
    expect(engine.getRefAddress(refId)?.sheetName).toBe("Data");

    engine.clearUndoRedoHistory();
    engine.renameWorkbook({
      workbookName,
      newWorkbookName: "RenamedBook",
    });

    expect(engine.hasWorkbook("RenamedBook")).toBe(true);
    expect(engine.getRefAddress(refId)?.workbookName).toBe("RenamedBook");
    expect(engine.undo()).toBe(true);
    expect(engine.hasWorkbook(workbookName)).toBe(true);
    expect(engine.getRefAddress(refId)?.workbookName).toBe(workbookName);
  });

  test("restores styles, range metadata, tables, and tracked references", () => {
    const engine = buildUndoableEngine();
    const style: DirectCellStyle = {
      areas: [range(sheetName, "A1", "B2")],
      style: { bold: true, backgroundColor: "#ffeeaa" },
    };
    let refId = "";

    engine.transact(() => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "Name"],
          ["B1", "Value"],
          ["A2", "alpha"],
          ["B2", 12],
        ])
      );
      engine.addCellStyle(style);
      engine.addRangeMetadata({
        id: "metadata-1",
        areas: [range(sheetName, "A1", "B2")],
        metadata: { kind: "test" },
      });
      engine.addTable({
        workbookName,
        sheetName,
        tableName: "Table1",
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      });
      refId = engine.createRef(range(sheetName, "A1", "B2"));
    });

    expect(engine.getUndoRedoState().undoDepth).toBe(1);
    expect(engine.getAllCellStyles()).toHaveLength(1);
    expect(engine.getAllRangeMetadata()).toHaveLength(1);
    expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(true);
    expect(engine.getRefAddress(refId)).toBeDefined();

    expect(engine.undo()).toBe(true);
    expect(engine.getAllCellStyles()).toHaveLength(0);
    expect(engine.getAllRangeMetadata()).toHaveLength(0);
    expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(false);
    expect(engine.getRefAddress(refId)).toBeUndefined();

    expect(engine.redo()).toBe(true);
    expect(engine.getAllCellStyles()).toEqual([style]);
    expect(engine.getAllRangeMetadata()).toHaveLength(1);
    expect(engine.hasTable({ workbookName, tableName: "Table1" })).toBe(true);
    expect(engine.getRefAddress(refId)).toBeDefined();
  });

  test("transact groups multiple cell edits into one undo step", () => {
    const engine = buildUndoableEngine();

    engine.transact(() => {
      engine.setCellContent(cell("A1"), 1);
      engine.setCellContent(cell("A2"), 2);
    });

    expect(engine.getUndoRedoState().undoDepth).toBe(1);
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(cell("A1"))).toBe("");
    expect(engine.getCellValue(cell("A2"))).toBe("");
  });

  test("clearUndoRedoHistory clears stacks without changing workbook state", () => {
    const engine = buildUndoableEngine();

    engine.setCellContent(cell("A1"), 42);
    expect(engine.canUndo()).toBe(true);

    engine.clearUndoRedoHistory();

    expect(engine.canUndo()).toBe(false);
    expect(engine.canRedo()).toBe(false);
    expect(engine.getCellValue(cell("A1"))).toBe(42);
  });

  test("resetToSerializedEngine restores data and clears history", () => {
    const source = FormulaEngine.buildEmpty();
    source.addWorkbook(workbookName);
    source.addSheet({ workbookName, sheetName });
    source.setCellContent(cell("A1"), 99);
    const serialized = source.serializeEngine();

    const engine = buildUndoableEngine();
    engine.setCellContent(cell("A1"), 1);
    expect(engine.canUndo()).toBe(true);

    engine.resetToSerializedEngine(serialized);

    expect(engine.getCellValue(cell("A1"))).toBe(99);
    expect(engine.canUndo()).toBe(false);
    expect(engine.canRedo()).toBe(false);
    expect(engine.undo()).toBe(false);
  });

  test("tracked reference creation and deletion emit updates", () => {
    const engine = buildUndoableEngine();
    let updates = 0;
    engine.onUpdate(() => {
      updates++;
    });

    const refId = engine.createRef(range(sheetName, "A1"));
    expect(updates).toBe(1);

    expect(engine.deleteRef(refId)).toBe(true);
    expect(updates).toBe(2);

    expect(engine.deleteRef(refId)).toBe(false);
    expect(updates).toBe(2);
  });
});
