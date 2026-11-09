import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SerializedCellValue } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

/**
 * Multi-step table edits that a user makes as one action, such as injecting a
 * dataset into a table of another shape, run inside `transact` so they are one
 * update notification and one undo step.
 */
describe("Table edits in one transaction", () => {
  const workbookName = "Book";
  const data = { workbookName, sheetName: "Data" };
  const summary = { workbookName, sheetName: "Summary" };
  const tableName = "Results";
  let engine: FormulaEngine;
  let updates: number;

  const summaryValue = (ref: string) =>
    engine.getCellValue({ ...summary, ...parseCellReference(ref) });
  const evaluate = (formula: string) =>
    engine.evaluateFormula(formula, { ...summary, rowIndex: 99, colIndex: 0 });
  const content = (sheet: { workbookName: string; sheetName: string }) =>
    Object.fromEntries(engine.getSheetSerialized(sheet));
  const table = () => engine.getTable({ tableName, workbookName })!;
  const headers = () =>
    Array.from(table().headers.values())
      .sort((a, b) => a.index - b.index)
      .map(({ name }) => name);

  const twoColumns = new Map<string, SerializedCellValue>([
    ["A1", "Name"],
    ["B1", "Score"],
    ["A2", "Ann"],
    ["B2", 1],
    ["A3", "Bob"],
    ["B3", 2],
  ]);
  const threeColumns = new Map<string, SerializedCellValue>([
    ["A1", "Sample"],
    ["B1", "Value"],
    ["C1", "Unit"],
    ["A2", "s1"],
    ["B2", 10],
    ["C2", "mg"],
    ["A3", "s2"],
    ["B3", 20],
    ["C3", "mg"],
    ["A4", "s3"],
    ["B4", 30],
    ["C4", "mg"],
  ]);

  function setUp(
    initialContent: Map<string, SerializedCellValue>,
    shape: { numRows: number; numCols: number }
  ) {
    engine.setSheetContent(data, initialContent);
    engine.addTable({
      tableName,
      ...data,
      start: "A1",
      numRows: { type: "number", value: shape.numRows },
      numCols: shape.numCols,
    });
    engine.setSheetContent(
      summary,
      new Map<string, SerializedCellValue>([
        ["A1", `=SUM(${tableName}[Score])`],
        ["A2", `=ROWS(${tableName}[#Data])`],
        ["A3", `=COLUMNS(${tableName}[#Data])`],
      ])
    );
    engine.clearUndoRedoHistory();
  }

  function inject(
    dataset: Map<string, SerializedCellValue>,
    shape: { numRows: number; numCols: number }
  ) {
    engine.transact(() => {
      // Resize first: the write then names every column of the new shape.
      engine.updateTable({
        tableName,
        workbookName,
        numRows: { type: "number", value: shape.numRows },
        numCols: shape.numCols,
      });
      engine.setSheetContent(data, dataset);
    });
  }

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet(data);
    engine.addSheet(summary);
    updates = 0;
    engine.onUpdate(() => {
      updates++;
    });
  });

  test("growing a table and writing a dataset is one update and one undo step", () => {
    setUp(twoColumns, { numRows: 2, numCols: 2 });
    expect([summaryValue("A1"), summaryValue("A2"), summaryValue("A3")]).toEqual(
      [3, 2, 2]
    );
    updates = 0;

    inject(threeColumns, { numRows: 3, numCols: 3 });

    expect(updates).toBe(1);
    expect(engine.getUndoRedoState()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
    });
    expect(headers()).toEqual(["Sample", "Value", "Unit"]);
    expect(table().endRow).toEqual({ type: "number", value: 3 });
    expect(content(data)).toEqual(Object.fromEntries(threeColumns));
    // The Score column was renamed to Value, and the reference followed it.
    expect(content(summary).A1).toBe(`=SUM(${tableName}[Value])`);
    expect([summaryValue("A1"), summaryValue("A2"), summaryValue("A3")]).toEqual(
      [60, 3, 3]
    );

    expect(engine.undo()).toBe(true);
    expect(updates).toBe(2);
    expect(engine.canUndo()).toBe(false);
    expect(headers()).toEqual(["Name", "Score"]);
    expect(table().endRow).toEqual({ type: "number", value: 2 });
    // No default name ("Column 3") is left behind from the resize.
    expect(content(data)).toEqual(Object.fromEntries(twoColumns));
    expect(content(summary).A1).toBe(`=SUM(${tableName}[Score])`);
    expect([summaryValue("A1"), summaryValue("A2"), summaryValue("A3")]).toEqual(
      [3, 2, 2]
    );

    expect(engine.redo()).toBe(true);
    expect(updates).toBe(3);
    expect(headers()).toEqual(["Sample", "Value", "Unit"]);
    expect(content(data)).toEqual(Object.fromEntries(threeColumns));
    expect([summaryValue("A1"), summaryValue("A2"), summaryValue("A3")]).toEqual(
      [60, 3, 3]
    );
  });

  test("shrinking a table and writing a dataset leaves nothing of the old shape", () => {
    setUp(threeColumns, { numRows: 3, numCols: 3 });
    updates = 0;

    inject(twoColumns, { numRows: 2, numCols: 2 });

    expect(updates).toBe(1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);
    expect(headers()).toEqual(["Name", "Score"]);
    expect(table().endRow).toEqual({ type: "number", value: 2 });
    expect(content(data)).toEqual(Object.fromEntries(twoColumns));
    expect([summaryValue("A2"), summaryValue("A3")]).toEqual([2, 2]);

    expect(engine.undo()).toBe(true);
    expect(headers()).toEqual(["Sample", "Value", "Unit"]);
    expect(table().endRow).toEqual({ type: "number", value: 3 });
    expect(content(data)).toEqual(Object.fromEntries(threeColumns));
    expect([summaryValue("A2"), summaryValue("A3")]).toEqual([3, 3]);
  });

  test("a dataset that reorders the table's columns keeps every column", () => {
    setUp(twoColumns, { numRows: 2, numCols: 2 });

    inject(
      new Map<string, SerializedCellValue>([
        ["A1", "Score"],
        ["B1", "Name"],
        ["C1", "Rank"],
        ["A2", 5],
        ["B2", "Cy"],
        ["C2", 1],
        ["A3", 7],
        ["B3", "Di"],
        ["C3", 2],
      ]),
      { numRows: 2, numCols: 3 }
    );

    expect(Array.from(table().headers.values())).toEqual([
      { name: "Score", index: 0 },
      { name: "Name", index: 1 },
      { name: "Rank", index: 2 },
    ]);
    expect(evaluate(`=SUM(${tableName}[Score])`)).toBe(12);
    expect(evaluate(`=SUM(${tableName}[Rank])`)).toBe(3);
    expect(evaluate(`=COUNTA(${tableName}[Name])`)).toBe(2);

    expect(engine.undo()).toBe(true);
    expect(Array.from(table().headers.values())).toEqual([
      { name: "Name", index: 0 },
      { name: "Score", index: 1 },
    ]);
    expect(evaluate(`=SUM(${tableName}[Score])`)).toBe(3);
  });

  test("a write that fails rolls back the resize", () => {
    setUp(twoColumns, { numRows: 2, numCols: 2 });
    const historyBefore = engine.getUndoRedoState();
    updates = 0;

    expect(() =>
      inject(
        new Map<string, SerializedCellValue>([
          ["A1", "Value"],
          ["B1", "Value"],
          ["C1", "Unit"],
        ]),
        { numRows: 2, numCols: 3 }
      )
    ).toThrow('Duplicate table header "Value"');

    expect(updates).toBe(0);
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
    expect(headers()).toEqual(["Name", "Score"]);
    expect(table().endRow).toEqual({ type: "number", value: 2 });
    expect(content(data)).toEqual(Object.fromEntries(twoColumns));
    expect([summaryValue("A1"), summaryValue("A3")]).toEqual([3, 2]);
  });

  test("creating a table and filling it is one undo step", () => {
    engine.clearUndoRedoHistory();
    updates = 0;

    engine.transact(() => {
      engine.addTable({
        tableName,
        ...data,
        start: "A1",
        numRows: { type: "number", value: 3 },
        numCols: 3,
      });
      engine.setSheetContent(data, threeColumns);
    });

    expect(updates).toBe(1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);
    expect(headers()).toEqual(["Sample", "Value", "Unit"]);
    expect(content(data)).toEqual(Object.fromEntries(threeColumns));
    expect(evaluate(`=SUM(${tableName}[Value])`)).toBe(60);

    expect(engine.undo()).toBe(true);
    expect(engine.hasTable({ tableName, workbookName })).toBe(false);
    expect(content(data)).toEqual({});

    expect(engine.redo()).toBe(true);
    expect(headers()).toEqual(["Sample", "Value", "Unit"]);
    expect(content(data)).toEqual(Object.fromEntries(threeColumns));
  });

  test("cloning a sheet and naming its table is one undo step", () => {
    setUp(twoColumns, { numRows: 2, numCols: 2 });
    updates = 0;

    engine.transact(() => {
      engine.cloneSheet({ ...data, newSheetName: "Data copy" });
      engine.renameTable(workbookName, {
        oldName: `${tableName}_2`,
        newName: "ResultsCopy",
      });
    });

    expect(updates).toBe(1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);
    expect(
      engine.getTable({ tableName: "ResultsCopy", workbookName })?.sheetName
    ).toBe("Data copy");
    expect(evaluate("=SUM(ResultsCopy[Score])")).toBe(3);

    expect(engine.undo()).toBe(true);
    expect(engine.hasSheet({ workbookName, sheetName: "Data copy" })).toBe(
      false
    );
    expect(engine.hasTable({ tableName: "ResultsCopy", workbookName })).toBe(
      false
    );
    expect(Array.from(engine.getTables(workbookName).keys())).toEqual([
      tableName,
    ]);
  });
});
