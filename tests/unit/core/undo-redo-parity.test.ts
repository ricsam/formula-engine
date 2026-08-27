import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type {
  CellAddress,
  ConditionalStyle,
  DirectCellDataType,
  DirectCellStyle,
  NamedExpression,
  RangeAddress,
  SerializedCellValue,
  TableDefinition,
} from "../../../src/core/types";
import type { EngineSnapshot } from "../../../src/core/engine-snapshot";
import { deserialize } from "../../../src/core/map-serializer";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";

type TestMetadata = {
  cell: { label: string; nested?: { value: number } };
  sheet: { title: string; frozenRows?: number };
  workbook: { theme: string; revision?: number };
  range: { kind: string; label: string };
};

type TestEngine = FormulaEngine<TestMetadata>;
type PersistentManagers = Omit<
  EngineSnapshot["managers"],
  "dependency" | "cache"
>;

function buildEngine(): TestEngine {
  const engine = FormulaEngine.buildEmpty<TestMetadata>();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  engine.clearUndoRedoHistory();
  return engine;
}

function cell(
  ref: string,
  targetSheetName = sheetName,
  targetWorkbookName = workbookName
): CellAddress {
  return {
    workbookName: targetWorkbookName,
    sheetName: targetSheetName,
    ...parseCellReference(ref),
  };
}

function finiteRange(
  start: string,
  end = start,
  targetSheetName = sheetName,
  targetWorkbookName = workbookName
): RangeAddress {
  const startAddress = parseCellReference(start);
  const endAddress = parseCellReference(end);
  return {
    workbookName: targetWorkbookName,
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

function persistentManagers(engine: TestEngine): PersistentManagers {
  const snapshot = deserialize(engine.serializeEngine()) as EngineSnapshot;
  const {
    dependency: _dependency,
    cache: _cache,
    ...persistent
  } = snapshot.managers;
  return persistent;
}

/**
 * Verifies the complete persisted state, history shape, and notification count
 * around one public mutation. The engine is left in the redone state.
 */
function expectUndoRedoParity<TResult>(
  engine: TestEngine,
  mutate: () => TResult
): TResult {
  engine.clearUndoRedoHistory();
  const before = persistentManagers(engine);
  let updates = 0;
  const unsubscribe = engine.onUpdate(() => {
    updates++;
  });

  const result = mutate();
  const after = persistentManagers(engine);

  expect(after).not.toEqual(before);
  expect(engine.getUndoRedoState()).toMatchObject({
    canUndo: true,
    canRedo: false,
    undoDepth: 1,
    redoDepth: 0,
  });
  expect(updates).toBe(1);

  expect(engine.undo()).toBe(true);
  expect(persistentManagers(engine)).toEqual(before);
  expect(updates).toBe(2);

  expect(engine.redo()).toBe(true);
  expect(persistentManagers(engine)).toEqual(after);
  expect(updates).toBe(3);
  unsubscribe();

  return result;
}

function sourceAddresses(): CellAddress[] {
  return [cell("A1"), cell("B1"), cell("A2"), cell("B2")];
}

function seedRichCopyState(engine: TestEngine): void {
  engine.setSheetContent(
    { workbookName, sheetName },
    new Map<string, SerializedCellValue>([
      ["A1", 1],
      ["B1", 10],
      ["A2", "=A1+1"],
      ["B2", "=B1+1"],
    ])
  );
  engine.setCellMetadata(cell("A1"), { label: "source-a1" });
  engine.setCellMetadata(cell("B1"), { label: "source-b1" });
  engine.addCellStyle({
    areas: [finiteRange("A1", "B2")],
    style: { bold: true, backgroundColor: "#aabbcc" },
  });
  engine.addConditionalStyle({
    areas: [finiteRange("A1", "B2")],
    condition: {
      type: "formula",
      formula: "A1>0",
      color: { l: 70, c: 40, h: 120 },
    },
  });
  engine.addCellDataType({
    areas: [finiteRange("A1", "B2")],
    dataType: "number",
  });
  engine.addRangeMetadata({
    areas: [finiteRange("A1", "B2")],
    metadata: { kind: "validation", label: "source-range" },
  });
}

function seedBaselineTable(engine: TestEngine): void {
  const baselineSheet = "BaselineSheet";
  if (!engine.hasSheet({ workbookName, sheetName: baselineSheet })) {
    engine.addSheet({ workbookName, sheetName: baselineSheet });
  }
  engine.setSheetContent(
    { workbookName, sheetName: baselineSheet },
    new Map<string, SerializedCellValue>([
      ["A1", "Key"],
      ["A2", "baseline"],
    ])
  );
  if (!engine.hasTable({ workbookName, tableName: "BaselineTable" })) {
    engine.addTable({
      workbookName,
      sheetName: baselineSheet,
      tableName: "BaselineTable",
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 1,
    });
  }
}

describe("incremental history persistent-state parity", () => {
  test("cell, sheet, and workbook metadata add/update/delete", () => {
    const engine = buildEngine();

    expectUndoRedoParity(engine, () =>
      engine.setCellMetadata(cell("A1"), {
        label: "first",
        nested: { value: 1 },
      })
    );
    expect(engine.getCellMetadata(cell("A1"))).toEqual({
      label: "first",
      nested: { value: 1 },
    });

    expectUndoRedoParity(engine, () =>
      engine.setCellMetadata(cell("A1"), {
        label: "second",
        nested: { value: 2 },
      })
    );
    expectUndoRedoParity(engine, () =>
      engine.setCellMetadata(cell("A1"), undefined)
    );
    expect(engine.getCellMetadata(cell("A1"))).toBeUndefined();

    expectUndoRedoParity(engine, () =>
      engine.setSheetMetadata(
        { workbookName, sheetName },
        { title: "Dashboard", frozenRows: 2 }
      )
    );
    expect(engine.getSheetMetadata({ workbookName, sheetName })).toEqual({
      title: "Dashboard",
      frozenRows: 2,
    });

    expectUndoRedoParity(engine, () =>
      engine.setWorkbookMetadata(workbookName, {
        theme: "ocean",
        revision: 3,
      })
    );
    expect(engine.getWorkbookMetadata(workbookName)).toEqual({
      theme: "ocean",
      revision: 3,
    });
  });

  test("range metadata add/remove/clear preserves generated identity", () => {
    const engine = buildEngine();

    const generatedId = expectUndoRedoParity(engine, () =>
      engine.addRangeMetadata({
        areas: [finiteRange("A1", "C3")],
        metadata: { kind: "validation", label: "generated" },
      })
    );
    expect(engine.getAllRangeMetadata().map(({ id }) => id)).toContain(
      generatedId
    );

    expectUndoRedoParity(engine, () => engine.removeRangeMetadata(generatedId));
    expect(engine.getAllRangeMetadata()).toHaveLength(0);

    const splitId = engine.addRangeMetadata({
      id: "split-range",
      areas: [finiteRange("A1", "C3")],
      metadata: { kind: "annotation", label: "split me" },
    });
    expect(splitId).toBe("split-range");

    expectUndoRedoParity(engine, () =>
      engine.clearRangeMetadata(finiteRange("B2"))
    );
    expect(engine.getRangeMetadataForCell(cell("B2"))).toHaveLength(0);
    expect(engine.getRangeMetadataForCell(cell("A1"))).toHaveLength(1);
    expect(engine.getRangeMetadataForCell(cell("C3"))).toHaveLength(1);
  });

  test("named expression add/update/rename/remove and bulk set", () => {
    const engine = buildEngine();

    expectUndoRedoParity(engine, () =>
      engine.addNamedExpression({ expressionName: "RATE", expression: "0.1" })
    );
    engine.setCellContent(cell("A1"), "=RATE*100");
    engine.addNamedExpression({
      expressionName: "DOUBLE_RATE",
      expression: "RATE*2",
      workbookName,
    });
    expect(engine.getCellValue(cell("A1"))).toBe(10);

    expectUndoRedoParity(engine, () =>
      engine.updateNamedExpression({
        expressionName: "RATE",
        expression: "0.25",
      })
    );
    expect(engine.getCellValue(cell("A1"))).toBe(25);

    expectUndoRedoParity(engine, () =>
      engine.renameNamedExpression({
        expressionName: "RATE",
        newName: "TAX_RATE",
      })
    );
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).get("A1")
    ).toBe("=TAX_RATE*100");
    expect(engine.hasNamedExpression({ expressionName: "TAX_RATE" })).toBe(
      true
    );

    expectUndoRedoParity(engine, () =>
      engine.removeNamedExpression({ expressionName: "TAX_RATE" })
    );
    expect(engine.hasNamedExpression({ expressionName: "TAX_RATE" })).toBe(
      false
    );

    const expressions = new Map<string, NamedExpression>([
      ["ONE", { name: "ONE", expression: "1" }],
      ["TWO", { name: "TWO", expression: "2" }],
    ]);
    expectUndoRedoParity(engine, () =>
      engine.setNamedExpressions({ type: "global", expressions })
    );
    expect(engine.hasNamedExpression({ expressionName: "ONE" })).toBe(true);
    expect(engine.hasNamedExpression({ expressionName: "TWO" })).toBe(true);
  });

  test("table add/update/header edit/rename/remove/reset", () => {
    const engine = buildEngine();
    seedBaselineTable(engine);
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "Name"],
        ["B1", "Value"],
        ["A2", "alpha"],
        ["B2", 10],
        ["A3", "beta"],
        ["B3", 20],
      ])
    );

    expectUndoRedoParity(engine, () =>
      engine.addTable({
        workbookName,
        sheetName,
        tableName: "Data",
        start: "A1",
        numRows: { type: "number", value: 1 },
        numCols: 2,
      })
    );
    expect(engine.hasTable({ workbookName, tableName: "Data" })).toBe(true);

    engine.setCellContent(cell("D1"), "=SUM(Data[Value])");
    engine.addNamedExpression({
      expressionName: "DATA_TOTAL",
      expression: "SUM(Data[Value])",
    });
    expectUndoRedoParity(engine, () =>
      engine.setCellContent(cell("B1"), "Amount")
    );
    expect(
      Array.from(engine.getTable({ workbookName, tableName: "Data" })!.headers)
    ).toEqual([
      ["Name", { name: "Name", index: 0 }],
      ["Amount", { name: "Amount", index: 1 }],
    ]);
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).get("D1")
    ).toBe("=SUM(Data[Amount])");

    expectUndoRedoParity(engine, () =>
      engine.updateTable({
        workbookName,
        tableName: "Data",
        numRows: { type: "number", value: 2 },
      })
    );
    expect(
      engine.getTable({ workbookName, tableName: "Data" })?.endRow
    ).toEqual({ type: "number", value: 2 });

    expectUndoRedoParity(engine, () =>
      engine.renameTable(workbookName, {
        oldName: "Data",
        newName: "Measurements",
      })
    );
    expect(engine.hasTable({ workbookName, tableName: "Measurements" })).toBe(
      true
    );
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).get("D1")
    ).toBe("=SUM(Measurements[Amount])");

    expectUndoRedoParity(engine, () =>
      engine.removeTable({ workbookName, tableName: "Measurements" })
    );
    expect(engine.hasTable({ workbookName, tableName: "Measurements" })).toBe(
      false
    );

    const resetDefinition: TableDefinition = {
      name: "ResetTable",
      workbookName,
      sheetName,
      start: { colIndex: 0, rowIndex: 0 },
      endRow: { type: "number", value: 2 },
      headers: new Map([
        ["Name", { name: "Name", index: 0 }],
        ["Amount", { name: "Amount", index: 1 }],
      ]),
    };
    expectUndoRedoParity(engine, () =>
      engine.resetTables(
        new Map([[workbookName, new Map([["ResetTable", resetDefinition]])]])
      )
    );
    expect(engine.hasTable({ workbookName, tableName: "ResetTable" })).toBe(
      true
    );
  });

  test("conditional/direct styles and cell data types", () => {
    const engine = buildEngine();
    const conditional: ConditionalStyle = {
      areas: [finiteRange("A1", "C3")],
      condition: {
        type: "formula",
        formula: "A1>0",
        color: { l: 65, c: 50, h: 20 },
      },
    };
    const direct: DirectCellStyle = {
      areas: [finiteRange("A1", "C3")],
      style: { bold: true, backgroundColor: "#ffcc00" },
    };
    const dataType: DirectCellDataType = {
      areas: [finiteRange("A1", "C3")],
      dataType: "number",
    };

    expectUndoRedoParity(engine, () => engine.addConditionalStyle(conditional));
    expectUndoRedoParity(engine, () =>
      engine.removeConditionalStyle(workbookName, 0)
    );
    expect(engine.getAllConditionalStyles()).toHaveLength(0);

    expectUndoRedoParity(engine, () => engine.addCellStyle(direct));
    expectUndoRedoParity(engine, () => engine.removeCellStyle(workbookName, 0));
    expect(engine.getAllCellStyles()).toHaveLength(0);

    engine.addConditionalStyle(conditional);
    engine.addCellStyle(direct);
    expectUndoRedoParity(engine, () =>
      engine.clearCellStyles(finiteRange("B2"))
    );
    expect(engine.getCellStyle(cell("B2"))).toBeUndefined();
    expect(engine.getCellStyle(cell("A1"))).toBeDefined();

    expectUndoRedoParity(engine, () => engine.addCellDataType(dataType));
    expect(engine.getCellDataType(cell("A1"))).toBe("number");
    expectUndoRedoParity(engine, () =>
      engine.clearCellDataTypes(finiteRange("B2"))
    );
    expect(engine.getCellDataType(cell("B2"))).toBe("general");
    expect(engine.getCellDataType(cell("A1"))).toBe("number");
  });

  test("paste, fill, and autofill restore all copied state and IDs", () => {
    const engine = buildEngine();
    seedRichCopyState(engine);

    expectUndoRedoParity(engine, () =>
      engine.pasteCells(sourceAddresses(), cell("D1"), {
        cut: false,
        include: "all",
        type: "formula",
      })
    );
    expect(engine.getCellValue(cell("D1"))).toBe(1);
    expect(engine.getCellMetadata(cell("D1"))).toEqual({ label: "source-a1" });
    expect(engine.getCellStyle(cell("D1"))?.bold).toBe(true);
    expect(engine.getCellDataType(cell("D1"))).toBe("number");
    expect(engine.getRangeMetadataForCell(cell("D1"))).toHaveLength(1);

    expectUndoRedoParity(engine, () =>
      engine.fillAreas(finiteRange("A1", "B2"), [finiteRange("G1", "J4")], {
        cut: false,
        include: "all",
        type: "formula",
      })
    );
    expect(engine.getCellMetadata(cell("G1"))).toEqual({ label: "source-a1" });
    expect(engine.getCellStyle(cell("J4"))?.bold).toBe(true);
    expect(engine.getCellDataType(cell("J4"))).toBe("number");
    expect(engine.getRangeMetadataForCell(cell("J4"))).toHaveLength(1);

    expectUndoRedoParity(engine, () =>
      engine.autoFill(
        { workbookName, sheetName },
        finiteRange("A1", "B2").range,
        [finiteRange("A5", "B8").range],
        "down"
      )
    );
    expect(engine.getCellMetadata(cell("A5"))).toEqual({ label: "source-a1" });
    expect(engine.getCellStyle(cell("B8"))?.bold).toBe(true);
    expect(engine.getCellDataType(cell("B8"))).toBe("number");
    expect(engine.getRangeMetadataForCell(cell("B8"))).toHaveLength(1);
  });

  test("replace, replaceAll, and clearSpreadsheetRange", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "draft draft"],
        ["A2", "draft"],
        ["B1", '=IF(A2="draft","draft","")'],
        ["C1", 42],
      ])
    );

    expectUndoRedoParity(engine, () =>
      engine.replace("draft", "ready", {
        workbookName,
        sheetName,
        cellReference: "A1",
        occurrenceIndex: 1,
      })
    );
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).get("A1")
    ).toBe("draft ready");

    expectUndoRedoParity(engine, () =>
      engine.replaceAll("draft", "final", { workbookName, sheetName })
    );
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).get("A2")
    ).toBe("final");

    expectUndoRedoParity(engine, () =>
      engine.clearSpreadsheetRange(finiteRange("A1", "B1"))
    );
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).has("A1")
    ).toBe(false);
    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).has("B1")
    ).toBe(false);
    expect(engine.getCellValue(cell("C1"))).toBe(42);
  });

  test("sheet add/create/remove restores complete scope and generated name", () => {
    const engine = buildEngine();
    seedBaselineTable(engine);

    expectUndoRedoParity(engine, () =>
      engine.addSheet({ workbookName, sheetName: "Explicit" })
    );
    expect(engine.hasSheet({ workbookName, sheetName: "Explicit" })).toBe(true);

    const created = expectUndoRedoParity(engine, () =>
      engine.createSheet({ workbookName, baseName: "Generated" })
    );
    expect(created.name).toBe("Generated1");
    expect(engine.hasSheet({ workbookName, sheetName: created.name })).toBe(
      true
    );

    engine.setSheetContent(
      { workbookName, sheetName: "Explicit" },
      new Map<string, SerializedCellValue>([
        ["A1", "Header"],
        ["A2", 5],
      ])
    );
    engine.setCellMetadata(cell("A2", "Explicit"), { label: "cell metadata" });
    engine.setSheetMetadata(
      { workbookName, sheetName: "Explicit" },
      { title: "Explicit sheet", frozenRows: 1 }
    );
    engine.addNamedExpression({
      expressionName: "LOCAL",
      expression: "5",
      workbookName,
      sheetName: "Explicit",
    });
    engine.addCellStyle({
      areas: [finiteRange("A1", "A2", "Explicit")],
      style: { italic: true },
    });
    engine.addCellDataType({
      areas: [finiteRange("A2", "A2", "Explicit")],
      dataType: "number",
    });
    engine.addRangeMetadata({
      areas: [finiteRange("A1", "A2", "Explicit")],
      metadata: { kind: "note", label: "sheet range" },
    });
    engine.addTable({
      workbookName,
      sheetName: "Explicit",
      tableName: "ExplicitTable",
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 1,
    });
    const refId = engine.createRef(finiteRange("A1", "A2", "Explicit"));

    expectUndoRedoParity(engine, () =>
      engine.removeSheet({ workbookName, sheetName: "Explicit" })
    );
    expect(engine.hasSheet({ workbookName, sheetName: "Explicit" })).toBe(
      false
    );
    expect(engine.getRefAddress(refId)).toBeUndefined();
  });

  test("workbook add restores its empty ancillary scopes", () => {
    const engine = buildEngine();
    seedBaselineTable(engine);

    expectUndoRedoParity(engine, () => engine.addWorkbook("AddedBook"));
    expect(engine.hasWorkbook("AddedBook")).toBe(true);
  });

  test("workbook clone/remove parity independently of empty-workbook creation", () => {
    const engine = buildEngine();
    seedBaselineTable(engine);
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "Header"],
        ["A2", 9],
      ])
    );
    engine.setCellMetadata(cell("A2"), { label: "cloned cell" });
    engine.setSheetMetadata(
      { workbookName, sheetName },
      { title: "Source sheet" }
    );
    engine.setWorkbookMetadata(workbookName, { theme: "source" });
    engine.addNamedExpression({
      expressionName: "WORKBOOK_VALUE",
      expression: "9",
      workbookName,
    });
    engine.addNamedExpression({
      expressionName: "SHEET_VALUE",
      expression: "A2",
      workbookName,
      sheetName,
    });
    engine.addTable({
      workbookName,
      sheetName,
      tableName: "SourceTable",
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 1,
    });
    engine.addCellStyle({
      areas: [finiteRange("A1", "A2")],
      style: { underline: true },
    });
    engine.addConditionalStyle({
      areas: [finiteRange("A1", "A2")],
      condition: {
        type: "formula",
        formula: "A2>0",
        color: { l: 60, c: 30, h: 240 },
      },
    });
    engine.addCellDataType({
      areas: [finiteRange("A2")],
      dataType: "number",
    });
    engine.addRangeMetadata({
      id: "source-range-id",
      areas: [finiteRange("A1", "A2")],
      metadata: { kind: "source", label: "clone me" },
    });

    expectUndoRedoParity(engine, () =>
      engine.cloneWorkbook(workbookName, "IndependentClone")
    );
    const clonedRangeIds = engine
      .getAllRangeMetadata()
      .filter((entry) =>
        entry.areas.some((area) => area.workbookName === "IndependentClone")
      )
      .map(({ id }) => id);
    expect(clonedRangeIds).toHaveLength(1);
    expect(clonedRangeIds[0]).not.toBe("source-range-id");
    expect(engine.getWorkbookMetadata("IndependentClone")).toEqual({
      theme: "source",
    });

    const clonedRefId = engine.createRef(
      finiteRange("A1", "A2", sheetName, "IndependentClone")
    );
    expectUndoRedoParity(engine, () =>
      engine.removeWorkbook("IndependentClone")
    );
    expect(engine.hasWorkbook("IndependentClone")).toBe(false);
    expect(engine.getRefAddress(clonedRefId)).toBeUndefined();
  });
});
