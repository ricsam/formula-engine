import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import {
  ENGINE_SNAPSHOT_VERSION,
  LEGACY_ENGINE_SNAPSHOT_VERSION,
} from "../../../src/core/engine-snapshot";
import { deserialize, serialize } from "../../../src/core/map-serializer";
import type {
  CellAddress,
  CellDataType,
  RangeAddress,
  SpreadsheetRange,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";

function buildEngine() {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  engine.clearUndoRedoHistory();
  return engine;
}

function cell(ref: string, targetSheet = sheetName): CellAddress {
  return {
    workbookName,
    sheetName: targetSheet,
    ...parseCellReference(ref),
  };
}

function finiteRange(start: string, end = start, targetSheet = sheetName): RangeAddress {
  const startAddress = parseCellReference(start);
  const endAddress = parseCellReference(end);
  return {
    workbookName,
    sheetName: targetSheet,
    range: {
      start: { col: startAddress.colIndex, row: startAddress.rowIndex },
      end: {
        col: { type: "number", value: endAddress.colIndex },
        row: { type: "number", value: endAddress.rowIndex },
      },
    },
  };
}

function addType(
  engine: FormulaEngine,
  dataType: CellDataType,
  range: RangeAddress
) {
  engine.addCellDataType({ areas: [range], dataType });
}

describe("cell data types", () => {
  test("defaults to General and resolves newest rules and mixed ranges", () => {
    const engine = buildEngine();

    expect(engine.getCellDataType(cell("A1"))).toBe("general");
    expect(engine.getDataTypeForRange(finiteRange("A1", "E5"))).toBe(
      "general"
    );

    addType(engine, "text", finiteRange("A1", "E5"));
    addType(engine, "text", finiteRange("B2", "D4"));
    addType(engine, "general", finiteRange("C3"));

    expect(engine.getCellDataType(cell("A1"))).toBe("text");
    expect(engine.getCellDataType(cell("C3"))).toBe("general");
    expect(engine.getDataTypeForRange(finiteRange("A1", "B2"))).toBe("text");
    expect(engine.getDataTypeForRange(finiteRange("A1", "E5"))).toBeUndefined();
  });

  test("resolves uniform and mixed infinite ranges without enumerating cells", () => {
    const engine = buildEngine();
    const wholeColumn: SpreadsheetRange = {
      start: { col: 0, row: 0 },
      end: {
        col: { type: "number", value: 0 },
        row: { type: "infinity", sign: "positive" },
      },
    };
    addType(engine, "number", { workbookName, sheetName, range: wholeColumn });

    expect(
      engine.getDataTypeForRange({ workbookName, sheetName, range: wholeColumn })
    ).toBe("number");
    expect(
      engine.getDataTypeForRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "infinity", sign: "positive" },
          },
        },
      })
    ).toBeUndefined();
  });

  test("clears a hole while retaining the surrounding rule", () => {
    const engine = buildEngine();
    addType(engine, "boolean", finiteRange("A1", "E5"));

    engine.clearCellDataTypes(finiteRange("C3"));

    expect(engine.getCellDataType(cell("C3"))).toBe("general");
    expect(engine.getCellDataType(cell("C2"))).toBe("boolean");
    expect(engine.getCellDataType(cell("B3"))).toBe("boolean");
    expect(engine.getAllCellDataTypes()[0]?.areas).toHaveLength(4);
  });

  test("serializes version 6 and migrates version 5 to General", () => {
    const engine = buildEngine();
    addType(engine, "text", finiteRange("A1", "A2"));
    engine.setCellContent(cell("A1"), "001");

    const serialized = engine.serializeEngine();
    const snapshot = deserialize(serialized) as any;
    expect(snapshot.version).toBe(ENGINE_SNAPSHOT_VERSION);
    expect(snapshot.managers.style.cellDataTypes).toHaveLength(1);

    const restored = FormulaEngine.buildEmpty();
    restored.resetToSerializedEngine(serialized);
    expect(restored.getCellDataType(cell("A1"))).toBe("text");

    snapshot.version = LEGACY_ENGINE_SNAPSHOT_VERSION;
    delete snapshot.managers.style.cellDataTypes;
    const migrated = FormulaEngine.buildEmpty();
    migrated.resetToSerializedEngine(serialize(snapshot));
    expect(migrated.getCellDataType(cell("A1"))).toBe("general");

    snapshot.version = 4;
    expect(() =>
      FormulaEngine.buildEmpty().resetToSerializedEngine(serialize(snapshot))
    ).toThrow();
  });

  test("undoes a transacted type and content conversion together", () => {
    const engine = buildEngine();
    engine.setCellContent(cell("A1"), 12);
    engine.clearUndoRedoHistory();

    engine.transact(() => {
      addType(engine, "text", finiteRange("A1"));
      engine.setCellContent(cell("A1"), "12");
    });

    expect(engine.getCellDataType(cell("A1"))).toBe("text");
    expect(engine.getSheetSerialized({ workbookName, sheetName }).get("A1")).toBe(
      "12"
    );
    expect(engine.undo()).toBe(true);
    expect(engine.getCellDataType(cell("A1"))).toBe("general");
    expect(engine.getSheetSerialized({ workbookName, sheetName }).get("A1")).toBe(
      12
    );
    expect(engine.redo()).toBe(true);
    expect(engine.getCellDataType(cell("A1"))).toBe("text");
  });

  test("updates types for clone, rename, and deletion lifecycles", () => {
    const engine = buildEngine();
    engine.addSheet({ workbookName, sheetName: "Second" });
    engine.addWorkbook("Other");
    engine.addSheet({ workbookName: "Other", sheetName });
    engine.addCellDataType({
      dataType: "boolean",
      areas: [
        finiteRange("A1"),
        finiteRange("B2", "B2", "Second"),
        { ...finiteRange("C3"), workbookName: "Other" },
      ],
    });

    engine.cloneWorkbook(workbookName, "Copy");
    expect(
      engine.getCellDataType({
        workbookName: "Copy",
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      })
    ).toBe("boolean");
    const clonedRule = engine
      .getAllCellDataTypes()
      .find((rule) => rule.areas.some((area) => area.workbookName === "Copy"));
    expect(clonedRule?.areas.every((area) => area.workbookName === "Copy")).toBe(
      true
    );

    engine.renameSheet({ workbookName, sheetName, newSheetName: "Renamed" });
    engine.renameWorkbook({ workbookName, newWorkbookName: "RenamedBook" });
    expect(
      engine.getCellDataType({
        workbookName: "RenamedBook",
        sheetName: "Renamed",
        colIndex: 0,
        rowIndex: 0,
      })
    ).toBe("boolean");

    engine.removeSheet({
      workbookName: "RenamedBook",
      sheetName: "Renamed",
    });
    expect(
      engine.getCellDataType({
        workbookName: "RenamedBook",
        sheetName: "Second",
        colIndex: 1,
        rowIndex: 1,
      })
    ).toBe("boolean");

    engine.removeWorkbook("RenamedBook");
    expect(
      engine.getCellDataType({
        workbookName: "Other",
        sheetName,
        colIndex: 2,
        rowIndex: 2,
      })
    ).toBe("boolean");
  });

  test("copies types with formatting while content-only paste leaves them alone", () => {
    const engine = buildEngine();
    engine.setCellContent(cell("A1"), "source");
    addType(engine, "text", finiteRange("A1"));
    addType(engine, "boolean", finiteRange("B1"));

    engine.pasteCells([cell("A1")], cell("B1"), {
      include: ["content"],
      type: "formula",
    });
    expect(engine.getCellDataType(cell("B1"))).toBe("boolean");

    engine.pasteCells([cell("A1")], cell("B1"), {
      include: ["style"],
      type: "formula",
    });
    expect(engine.getCellDataType(cell("B1"))).toBe("text");

    addType(engine, "boolean", finiteRange("D1"));
    engine.pasteCells([cell("C1")], cell("D1"), {
      include: ["style"],
      type: "formula",
    });
    expect(engine.getCellDataType(cell("D1"))).toBe("general");

    engine.addSheet({ workbookName, sheetName: "Sheet2" });
    addType(engine, "boolean", finiteRange("A1", "A1", "Sheet2"));
    engine.pasteCells([cell("A1")], cell("A1", "Sheet2"), {
      include: ["style"],
      type: "formula",
    });
    expect(engine.getCellDataType(cell("A1", "Sheet2"))).toBe("text");
  });

  test("preserves false and zero content while copying typed cells", () => {
    const engine = buildEngine();
    engine.setCellContent(cell("A1"), 0);
    engine.setCellContent(cell("B1"), false);
    addType(engine, "number", finiteRange("A1"));
    addType(engine, "boolean", finiteRange("B1"));

    engine.pasteCells([cell("A1"), cell("B1")], cell("C1"), {
      include: "all",
      type: "formula",
    });

    const content = engine.getSheetSerialized({ workbookName, sheetName });
    expect(content.get("C1")).toBe(0);
    expect(content.get("D1")).toBe(false);
    expect(engine.getCellDataType(cell("C1"))).toBe("number");
    expect(engine.getCellDataType(cell("D1"))).toBe("boolean");
  });

  test("moves a type on cut and repeats type patterns for fill and autofill", () => {
    const engine = buildEngine();
    engine.setCellContent(cell("A1"), "one");
    engine.setCellContent(cell("B1"), "two");
    addType(engine, "text", finiteRange("A1"));
    addType(engine, "boolean", finiteRange("B1"));

    engine.pasteCells([cell("A1")], cell("E1"), {
      cut: true,
      include: "all",
      type: "formula",
    });
    expect(engine.getCellDataType(cell("A1"))).toBe("general");
    expect(engine.getCellDataType(cell("E1"))).toBe("text");

    engine.fillAreas(finiteRange("E1", "F1"), [finiteRange("E3", "H3")], {
      include: ["style"],
      type: "formula",
    });
    expect(engine.getCellDataType(cell("E3"))).toBe("text");
    expect(engine.getCellDataType(cell("F3"))).toBe("general");
    expect(engine.getCellDataType(cell("G3"))).toBe("text");
    expect(engine.getCellDataType(cell("H3"))).toBe("general");

    addType(engine, "boolean", finiteRange("F4", "H4"));
    engine.smartPaste(
      [cell("E1"), cell("F1")],
      {
        workbookName,
        sheetName,
        areas: [finiteRange("E4", "H4").range],
      },
      { include: ["style"], type: "formula" }
    );
    expect(engine.getCellDataType(cell("E4"))).toBe("text");
    expect(engine.getCellDataType(cell("F4"))).toBe("general");
    expect(engine.getCellDataType(cell("G4"))).toBe("text");
    expect(engine.getCellDataType(cell("H4"))).toBe("general");

    engine.autoFill(
      { workbookName, sheetName },
      finiteRange("A1", "B1").range,
      [finiteRange("A2", "B2").range],
      "down"
    );
    expect(engine.getCellDataType(cell("A2"))).toBe("general");
    expect(engine.getCellDataType(cell("B2"))).toBe("boolean");
  });
});
