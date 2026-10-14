import { describe, expect, spyOn, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { deserialize, serialize } from "../../../src/core/map-serializer";
import { parseCellReference } from "../../../src/core/utils";
import type { RangeAddress } from "../../../src/core/types";

const sheet = { workbookName: "Book", sheetName: "Data" };
const address = (ref: string) => ({ ...sheet, ...parseCellReference(ref) });
const area: RangeAddress = {
  ...sheet,
  range: { start: { row: 0, col: 0 }, end: {
    row: { type: "number", value: 2 }, col: { type: "number", value: 2 },
  } },
};
function create() {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(sheet);
  engine.addSheet(sheet);
  return engine;
}

describe("compact canonical persistence", () => {
  test("skips dependency snapshot traversal and preserves every canonical manager", () => {
    const engine = create();
    engine.setSheetContent(sheet, new Map<string, string | number>([
      ["A1", "Value"], ["A2", 10], ["A3", 20],
      ["B1", "=SUM(A2:A3)*RATE"], ["D1", "=SEQUENCE(2,2)"],
    ]));
    engine.addNamedExpression({ ...sheet, expressionName: "RATE", expression: "2" });
    engine.addTable({ ...sheet, tableName: "Values", start: "A1", numRows: { type: "number", value: 2 }, numCols: 1 });
    engine.addCellStyle({ areas: [area], style: { color: "#123456", wrapText: true, fontFamily: "Arial, sans-serif" } });
    engine.addRangeMetadata({ areas: [area], metadata: { purpose: "synthetic" } });
    engine.setWorkbookMetadata(sheet.workbookName, { defaultCellStyle: { color: "#222222" } });
    engine.setSheetMetadata(sheet, { hidden: false });
    engine.setCellMetadata(address("A2"), { note: "retained" });
    engine.addCellDataType({ areas: [area], dataType: "number" });
    engine.addConditionalStyle({ areas: [area], condition: {
      type: "formula", formula: "=A2>5", color: { l: 50, c: 20, h: 120 },
    } });
    const reference = engine.createRef(area);
    expect(engine.getCellValue(address("B1"))).toBe(60);
    expect(engine.getCellValue(address("D1"))).toBe(1);
    const warm = deserialize(engine.serializeEngine()) as any;
    const snapshotSpy = spyOn(engine._dependencyManager, "toSnapshot");
    const compact = engine.serializeEngine({ includeEvaluationState: false });
    expect(snapshotSpy).not.toHaveBeenCalled();
    snapshotSpy.mockRestore();
    const cold = deserialize(compact) as any;
    const { dependency, cache, ...canonical } = warm.managers;
    expect(cold.managers).toEqual(canonical);
    expect(cold.managers).not.toHaveProperty("dependency");
    expect(cold.managers).not.toHaveProperty("cache");

    const loaded = create();
    loaded.setCellContent(address("B1"), "=999");
    expect(loaded.getCellValue(address("B1"))).toBe(999);
    loaded.resetToSerializedEngine(compact);
    expect(loaded.getState()).toEqual(engine.getState());
    expect(loaded.getRefAddress(reference)).toEqual(area);
    expect(loaded.getCellStyle(address("A2"))?.fontFamily).toBe("Arial, sans-serif");
    expect(loaded.getCellValue(address("B1"))).toBe(60);
    expect(loaded.getCellValue(address("E2"))).toBe(4);
    loaded.setCellContent(address("A2"), 15);
    expect(loaded.getCellValue(address("B1"))).toBe(70);
    loaded.undo();
    expect(loaded.getCellValue(address("B1"))).toBe(60);
    loaded.redo();
    expect(loaded.getCellValue(address("B1"))).toBe(70);
  });

  test("canonical bytes do not grow as evaluation caches warm", () => {
    const engine = create();
    const cells = new Map<string, number | string>();
    for (let i = 1; i <= 200; i++) {
      cells.set(`A${i}`, i);
      cells.set(`B${i}`, `=SUM(A1:A${i})`);
    }
    engine.setSheetContent(sheet, cells);
    const before = engine.serializeEngine({ includeEvaluationState: false });
    for (let i = 1; i <= 200; i++) engine.getCellValue(address(`B${i}`));
    expect(engine.serializeEngine({ includeEvaluationState: false })).toBe(before);
    expect(engine.serializeEngine().length).toBeGreaterThan(before.length * 5);
  });

  test("continues importing existing version 5 and 6 warm snapshots", () => {
    const source = create();
    source.setCellContent(address("A1"), "=1+2");
    expect(source.getCellValue(address("A1"))).toBe(3);
    for (const version of [5, 6]) {
      const snapshot = deserialize(source.serializeEngine()) as any;
      snapshot.version = version;
      if (version === 5) delete snapshot.managers.style.cellDataTypes;
      const loaded = FormulaEngine.buildEmpty();
      loaded.resetToSerializedEngine(serialize(snapshot));
      expect(loaded.getCellValue(address("A1"))).toBe(3);
    }
  });

  test("rejects partial warm state before replacing current workbook data", () => {
    const engine = create();
    engine.setCellContent(address("A1"), 42);
    const snapshot = deserialize(engine.serializeEngine()) as any;
    delete snapshot.managers.cache;
    expect(() => engine.resetToSerializedEngine(serialize(snapshot))).toThrow("both dependency and cache");
    expect(engine.getCellValue(address("A1"))).toBe(42);
  });
});
