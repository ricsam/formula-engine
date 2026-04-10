import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { deserialize, serialize } from "../../../src/core/map-serializer";
import { cellAddressToKey, rangeAddressToKey, parseCellReference } from "../../../src/core/utils";

const workbookName = "TestWorkbook";
const sheetName = "TestSheet";

function address(ref: string) {
  return {
    workbookName,
    sheetName,
    ...parseCellReference(ref),
  };
}

function buildEngine() {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  return engine;
}

describe("Warm-cache serialization", () => {
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = buildEngine();
  });

  test("roundtrips scalar formula values and cache metadata", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", 5],
        ["A2", 7],
        ["B1", "=A1+A2"],
      ])
    );

    expect(engine.getCellValue(address("B1"))).toBe(12);

    const snapshot = deserialize(engine.serializeEngine()) as any;
    expect(snapshot.version).toBe(3);
    expect(snapshot.managers.cache.evaluationOrders.length).toBeGreaterThan(0);
    expect(
      snapshot.managers.dependency.nodes.some((node: any) => node.kind === "ast")
    ).toBe(true);

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(engine.serializeEngine());

    expect(hydratedEngine.getCellValue(address("B1"))).toBe(12);
    expect(
      hydratedEngine.getCellEvaluationResult(address("B1"))
    ).toMatchObject({
      type: "value",
      result: { type: "number", value: 12 },
    });
  });

  test("roundtrips resolved blank frontier cells", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", 1],
        ["B1", 2],
        ["C1", "=SUM(A1:B1)"],
      ])
    );

    expect(engine.getCellValue(address("D1"))).toBe("");

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(engine.serializeEngine());

    const node = hydratedEngine._dependencyManager.getCellValueOrEmptyCellNode(
      cellAddressToKey(address("D1"))
    );

    expect(node.resolved).toBe(true);
    expect(hydratedEngine.getCellValue(address("D1"))).toBe("");
  });

  test("roundtrips finite spill values and spill registry", () => {
    engine.setCellContent(address("A1"), "=SEQUENCE(2,2)");

    expect(engine.getCellValue(address("A1"))).toBe(1);
    expect(engine.getCellValue(address("B1"))).toBe(2);
    expect(engine.getCellValue(address("A2"))).toBe(3);
    expect(engine.getCellValue(address("B2"))).toBe(4);

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(engine.serializeEngine());

    expect(hydratedEngine.getCellValue(address("A1"))).toBe(1);
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(2);
    expect(hydratedEngine.getCellValue(address("A2"))).toBe(3);
    expect(hydratedEngine.getCellValue(address("B2"))).toBe(4);
  });

  test("clearing a warmed formula removes orphaned AST snapshots", () => {
    engine.setCellContent(address("A1"), "=1+1");

    expect(engine.getCellValue(address("A1"))).toBe(2);

    engine.setCellContent(address("A1"), undefined);

    const snapshot = deserialize(engine.serializeEngine()) as any;

    expect(
      snapshot.managers.dependency.nodes.filter(
        (node: any) => node.kind === "ast"
      )
    ).toEqual([]);
  });

  test("roundtrips open-ended range consumers that were already hot", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["B1", "=SUM(A:A)"],
      ])
    );

    expect(engine.getCellValue(address("B1"))).toBe(60);

    const rangeNode = engine._dependencyManager.getRangeNode(
      rangeAddressToKey({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 0 },
            row: { type: "infinity", sign: "positive" },
          },
        },
      })
    );

    expect(rangeNode.resolved).toBe(true);

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(engine.serializeEngine());

    expect(hydratedEngine.getCellValue(address("B1"))).toBe(60);

    const hydratedRangeNode = hydratedEngine._dependencyManager.getRangeNode(
      rangeAddressToKey({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 0 },
            row: { type: "infinity", sign: "positive" },
          },
        },
      })
    );
    expect(hydratedRangeNode.resolved).toBe(true);
  });

  test("roundtrips table-scoped current-row ASTs", () => {
    const secondSheetName = "Sheet2";
    engine.addSheet({ workbookName, sheetName: secondSheetName });

    const sharedContent = new Map<string, string | number>([
      ["A1", "Identifier"],
      ["B1", "Calc"],
      ["A2", "abc"],
      ["B2", "=[@Identifier]"],
    ]);

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map(sharedContent)
    );
    engine.setSheetContent(
      { workbookName, sheetName: secondSheetName },
      new Map(sharedContent)
    );

    const secondSheetAddress = (ref: string) => ({
      workbookName,
      sheetName: secondSheetName,
      ...parseCellReference(ref),
    });

    expect(engine.getCellValue(address("B2"), true)).toBe(
      "#REF! in ast:[@Identifier] Table undefined not found"
    );
    expect(engine.getCellValue(secondSheetAddress("B2"), true)).toBe(
      "#REF! in ast:[@Identifier] Table undefined not found"
    );

    engine.addTable({
      tableName: "Sheet1Table",
      sheetName,
      workbookName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 2,
    });

    expect(engine.getCellValue(address("B2"))).toBe("abc");
    expect(engine.getCellValue(secondSheetAddress("B2"), true)).toBe(
      "#REF! in ast:[@Identifier] Table undefined not found"
    );

    const serialized = engine.serializeEngine();
    const snapshot = deserialize(serialized) as any;

    expect(
      snapshot.managers.dependency.nodes.some(
        (node: any) =>
          node.kind === "ast" &&
          node.key === "ast:[@Identifier]" &&
          node.contextDependency?.workbookName === workbookName &&
          node.contextDependency?.rowIndex === 1 &&
          node.contextDependency?.tableName === "Sheet1Table"
      )
    ).toBe(true);

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(serialized);

    expect(hydratedEngine.getCellValue(address("B2"))).toBe("abc");
    expect(hydratedEngine.getCellValue(secondSheetAddress("B2"), true)).toBe(
      "#REF! in ast:[@Identifier] Table undefined not found"
    );
  });

  test("edits, undo, and redo invalidate stale cache state before reserializing", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", 1],
        ["B1", "=A1+1"],
      ])
    );

    expect(engine.getCellValue(address("B1"))).toBe(2);

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(engine.serializeEngine());
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(2);

    hydratedEngine.setCellContent(address("A1"), 10);
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(11);

    const afterEdit = FormulaEngine.buildEmpty();
    afterEdit.resetToSerializedEngine(hydratedEngine.serializeEngine());
    expect(afterEdit.getCellValue(address("B1"))).toBe(11);

    expect(hydratedEngine.undo()).toBe(true);
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(2);

    const afterUndo = FormulaEngine.buildEmpty();
    afterUndo.resetToSerializedEngine(hydratedEngine.serializeEngine());
    expect(afterUndo.getCellValue(address("B1"))).toBe(2);

    expect(hydratedEngine.redo()).toBe(true);
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(11);
  });

  test("rejects legacy serialized engine payloads", () => {
    engine.setCellContent(address("A1"), 1);

    const legacyPayload = serialize(engine.getState());
    const hydratedEngine = FormulaEngine.buildEmpty();

    expect(() => hydratedEngine.resetToSerializedEngine(legacyPayload)).toThrow(
      "Unsupported serialized engine format. Expected EngineSnapshot version 3."
    );
  });

  test("tolerates dangling snapshot node ids from older warm-cache saves", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", 5],
        ["B1", "=A1+1"],
      ])
    );

    expect(engine.getCellValue(address("B1"))).toBe(6);

    const snapshot = deserialize(engine.serializeEngine()) as any;
    const danglingNodeId =
      'ast:DataTable[Condition]&DataTable[Identifier]::{"workbookName":"Responder Compound Pulse Workbook.2"}';

    const firstNode = snapshot.managers.dependency.nodes.find(
      (node: any) => node.kind === "cell-value"
    );
    firstNode.dependencies.push(danglingNodeId);

    const firstEvaluationOrder = snapshot.managers.cache.evaluationOrders[0];
    firstEvaluationOrder.evaluationOrder.push(danglingNodeId);

    const hydratedEngine = FormulaEngine.buildEmpty();
    expect(() =>
      hydratedEngine.resetToSerializedEngine(serialize(snapshot))
    ).not.toThrow();
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(6);
  });
});
