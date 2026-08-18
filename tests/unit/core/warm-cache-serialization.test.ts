import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { ENGINE_SNAPSHOT_VERSION } from "../../../src/core/engine-snapshot";
import { deserialize, serialize } from "../../../src/core/map-serializer";
import {
  cellAddressToKey,
  rangeAddressToKey,
  parseCellReference,
} from "../../../src/core/utils";
import { FormulaError } from "../../../src/core/types";
import { NO_TABLE_CONTEXT_NAME } from "../../../src/evaluator/evaluation-context";
import { AstEvaluationNode } from "../../../src/evaluator/dependency-nodes/ast-evaluation-node";
import { parseFormula } from "../../../src/parser/parser";

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
    expect(snapshot.version).toBe(ENGINE_SNAPSHOT_VERSION);
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

  test("roundtrips table-scoped current-row ASTs without mixing table and no-table contexts", () => {
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
          node.contextDependency?.sheetName === sheetName &&
          node.contextDependency?.workbookName === workbookName &&
          node.contextDependency?.rowIndex === 1 &&
          node.contextDependency?.tableName === "Sheet1Table"
      )
    ).toBe(true);
    expect(
      snapshot.managers.dependency.nodes.some(
        (node: any) =>
          node.kind === "ast" &&
          node.key === "ast:[@Identifier]" &&
          node.contextDependency?.sheetName === secondSheetName &&
          node.contextDependency?.workbookName === workbookName &&
          node.contextDependency?.rowIndex === 1 &&
          node.contextDependency?.tableName === NO_TABLE_CONTEXT_NAME
      )
    ).toBe(true);
    expect(
      snapshot.managers.dependency.nodes.some(
        (node: any) =>
          node.kind === "cell-value" &&
          node.key === cellAddressToKey(address("B2"))
      )
    ).toBe(true);
    expect(
      snapshot.managers.dependency.nodes.some(
        (node: any) =>
          node.kind === "cell-value" &&
          node.key === cellAddressToKey(secondSheetAddress("B2"))
      )
    ).toBe(true);

    const hydratedEngine = FormulaEngine.buildEmpty();
    hydratedEngine.resetToSerializedEngine(serialized);

    expect(hydratedEngine.getCellValue(address("B2"))).toBe("abc");
    expect(hydratedEngine.getCellValue(secondSheetAddress("B2"), true)).toBe(
      "#REF! in ast:[@Identifier] Table undefined not found"
    );
  });

  test("edits invalidate stale cache state before reserializing", () => {
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

    hydratedEngine.setCellContent(address("A1"), 20);
    expect(hydratedEngine.getCellValue(address("B1"))).toBe(21);

    const afterSecondEdit = FormulaEngine.buildEmpty();
    afterSecondEdit.resetToSerializedEngine(hydratedEngine.serializeEngine());
    expect(afterSecondEdit.getCellValue(address("B1"))).toBe(21);
  });

  test("resetting an engine replaces stale named-expression scopes before reserializing", () => {
    engine.addNamedExpression({
      expressionName: "LOCAL_RATE",
      expression: "0.2",
      sheetName,
      workbookName,
    });

    const dirtyEngine = FormulaEngine.buildEmpty();
    dirtyEngine.addWorkbook("a");
    dirtyEngine.addSheet({ workbookName: "a", sheetName: "Sheet1" });
    dirtyEngine.addNamedExpression({
      expressionName: "STALE_RATE",
      expression: "0.1",
      workbookName: "a",
    });

    dirtyEngine.resetToSerializedEngine(engine.serializeEngine());

    const snapshot = deserialize(dirtyEngine.serializeEngine()) as any;
    expect(snapshot.managers.namedExpression.workbookExpressions.has("a")).toBe(
      false
    );
    expect(snapshot.managers.namedExpression.sheetExpressions.has("a")).toBe(
      false
    );
    expect(
      dirtyEngine.hasNamedExpression({
        expressionName: "LOCAL_RATE",
        sheetName,
        workbookName,
      })
    ).toBe(true);
  });

  test("serializeEngine filters named-expression scopes to existing workbooks and sheets", () => {
    engine._namedExpressionManager.addNamedExpression({
      expressionName: "STALE_WORKBOOK",
      expression: "1",
      workbookName: "a",
    });
    engine._namedExpressionManager.addNamedExpression({
      expressionName: "STALE_SHEET",
      expression: "2",
      sheetName: "MissingSheet",
      workbookName,
    });

    const snapshot = deserialize(engine.serializeEngine()) as any;

    expect(snapshot.managers.namedExpression.workbookExpressions.has("a")).toBe(
      false
    );
    expect(
      snapshot.managers.namedExpression.sheetExpressions.has(workbookName)
    ).toBe(true);
    expect(
      snapshot.managers.namedExpression.sheetExpressions
        .get(workbookName)
        ?.has(sheetName)
    ).toBe(true);
    expect(
      snapshot.managers.namedExpression.sheetExpressions
        .get(workbookName)
        ?.has("MissingSheet")
    ).toBe(false);
  });

  test("ignores orphan named-expression scopes in serialized snapshots", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", "Value"],
        ["A2", 1],
      ])
    );
    engine.addTable({
      tableName: "Data",
      sheetName,
      workbookName,
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 1,
    });

    const snapshot = deserialize(engine.serializeEngine()) as any;
    snapshot.managers.namedExpression.workbookExpressions.set("a", new Map());
    snapshot.managers.namedExpression.sheetExpressions.set(
      "a",
      new Map([["Sheet1", new Map()]])
    );

    const hydratedEngine = FormulaEngine.buildEmpty();
    expect(() =>
      hydratedEngine.resetToSerializedEngine(serialize(snapshot))
    ).not.toThrow();
    expect(
      hydratedEngine.getState().tables.get(workbookName)?.has("Data")
    ).toBe(true);

    const cleanedSnapshot = deserialize(hydratedEngine.serializeEngine()) as any;
    expect(
      cleanedSnapshot.managers.namedExpression.workbookExpressions.has("a")
    ).toBe(false);
    expect(
      cleanedSnapshot.managers.namedExpression.sheetExpressions.has("a")
    ).toBe(false);
  });

  test("falls back to cold dependency state when an AST snapshot key is invalid", () => {
    engine.setCellContent(address("A1"), "=1+1");

    expect(engine.getCellValue(address("A1"))).toBe(2);

    const snapshot = deserialize(engine.serializeEngine()) as any;
    const astSnapshot = snapshot.managers.dependency.nodes.find(
      (node: any) => node.kind === "ast"
    );

    expect(astSnapshot).toBeDefined();
    astSnapshot.key = "ast:SUM(";
    astSnapshot.snapshotId = "ast:SUM(::{}";

    const originalConsoleWarn = console.warn;
    const warnings: unknown[][] = [];
    console.warn = (...args: unknown[]) => {
      warnings.push(args);
    };

    try {
      const hydratedEngine = FormulaEngine.buildEmpty();
      expect(() =>
        hydratedEngine.resetToSerializedEngine(serialize(snapshot))
      ).not.toThrow();
      expect(hydratedEngine.getCellValue(address("A1"))).toBe(2);
    } finally {
      console.warn = originalConsoleWarn;
    }

    expect(
      warnings.some(
        (entry) =>
          entry[0] === "[FormulaEngine snapshot]" &&
          String(entry[1]).includes("Failed to create warm snapshot node")
      )
    ).toBe(true);
  });

  test("serializes when an error references an equivalent AST snapshot identity", () => {
    engine.setCellContent(address("A1"), '="a"+1');

    expect(engine.getCellValue(address("A1"))).toBe(FormulaError.VALUE);

    const cellNode = engine._dependencyManager.getCellValueNode(
      cellAddressToKey(address("A1"))
    );
    const astEntries = engine._dependencyManager.asts.get('ast:"a"+1');
    const liveAstNode = astEntries
      ? Array.from(astEntries.entries.values())[0]?.evalNode
      : undefined;

    expect(liveAstNode).toBeDefined();

    const duplicateAstNode = new AstEvaluationNode(parseFormula('"a"+1'), {});
    duplicateAstNode.setEvaluationResult({
      type: "error",
      err: FormulaError.VALUE,
      errAddress: duplicateAstNode,
      message: "simulated equivalent stale AST reference",
    });
    duplicateAstNode.setContextDependency(liveAstNode!.getContextDependency());
    duplicateAstNode.resolve();

    cellNode.setEvaluationResult({
      type: "error",
      err: FormulaError.VALUE,
      errAddress: duplicateAstNode,
      message: "simulated equivalent stale AST reference",
    });

    expect(() => engine.serializeEngine()).not.toThrow();
  });

  test("rejects legacy serialized engine payloads", () => {
    engine.setCellContent(address("A1"), 1);

    const legacyPayload = serialize(engine.getState());
    const hydratedEngine = FormulaEngine.buildEmpty();

    expect(() => hydratedEngine.resetToSerializedEngine(legacyPayload)).toThrow(
      `Unsupported serialized engine format. Expected EngineSnapshot version ${ENGINE_SNAPSHOT_VERSION}.`
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
