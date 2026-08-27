import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type {
  CellAddress,
  FormulaEngineOptions,
  RangeAddress,
  SerializedCellValue,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";

function address(
  ref: string,
  targetWorkbookName = workbookName,
  targetSheetName = sheetName
): CellAddress {
  return {
    workbookName: targetWorkbookName,
    sheetName: targetSheetName,
    ...parseCellReference(ref),
  };
}

function finiteRange(start: string, end = start): RangeAddress {
  const startAddress = parseCellReference(start);
  const endAddress = parseCellReference(end);
  return {
    workbookName,
    sheetName,
    range: {
      start: { col: startAddress.colIndex, row: startAddress.rowIndex },
      end: {
        col: { type: "number", value: endAddress.colIndex },
        row: { type: "number", value: endAddress.rowIndex },
      },
    },
  };
}

function buildEngine(options?: FormulaEngineOptions): FormulaEngine {
  const engine = FormulaEngine.buildEmpty(options);
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  engine.clearUndoRedoHistory();
  return engine;
}

describe("FormulaEngine incremental history", () => {
  test("uses bounded incremental-history defaults", () => {
    const engine = FormulaEngine.buildEmpty();

    expect(engine.getUndoRedoState()).toEqual({
      enabled: true,
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      maxEntries: 100,
      maxBytes: 64 * 1024 * 1024,
      undoBytes: 0,
      redoBytes: 0,
    });
  });

  test("1,000 cell edits stay bounded with a large unrelated workbook", () => {
    const smallEngine = buildEngine();
    smallEngine.setCellContent(address("A1"), "before");
    smallEngine.clearUndoRedoHistory();

    const largeEngine = buildEngine();
    largeEngine.setCellContent(address("A1"), "before");
    largeEngine.addWorkbook("UnrelatedBook");
    largeEngine.addSheet({
      workbookName: "UnrelatedBook",
      sheetName: "LargeSheet",
    });

    const unrelatedContent = new Map<string, SerializedCellValue>();
    for (let row = 1; row <= 10_000; row++) {
      unrelatedContent.set(`A${row}`, `unrelated-${row}-${"x".repeat(64)}`);
    }
    largeEngine.setSheetContent(
      { workbookName: "UnrelatedBook", sheetName: "LargeSheet" },
      unrelatedContent
    );
    largeEngine.clearUndoRedoHistory();

    for (let rowIndex = 1; rowIndex <= 1_000; rowIndex++) {
      const editedAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex,
      };
      smallEngine.setCellContent(editedAddress, `after-${rowIndex}`);
      largeEngine.setCellContent(editedAddress, `after-${rowIndex}`);
    }

    const smallState = smallEngine.getUndoRedoState();
    const largeState = largeEngine.getUndoRedoState();

    expect(smallState.undoDepth).toBe(100);
    expect(largeState.undoDepth).toBe(100);
    expect(smallState.undoBytes).toBeGreaterThan(0);
    expect(largeState.undoBytes).toBe(smallState.undoBytes);
  });

  test("a style edit retains no unrelated style state", () => {
    const styleAtRow = (rowIndex: number) => ({
      areas: [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: rowIndex },
            end: {
              col: { type: "number" as const, value: 0 },
              row: { type: "number" as const, value: rowIndex },
            },
          },
        },
      ],
      style: { bold: true },
    });

    const smallEngine = buildEngine();
    smallEngine.addCellStyle(styleAtRow(20_001));
    const smallEditBytes = smallEngine.getUndoRedoState().undoBytes;

    const largeEngine = buildEngine();
    for (let rowIndex = 0; rowIndex < 20_000; rowIndex++) {
      // Seed outside an engine transaction so this setup is not retained in
      // history; the edit below must journal only its own sparse delta.
      largeEngine._styleManager.addCellStyle(styleAtRow(rowIndex));
    }
    largeEngine.clearUndoRedoHistory();
    largeEngine.addCellStyle(styleAtRow(20_001));

    expect(largeEngine.getUndoRedoState()).toMatchObject({
      undoDepth: 1,
      undoBytes: smallEditBytes,
    });
    expect(largeEngine.undo()).toBe(true);
    expect(largeEngine._styleManager.getAllCellStyles()).toHaveLength(20_000);
    expect(largeEngine.redo()).toBe(true);
    expect(largeEngine._styleManager.getAllCellStyles()).toHaveLength(20_001);
  });

  test("cut style patches replay in their original sequential order", () => {
    const engine = buildEngine();
    engine.addCellStyle({
      areas: [finiteRange("A1", "A2")],
      style: { bold: true },
    });
    engine.clearUndoRedoHistory();
    const before = structuredClone(engine.getAllCellStyles());

    engine.pasteCells([address("A1"), address("A2")], address("C1"), {
      cut: true,
      type: "formula",
      include: "all",
    });
    const after = structuredClone(engine.getAllCellStyles());

    expect(engine.undo()).toBe(true);
    expect(engine.getAllCellStyles()).toEqual(before);
    expect(engine.redo()).toBe(true);
    expect(engine.getAllCellStyles()).toEqual(after);
  });

  test("large indexed patches remain atomic across undo and redo", () => {
    const engine = buildEngine();
    for (let index = 0; index < 1_030; index++) {
      engine._styleManager.addCellStyle({
        areas: [finiteRange("A1")],
        style: { fontSize: 10 + (index % 5) },
      });
    }
    engine.clearUndoRedoHistory();

    engine.clearCellStyles(finiteRange("A1"));
    expect(engine.getAllCellStyles()).toHaveLength(0);
    expect(engine.undo()).toBe(true);
    expect(engine.getAllCellStyles()).toHaveLength(1_030);
    expect(engine.redo()).toBe(true);
    expect(engine.getAllCellStyles()).toHaveLength(0);
  });

  test("restores exact cell insertion order for sequential deletions", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "a"],
        ["B1", "b"],
        ["C1", "c"],
      ])
    );
    engine.clearUndoRedoHistory();

    engine.transact(() => {
      engine.setCellContent(address("A1"), undefined);
      engine.setCellContent(address("B1"), undefined);
    });
    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual([["C1", "c"]]);

    expect(engine.undo()).toBe(true);
    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual([
      ["A1", "a"],
      ["B1", "b"],
      ["C1", "c"],
    ]);
    expect(engine.redo()).toBe(true);
    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual([["C1", "c"]]);
  });

  test("setSheetContent records reorder-only changes and restores exact order", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "a"],
        ["B1", "b"],
      ])
    );
    engine.clearUndoRedoHistory();

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["B1", "b"],
        ["A1", "a"],
      ])
    );
    expect(engine.getUndoRedoState().canUndo).toBe(true);
    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }).keys())
    ).toEqual(["B1", "A1"]);

    expect(engine.undo()).toBe(true);
    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }).keys())
    ).toEqual(["A1", "B1"]);
    expect(engine.redo()).toBe(true);
    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }).keys())
    ).toEqual(["B1", "A1"]);
  });

  test("restores exact cell-metadata insertion order", () => {
    const engine = buildEngine();
    engine.setCellMetadata(address("A1"), { label: "a" });
    engine.setCellMetadata(address("B1"), { label: "b" });
    engine.setCellMetadata(address("C1"), { label: "c" });
    engine.clearUndoRedoHistory();

    engine.transact(() => {
      engine.setCellMetadata(address("A1"), undefined);
      engine.setCellMetadata(address("B1"), undefined);
    });
    const metadataKeys = () =>
      Array.from(
        engine._workbookManager
          .getSheetMetadataSerialized({ workbookName, sheetName })
          .keys()
      );
    expect(metadataKeys()).toEqual(["C1"]);
    expect(engine.undo()).toBe(true);
    expect(metadataKeys()).toEqual(["A1", "B1", "C1"]);
    expect(engine.redo()).toBe(true);
    expect(metadataKeys()).toEqual(["C1"]);
  });

  test("records style precedence changes when existing rules are reordered", () => {
    const engine = buildEngine();
    const first = {
      areas: [finiteRange("A1")],
      style: { backgroundColor: "#ff0000" },
    };
    const second = {
      areas: [finiteRange("A1")],
      style: { backgroundColor: "#0000ff" },
    };
    engine.addCellStyle(first);
    engine.addCellStyle(second);
    engine.clearUndoRedoHistory();

    engine.transact(() => {
      engine.removeCellStyle(workbookName, 0);
      engine.addCellStyle(first);
    });
    expect(engine.getAllCellStyles()).toEqual([second, first]);
    expect(engine.undo()).toBe(true);
    expect(engine.getAllCellStyles()).toEqual([first, second]);
    expect(engine.redo()).toBe(true);
    expect(engine.getAllCellStyles()).toEqual([second, first]);
  });

  test("coalesces repeated edits in one transaction and emits once", () => {
    const singleEditEngine = buildEngine();
    singleEditEngine.setCellContent(address("A1"), "before");
    singleEditEngine.clearUndoRedoHistory();
    singleEditEngine.setCellContent(address("A1"), "final");
    const singleEditBytes = singleEditEngine.getUndoRedoState().undoBytes;

    const transactedEngine = buildEngine();
    transactedEngine.setCellContent(address("A1"), "before");
    transactedEngine.clearUndoRedoHistory();

    let updates = 0;
    const unsubscribe = transactedEngine.onUpdate(() => {
      updates++;
    });

    transactedEngine.transact(() => {
      for (let index = 0; index < 1_000; index++) {
        transactedEngine.setCellContent(address("A1"), `intermediate-${index}`);
      }
      transactedEngine.setCellContent(address("A1"), "final");
    });
    unsubscribe();

    const state = transactedEngine.getUndoRedoState();
    expect(updates).toBe(1);
    expect(state.undoDepth).toBe(1);
    expect(state.undoBytes).toBeGreaterThan(0);
    expect(state.undoBytes).toBeLessThanOrEqual(singleEditBytes * 2 + 256);

    expect(transactedEngine.undo()).toBe(true);
    expect(transactedEngine.getCellValue(address("A1"))).toBe("before");
    expect(transactedEngine.redo()).toBe(true);
    expect(transactedEngine.getCellValue(address("A1"))).toBe("final");
  });

  test("does not retain a net-zero ancillary transaction", () => {
    const engine = buildEngine();
    engine.setCellContent(address("A1"), "retained-before-no-op");
    const before = engine.getUndoRedoState();

    engine.transact(() => {
      const refId = engine.createRef(finiteRange("B1"));
      engine.deleteRef(refId);
      engine.addCellStyle({
        areas: [finiteRange("C1")],
        style: { bold: true },
      });
      engine.removeCellStyle(workbookName, 0);
    });

    expect(engine.getUndoRedoState()).toMatchObject({
      undoDepth: before.undoDepth,
      undoBytes: before.undoBytes,
    });
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(address("A1"))).toBe("");
  });

  test("coalesces from an empty value using the original before state", () => {
    const engine = buildEngine();

    engine.transact(() => {
      engine.setCellContent(address("A1"), "first");
      engine.setCellContent(address("A1"), "second");
    });

    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(address("A1"))).toBe("");
    expect(engine.redo()).toBe(true);
    expect(engine.getCellValue(address("A1"))).toBe("second");
  });

  test("undo and redo preserve unrelated resolved evaluation nodes", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["B1", "=A1+1"],
        ["C1", "=10+1"],
      ])
    );

    expect(engine.getCellValue(address("B1"))).toBe(2);
    expect(engine.getCellValue(address("C1"))).toBe(11);

    const unrelatedNodeKey = `cell-value:${workbookName}:${sheetName}:C1`;
    const unrelatedNode =
      engine._dependencyManager.getCellValueNode(unrelatedNodeKey);
    expect(unrelatedNode.resolved).toBe(true);
    engine.clearUndoRedoHistory();

    engine.setCellContent(address("A1"), 5);
    expect(engine.getCellValue(address("B1"))).toBe(6);
    expect(unrelatedNode.resolved).toBe(true);

    expect(engine.undo()).toBe(true);
    expect(engine._dependencyManager.getCellValueNode(unrelatedNodeKey)).toBe(
      unrelatedNode
    );
    expect(unrelatedNode.resolved).toBe(true);
    expect(engine.getCellValue(address("B1"))).toBe(2);

    expect(engine.redo()).toBe(true);
    expect(engine._dependencyManager.getCellValueNode(unrelatedNodeKey)).toBe(
      unrelatedNode
    );
    expect(unrelatedNode.resolved).toBe(true);
    expect(engine.getCellValue(address("B1"))).toBe(6);
  });

  test("unrelated manager history leaves the formula cache intact", () => {
    const engine = buildEngine();
    engine.setCellContent(address("A1"), "=40+2");
    expect(engine.getCellValue(address("A1"))).toBe(42);

    const nodeKey = `cell-value:${workbookName}:${sheetName}:A1`;
    const node = engine._dependencyManager.getCellValueNode(nodeKey);
    expect(node.resolved).toBe(true);
    engine.clearUndoRedoHistory();

    engine.createRef(finiteRange("B1"));
    expect(engine.undo()).toBe(true);
    expect(engine._dependencyManager.getCellValueNode(nodeKey)).toBe(node);
    expect(node.resolved).toBe(true);

    engine.addNamedExpression({
      expressionName: "UNUSED_HISTORY_NAME",
      expression: "1",
    });
    engine.clearUndoRedoHistory();
    engine.updateNamedExpression({
      expressionName: "UNUSED_HISTORY_NAME",
      expression: "2",
    });
    expect(engine.undo()).toBe(true);
    expect(engine._dependencyManager.getCellValueNode(nodeKey)).toBe(node);
    expect(node.resolved).toBe(true);
  });

  test("bulk invalidation deduplication uses bounded transient state", () => {
    const engine = FormulaEngine.buildEmpty({
      undoRedo: { maxBytes: 1_024 },
    });
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });

    const content = new Map<string, SerializedCellValue>();
    for (let row = 1; row <= 10_000; row++) {
      content.set(`A${row}`, row);
    }
    engine.setSheetContent({ workbookName, sheetName }, content);
    engine.clearUndoRedoHistory();

    const internals = engine as unknown as {
      observerInvalidatedCellKeys: Set<string>;
    };
    const evaluationManager = engine._evaluationManager;
    const invalidate =
      evaluationManager.invalidateFromMutation.bind(evaluationManager);
    let maxDeduplicationKeys = 0;
    evaluationManager.invalidateFromMutation = (footprint) => {
      maxDeduplicationKeys = Math.max(
        maxDeduplicationKeys,
        internals.observerInvalidatedCellKeys.size
      );
      invalidate(footprint);
    };

    engine.clearSpreadsheetRange(finiteRange("A1", "A10000"));
    expect(maxDeduplicationKeys).toBeGreaterThan(0);
    expect(maxDeduplicationKeys).toBeLessThanOrEqual(4_096);
  });

  test("rolls back an escaping transaction exception without notifying", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["B1", "=A1*2"],
      ])
    );
    expect(engine.getCellValue(address("B1"))).toBe(2);
    engine.clearUndoRedoHistory();

    let updates = 0;
    const unsubscribe = engine.onUpdate(() => {
      updates++;
    });

    expect(() =>
      engine.transact(() => {
        engine.setCellContent(address("A1"), 5);
        engine.setCellContent(address("C1"), 10);
        expect(engine.getCellValue(address("B1"))).toBe(10);
        throw new Error("abort transaction");
      })
    ).toThrow("abort transaction");
    unsubscribe();

    expect(updates).toBe(0);
    expect(engine.getCellValue(address("A1"))).toBe(1);
    expect(engine.getCellValue(address("B1"))).toBe(2);
    expect(engine.getCellValue(address("C1"))).toBe("");
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
  });

  test("rejects and rolls back asynchronous transaction callbacks", async () => {
    const engine = buildEngine();
    engine.clearUndoRedoHistory();

    const unsupportedTransact = engine.transact.bind(engine) as unknown as (
      callback: () => Promise<void>
    ) => Promise<void>;
    await expect(
      unsupportedTransact(async () => {
        engine.setCellContent(address("A1"), 1);
        await Promise.resolve();
        engine.setCellContent(address("B1"), 2);
      })
    ).rejects.toThrow("transact callback must be synchronous");

    expect(engine.getCellValue(address("A1"))).toBe("");
    expect(engine.getCellValue(address("B1"))).toBe("");
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
    });
  });

  test("an oversized entry creates a barrier without reverting its mutation", () => {
    const engine = buildEngine({
      undoRedo: { maxEntries: 100, maxBytes: 4 * 1024 },
    });

    engine.setCellContent(address("A1"), 1);
    engine.setCellContent(address("A2"), 2);
    expect(engine.getUndoRedoState().undoDepth).toBe(2);
    expect(engine.undo()).toBe(true);
    expect(engine.getUndoRedoState()).toMatchObject({
      undoDepth: 1,
      redoDepth: 1,
    });

    const oversizedValue = "x".repeat(100_000);
    engine.setCellContent(address("A3"), oversizedValue);

    expect(engine.getCellValue(address("A3"))).toBe(oversizedValue);
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
    expect(engine.undo()).toBe(false);

    engine.setCellContent(address("A4"), 4);
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(address("A4"))).toBe("");
    expect(engine.getCellValue(address("A3"))).toBe(oversizedValue);
    expect(engine.undo()).toBe(false);
  });

  test("an explicit transaction rejects and rolls back an oversized mutation", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 1_000 } });
    engine.setCellContent(address("A1"), "prior");
    const historyBefore = engine.getUndoRedoState();

    expect(() =>
      engine.transact(() => {
        engine.setCellContent(address("A1"), "x".repeat(100_000));
      })
    ).toThrow("exceeded undoRedo.maxBytes");

    expect(engine.getCellValue(address("A1"))).toBe("prior");
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("an explicit transaction budgets the final history envelope atomically", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 300 } });
    engine.setCellContent(address("A1"), "before");
    engine.clearUndoRedoHistory();
    const historyBefore = engine.getUndoRedoState();

    expect(() =>
      engine.transact(() => {
        engine.setCellContent(address("A1"), "after");
      })
    ).toThrow("exceeded undoRedo.maxBytes");

    expect(
      engine.getSheetSerialized({ workbookName, sheetName }).get("A1")
    ).toBe("before");
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("an oversized explicit range clear restores every buffered cell", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 4 * 1024 } });
    const content = new Map<string, SerializedCellValue>();
    for (let rowIndex = 1; rowIndex <= 100; rowIndex++) {
      content.set(`A${rowIndex}`, `value-${rowIndex}-${"x".repeat(100)}`);
    }
    engine.setSheetContent({ workbookName, sheetName }, content);
    engine.clearUndoRedoHistory();
    const before = Array.from(
      engine.getSheetSerialized({ workbookName, sheetName })
    );
    const historyBefore = engine.getUndoRedoState();

    expect(() =>
      engine.transact(() => {
        engine.clearSpreadsheetRange(finiteRange("A1", "A100"));
      })
    ).toThrow("exceeded undoRedo.maxBytes");

    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual(before);
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("an oversized explicit batched paste leaves no unjournaled writes", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 4 * 1024 } });
    const content = new Map<string, SerializedCellValue>();
    const source: CellAddress[] = [];
    for (let rowIndex = 1; rowIndex <= 100; rowIndex++) {
      content.set(`A${rowIndex}`, `value-${rowIndex}-${"x".repeat(100)}`);
      source.push(address(`A${rowIndex}`));
    }
    engine.setSheetContent({ workbookName, sheetName }, content);
    engine.clearUndoRedoHistory();
    const before = Array.from(
      engine.getSheetSerialized({ workbookName, sheetName })
    );
    const historyBefore = engine.getUndoRedoState();

    expect(() =>
      engine.transact(() => {
        engine.pasteCells(source, address("B1"), {
          cut: false,
          type: "formula",
          include: ["content"],
        });
      })
    ).toThrow("exceeded undoRedo.maxBytes");

    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual(before);
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("an oversized explicit cut restores sequential source deletions", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 4 * 1024 } });
    const content = new Map<string, SerializedCellValue>();
    const source: CellAddress[] = [];
    for (let rowIndex = 1; rowIndex <= 100; rowIndex++) {
      content.set(`A${rowIndex}`, `value-${rowIndex}-${"x".repeat(100)}`);
      source.push(address(`A${rowIndex}`));
    }
    engine.setSheetContent({ workbookName, sheetName }, content);
    engine.clearUndoRedoHistory();
    const before = Array.from(
      engine.getSheetSerialized({ workbookName, sheetName })
    );
    const historyBefore = engine.getUndoRedoState();

    expect(() =>
      engine.transact(() => {
        engine.pasteCells(source, address("B1"), {
          cut: true,
          type: "formula",
          include: ["content"],
        });
      })
    ).toThrow("exceeded undoRedo.maxBytes");

    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual(before);
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("explicit transaction budgeting accounts for coalesced replacements", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 500 } });
    engine.setCellContent(address("A1"), "before");
    engine.clearUndoRedoHistory();

    engine.transact(() => {
      engine.setCellContent(address("A1"), "one");
      engine.setCellContent(address("A1"), "two");
    });

    expect(engine.getCellValue(address("A1"))).toBe("two");
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: true,
      undoDepth: 1,
    });
    expect(engine.undo()).toBe(true);
    expect(engine.getCellValue(address("A1"))).toBe("before");
  });

  test("an explicit transaction rejects unsafe metadata atomically", () => {
    const engine = buildEngine();
    engine.setCellContent(address("A1"), "prior-history");
    const historyBefore = engine.getUndoRedoState();

    expect(() =>
      engine.transact(() => {
        engine.setCellMetadata(address("B1"), {
          callback: () => "not retainable",
        });
      })
    ).toThrow("cannot be retained safely");

    expect(engine.getCellMetadata(address("B1"))).toBeUndefined();
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("an oversized sheet replacement aborts before an explicit transaction mutates", () => {
    const engine = buildEngine({ undoRedo: { maxBytes: 4 * 1024 } });
    engine.setCellContent(address("A1"), "before");
    const historyBefore = engine.getUndoRedoState();
    const content = new Map<string, SerializedCellValue>();
    for (let rowIndex = 1; rowIndex <= 3_000; rowIndex++) {
      content.set(`A${rowIndex}`, `large-${rowIndex}-${"x".repeat(32)}`);
    }

    expect(() =>
      engine.transact(() => {
        engine.setSheetContent({ workbookName, sheetName }, content);
      })
    ).toThrow("exceeded undoRedo.maxBytes");

    expect(
      Array.from(engine.getSheetSerialized({ workbookName, sheetName }))
    ).toEqual([["A1", "before"]]);
    expect(engine.getUndoRedoState()).toEqual(historyBefore);
  });

  test("an oversized bulk sheet replacement is one non-undoable barrier", () => {
    const engine = buildEngine({
      undoRedo: { maxEntries: 100, maxBytes: 8 * 1024 },
    });
    engine.setCellContent(address("A1"), "history-before-import");

    const content = new Map<string, SerializedCellValue>();
    for (let rowIndex = 1; rowIndex <= 3_000; rowIndex++) {
      content.set(`A${rowIndex}`, `imported-${rowIndex}-${"x".repeat(32)}`);
    }

    engine.setSheetContent({ workbookName, sheetName }, content);

    expect(engine.getCellValue(address("A1"))).toBe(
      `imported-1-${"x".repeat(32)}`
    );
    expect(engine.getCellValue(address("A3000"))).toBe(
      `imported-3000-${"x".repeat(32)}`
    );
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
    expect(engine.undo()).toBe(false);
  });

  test("workbook rename replay preserves workbook order", () => {
    const engine = buildEngine();
    engine.addWorkbook("Second");
    engine.addWorkbook("Third");
    engine.clearUndoRedoHistory();

    engine.renameWorkbook({
      workbookName: "Second",
      newWorkbookName: "Renamed",
    });
    expect(Array.from(engine.getWorkbooks().keys())).toEqual([
      workbookName,
      "Renamed",
      "Third",
    ]);

    expect(engine.undo()).toBe(true);
    expect(Array.from(engine.getWorkbooks().keys())).toEqual([
      workbookName,
      "Second",
      "Third",
    ]);
    expect(engine.redo()).toBe(true);
    expect(Array.from(engine.getWorkbooks().keys())).toEqual([
      workbookName,
      "Renamed",
      "Third",
    ]);
  });

  test("binary metadata changes are not mistaken for no-ops", () => {
    const engine = FormulaEngine.buildEmpty<{ cell: ArrayBuffer }>();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    engine.setCellMetadata(address("A1"), new Uint8Array([1]).buffer);
    engine.clearUndoRedoHistory();

    engine.setCellMetadata(address("A1"), new Uint8Array([2]).buffer);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);
    expect(engine.undo()).toBe(true);
    expect(
      Array.from(new Uint8Array(engine.getCellMetadata(address("A1"))!))
    ).toEqual([1]);
    expect(engine.redo()).toBe(true);
    expect(
      Array.from(new Uint8Array(engine.getCellMetadata(address("A1"))!))
    ).toEqual([2]);
  });

  test("metadata Map order and RegExp state round-trip exactly", () => {
    const engine = FormulaEngine.buildEmpty<{
      cell: Map<string, number> | RegExp;
    }>();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });

    engine.setCellMetadata(
      address("A1"),
      new Map([
        ["first", 1],
        ["second", 2],
      ])
    );
    engine.clearUndoRedoHistory();
    engine.setCellMetadata(
      address("A1"),
      new Map([
        ["second", 2],
        ["first", 1],
      ])
    );
    expect(engine.undo()).toBe(true);
    expect(
      Array.from(
        (engine.getCellMetadata(address("A1")) as Map<string, number>).keys()
      )
    ).toEqual(["first", "second"]);
    expect(engine.redo()).toBe(true);
    expect(
      Array.from(
        (engine.getCellMetadata(address("A1")) as Map<string, number>).keys()
      )
    ).toEqual(["second", "first"]);

    const before = /before/gi;
    before.lastIndex = 3;
    engine.setCellMetadata(address("B1"), before);
    engine.clearUndoRedoHistory();
    const after = /after/m;
    after.lastIndex = 1;
    engine.setCellMetadata(address("B1"), after);
    expect(engine.undo()).toBe(true);
    const restoredBefore = engine.getCellMetadata(address("B1")) as RegExp;
    expect([
      restoredBefore.source,
      restoredBefore.flags,
      restoredBefore.lastIndex,
    ]).toEqual(["before", "gi", 3]);
    expect(engine.redo()).toBe(true);
    const restoredAfter = engine.getCellMetadata(address("B1")) as RegExp;
    expect([
      restoredAfter.source,
      restoredAfter.flags,
      restoredAfter.lastIndex,
    ]).toEqual(["after", "m", 1]);
  });

  test("unsupported metadata creates a barrier without invoking accessors", () => {
    const engine = FormulaEngine.buildEmpty<{ cell: object }>();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    engine.setCellContent(address("A1"), "older-history");

    let getterReads = 0;
    const metadata = {
      callback: () => "opaque closure",
    } as { callback: () => string; dangerous?: string };
    Object.defineProperty(metadata, "dangerous", {
      enumerable: true,
      get() {
        getterReads++;
        throw new Error("getter must not run");
      },
    });

    expect(() => engine.setCellMetadata(address("B1"), metadata)).not.toThrow();
    expect(getterReads).toBe(0);
    expect(engine.getCellMetadata(address("B1"))).toBe(metadata);
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
  });

  test("rejects custom built-in prototypes and typed-view accessors atomically", () => {
    const engine = FormulaEngine.buildEmpty<{ cell: unknown }>({
      undoRedo: { maxBytes: 4 * 1024 },
    });
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });

    class CustomError extends Error {}
    Object.defineProperty(CustomError.prototype, "payload", {
      value: new Uint8Array(1024),
    });
    expect(() =>
      engine.transact(() => {
        engine.setCellMetadata(address("A1"), new CustomError("unsafe"));
      })
    ).toThrow("cannot be retained safely");
    expect(engine.getCellMetadata(address("A1"))).toBeUndefined();

    class CustomBytes extends Uint8Array {}
    Object.defineProperty(CustomBytes.prototype, "payload", {
      value: new Uint8Array(1024),
    });
    expect(() =>
      engine.transact(() => {
        engine.setCellMetadata(address("A1"), new CustomBytes(1));
      })
    ).toThrow("cannot be retained safely");
    expect(engine.getCellMetadata(address("A1"))).toBeUndefined();

    let getterReads = 0;
    const bytes = new Uint8Array(1);
    Object.defineProperty(bytes, "byteLength", {
      get() {
        getterReads++;
        throw new Error("must not run");
      },
    });
    expect(() =>
      engine.transact(() => {
        engine.setCellMetadata(address("A1"), bytes);
      })
    ).toThrow("cannot be retained safely");
    expect(getterReads).toBe(0);
    expect(engine.getCellMetadata(address("A1"))).toBeUndefined();

    for (const proxy of [
      new Proxy(new ArrayBuffer(1), {}),
      new Proxy(new Date(), {}),
      new Proxy(/unsafe/, {}),
    ]) {
      expect(() =>
        engine.transact(() => {
          engine.setCellMetadata(address("A1"), proxy);
        })
      ).toThrow("cannot be retained safely");
      expect(engine.getCellMetadata(address("A1"))).toBeUndefined();
    }
  });

  test("large typed metadata is rejected in O(1) without a history clone", () => {
    const engine = FormulaEngine.buildEmpty<{ cell: Uint8Array }>({
      undoRedo: { maxBytes: 1_024 },
    });
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    const binary = new Uint8Array(4 * 1024 * 1024);

    const startedAt = performance.now();
    engine.setCellMetadata(address("A1"), binary);
    const elapsedMs = performance.now() - startedAt;

    expect(engine.getCellMetadata(address("A1"))).toBe(binary);
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      undoDepth: 0,
      undoBytes: 0,
    });
    expect(elapsedMs).toBeLessThan(500);
  });

  test("large range metadata is budgeted before manager detachment", () => {
    const engine = FormulaEngine.buildEmpty<{ range: Uint8Array }>({
      undoRedo: { maxBytes: 1_024 },
    });
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    const binary = new Uint8Array(4 * 1024 * 1024);

    const startedAt = performance.now();
    engine.addRangeMetadata({
      areas: [finiteRange("A1")],
      metadata: binary,
    });
    const elapsedMs = performance.now() - startedAt;

    expect(engine.getAllRangeMetadata()[0]?.metadata).toBe(binary);
    expect(engine.getUndoRedoState()).toMatchObject({
      canUndo: false,
      undoDepth: 0,
      undoBytes: 0,
    });
    expect(elapsedMs).toBeLessThan(500);
  });

  test("a throwing update listener does not duplicate history entries", () => {
    const engine = buildEngine();
    engine.setCellContent(address("A1"), 1);
    const unsubscribe = engine.onUpdate(() => {
      throw new Error("listener failed");
    });

    expect(() => engine.undo()).toThrow("listener failed");
    unsubscribe();
    expect(engine.getCellValue(address("A1"))).toBe("");
    expect(engine.getUndoRedoState()).toMatchObject({
      undoDepth: 0,
      redoDepth: 1,
    });
    expect(engine.redo()).toBe(true);
    expect(engine.getCellValue(address("A1"))).toBe(1);
  });
});
