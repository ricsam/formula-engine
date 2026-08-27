import { describe, expect, test } from "bun:test";
import {
  createHistoryEntry,
  estimateHistoryStepsBytes,
  estimateHistoryValueBytes,
  isHistoryValueSafelyRetainable,
  type HistoryStep,
} from "../../../src/core/history";
import { UndoRedoManager } from "../../../src/core/managers/undo-redo-manager";

interface TestStep extends HistoryStep {
  readonly kind: "test";
  readonly id: number;
}

function entry(id: number, estimatedBytes: number) {
  return createHistoryEntry<TestStep>([{ kind: "test", id, estimatedBytes }], {
    estimatedBytes,
  });
}

describe("UndoRedoManager", () => {
  test("retains incremental entries and reports byte accounting", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 3,
      maxBytes: 100,
    });

    expect(manager.record(entry(1, 20))).toBe("recorded");
    expect(manager.record(entry(2, 30))).toBe("recorded");

    expect(manager.getState()).toEqual({
      enabled: true,
      canUndo: true,
      canRedo: false,
      undoDepth: 2,
      redoDepth: 0,
      maxEntries: 3,
      maxBytes: 100,
      undoBytes: 50,
      redoBytes: 0,
    });
    expect(manager.popUndo()?.steps[0]?.id).toBe(2);
    expect(manager.getState().undoBytes).toBe(20);
  });

  test("evicts the oldest undo entries to satisfy entry and byte limits", () => {
    const entryLimited = new UndoRedoManager<TestStep>({
      maxEntries: 2,
      maxBytes: 1_000,
    });
    entryLimited.record(entry(1, 10));
    entryLimited.record(entry(2, 10));
    entryLimited.record(entry(3, 10));

    expect(entryLimited.getState()).toMatchObject({
      undoDepth: 2,
      undoBytes: 20,
    });
    expect(entryLimited.popUndo()?.steps[0]?.id).toBe(3);
    expect(entryLimited.popUndo()?.steps[0]?.id).toBe(2);

    const byteLimited = new UndoRedoManager<TestStep>({
      maxEntries: 10,
      maxBytes: 25,
    });
    byteLimited.record(entry(1, 10));
    byteLimited.record(entry(2, 10));
    byteLimited.record(entry(3, 10));

    expect(byteLimited.getState()).toMatchObject({
      undoDepth: 2,
      undoBytes: 20,
    });
    expect(byteLimited.popUndo()?.steps[0]?.id).toBe(3);
    expect(byteLimited.popUndo()?.steps[0]?.id).toBe(2);
  });

  test("moves entries between stacks during replay without clearing either", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 4,
      maxBytes: 100,
    });
    manager.record(entry(1, 10));
    manager.record(entry(2, 20));

    const undone = manager.popUndo();
    expect(undone).toBeDefined();
    manager.pushRedoFromReplay(undone!);

    expect(manager.getState()).toMatchObject({
      undoDepth: 1,
      redoDepth: 1,
      undoBytes: 10,
      redoBytes: 20,
    });

    const redone = manager.popRedo();
    expect(redone).toBeDefined();
    manager.pushUndoFromReplay(redone!);

    expect(manager.getState()).toMatchObject({
      undoDepth: 2,
      redoDepth: 0,
      undoBytes: 30,
      redoBytes: 0,
    });
  });

  test("a new mutation clears redo history", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 4,
      maxBytes: 100,
    });
    manager.record(entry(1, 10));
    const undone = manager.popUndo();
    manager.pushRedoFromReplay(undone!);

    manager.record(entry(2, 15));

    expect(manager.getState()).toMatchObject({
      undoDepth: 1,
      redoDepth: 0,
      undoBytes: 15,
      redoBytes: 0,
    });
  });

  test("an oversized mutation creates a history barrier", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 4,
      maxBytes: 25,
    });
    manager.record(entry(1, 10));
    manager.record(entry(2, 10));
    const undone = manager.popUndo();
    manager.pushRedoFromReplay(undone!);

    expect(manager.record(entry(3, 26))).toBe("oversized");
    expect(manager.getState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
  });

  test("an explicit barrier releases both replay stacks", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 10,
      maxBytes: 10_000,
    });
    manager.record(entry(1, 100));
    manager.record(entry(2, 200));
    manager.pushRedoFromReplay(manager.popUndo()!);

    manager.recordBarrier();

    expect(manager.getState()).toMatchObject({
      canUndo: false,
      canRedo: false,
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
  });

  test("retains zero-byte entries and bounds them by entry count", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 1,
      maxBytes: 25,
    });

    manager.record(entry(1, 0));
    manager.record(entry(2, 0));

    expect(manager.getState()).toMatchObject({
      undoDepth: 1,
      undoBytes: 0,
    });
    expect(manager.popUndo()?.steps[0]?.id).toBe(2);
  });

  test("clear releases both stacks and resets accounting", () => {
    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 4,
      maxBytes: 100,
    });
    manager.record(entry(1, 10));
    const undone = manager.popUndo();
    manager.pushRedoFromReplay(undone!);

    manager.clear();

    expect(manager.getState()).toMatchObject({
      undoDepth: 0,
      redoDepth: 0,
      undoBytes: 0,
      redoBytes: 0,
    });
  });

  test("validates limits and entry byte sizes", () => {
    expect(() => new UndoRedoManager({ maxEntries: 0, maxBytes: 10 })).toThrow(
      "undoRedo.maxEntries must be a positive safe integer"
    );
    expect(
      () => new UndoRedoManager({ maxEntries: 10, maxBytes: 1.5 })
    ).toThrow("undoRedo.maxBytes must be a positive safe integer");

    const manager = new UndoRedoManager<TestStep>({
      maxEntries: 10,
      maxBytes: 100,
    });
    expect(() =>
      manager.record({
        steps: [{ kind: "test", id: 1 }],
        estimatedBytes: Number.NaN,
      })
    ).toThrow("history entry estimatedBytes");
  });
});

describe("history size estimation", () => {
  test("uses explicit step estimates when available", () => {
    expect(
      estimateHistoryStepsBytes<TestStep>([
        { kind: "test", id: 1, estimatedBytes: 11 },
        { kind: "test", id: 2, estimatedBytes: 13 },
      ])
    ).toBe(24);
  });

  test("estimates maps, typed arrays, shared references, and cycles", () => {
    const shared = { value: "abc" };
    const cyclic: { self?: unknown } = {};
    cyclic.self = cyclic;
    const payload = new Map<unknown, unknown>([
      ["typed", new Uint8Array(16)],
      [shared, new Set([shared, cyclic])],
    ]);

    const bytes = estimateHistoryValueBytes(payload);

    expect(Number.isSafeInteger(bytes)).toBe(true);
    expect(bytes).toBeGreaterThan(16);
  });

  test("copies the steps array into the entry envelope", () => {
    const steps: TestStep[] = [{ kind: "test", id: 1 }];
    const historyEntry = createHistoryEntry(steps);
    steps.push({ kind: "test", id: 2 });

    expect(historyEntry.steps).toHaveLength(1);
    expect(historyEntry.estimatedBytes).toBeGreaterThan(0);
  });

  test("rejects custom built-in prototypes and shadowed view intrinsics", () => {
    class CustomError extends Error {}
    class CustomBuffer extends ArrayBuffer {}
    class CustomBytes extends Uint8Array {}
    class CustomDate extends Date {}
    class CustomMap extends Map<string, string> {}

    expect(isHistoryValueSafelyRetainable(new CustomError())).toBe(false);
    expect(isHistoryValueSafelyRetainable(new CustomBuffer(1))).toBe(false);
    expect(isHistoryValueSafelyRetainable(new CustomBytes(1))).toBe(false);
    expect(isHistoryValueSafelyRetainable(new CustomDate())).toBe(false);
    expect(isHistoryValueSafelyRetainable(new CustomMap())).toBe(false);

    const bytes = new Uint8Array(1);
    Object.defineProperty(bytes, "constructor", {
      value: class EvilBytes extends Uint8Array {},
    });
    expect(isHistoryValueSafelyRetainable(bytes)).toBe(false);

    expect(
      isHistoryValueSafelyRetainable(new Proxy(new ArrayBuffer(1), {}))
    ).toBe(false);
    expect(isHistoryValueSafelyRetainable(new Proxy(new Date(), {}))).toBe(
      false
    );
    expect(isHistoryValueSafelyRetainable(new Proxy(/unsafe/, {}))).toBe(false);
  });
});
