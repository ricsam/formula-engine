import { describe, expect, test } from "bun:test";
import type { EvaluationManager } from "../../../../src/core/managers/evaluation-manager";
import {
  RangeMetadataManager,
  type RangeMetadataDataChange,
} from "../../../../src/core/managers/range-metadata-manager";
import {
  ReferenceManager,
  type ReferenceDataChange,
} from "../../../../src/core/managers/reference-manager";
import {
  StyleManager,
  type StyleDataChange,
} from "../../../../src/core/managers/style-manager";
import type {
  DirectCellStyle,
  RangeAddress,
  TrackedReference,
} from "../../../../src/core/types";
import { parseCellReference } from "../../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";

function range(
  start: string,
  end = start,
  targetSheet = sheetName,
  targetWorkbook = workbookName
): RangeAddress {
  const startAddress = parseCellReference(start);
  const endAddress = parseCellReference(end);
  return {
    workbookName: targetWorkbook,
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

function styleManager(observer: (changes: readonly StyleDataChange[]) => void) {
  return new StyleManager({} as EvaluationManager, observer);
}

describe("StyleManager mutation observer", () => {
  test("does not build or emit deltas while observation is inactive", () => {
    let active = false;
    const batches: StyleDataChange[][] = [];
    const manager = new StyleManager(
      {} as EvaluationManager,
      (changes) => batches.push([...changes]),
      () => active
    );

    manager.addCellStyle({ areas: [range("A1")], style: { bold: true } });
    expect(batches).toHaveLength(0);

    active = true;
    manager.addCellStyle({ areas: [range("A2")], style: { italic: true } });
    expect(batches.flat()).toHaveLength(1);
  });

  test("emits detached add/remove deltas and skips failed removals", () => {
    const batches: StyleDataChange[][] = [];
    const manager = styleManager((changes) => batches.push([...changes]));
    const style: DirectCellStyle = {
      areas: [range("A1")],
      style: { bold: true },
    };

    manager.addCellStyle(style);
    style.style.bold = false;
    style.areas[0]!.sheetName = "MutatedElsewhere";

    expect(batches).toHaveLength(1);
    expect(batches[0]?.[0]).toEqual({
      kind: "cell-style",
      after: {
        index: 0,
        value: {
          areas: [range("A1")],
          style: { bold: true },
        },
      },
    });

    expect(manager.removeCellStyle(workbookName, 99)).toBe(false);
    expect(batches).toHaveLength(1);
  });

  test("reports only changed rules and replays ordered deltas silently", () => {
    const batches: StyleDataChange[][] = [];
    const manager = styleManager((changes) => batches.push([...changes]));
    manager.addCellStyle({
      areas: [range("A1", "C3")],
      style: { backgroundColor: "#ff0000" },
    });
    manager.addCellStyle({
      areas: [range("A1", "C3", "OtherSheet")],
      style: { backgroundColor: "#0000ff" },
    });
    const before = structuredClone(manager.getAllCellStyles());
    batches.length = 0;

    manager.clearCellStylesInRange(range("B2"));

    const after = structuredClone(manager.getAllCellStyles());
    const changes = batches.flat();
    expect(changes).toHaveLength(2);
    expect(changes.every((change) => change.kind === "cell-style")).toBe(true);
    expect(
      changes.some(
        (change) =>
          change.before?.value.areas[0]?.sheetName === "OtherSheet" ||
          change.after?.value.areas[0]?.sheetName === "OtherSheet"
      )
    ).toBe(false);

    const notificationCount = batches.length;
    manager.applyHistoryChanges(changes, "undo");
    expect(manager.getAllCellStyles()).toEqual(before);
    expect(batches).toHaveLength(notificationCount);

    manager.applyHistoryChanges(changes, "redo");
    expect(manager.getAllCellStyles()).toEqual(after);
    expect(batches).toHaveLength(notificationCount);
  });

  test("preserves atomic patches across a large batched operation", () => {
    const batchSizes: number[] = [];
    const manager = styleManager((changes) => batchSizes.push(changes.length));

    manager.batchMutations(() => {
      for (let index = 0; index < 1_030; index++) {
        manager.addCellStyle({
          areas: [range(`A${index + 1}`)],
          style: { fontSize: 10 + (index % 5) },
        });
      }
    });

    expect(batchSizes).toEqual([1_030]);
  });
});

describe("RangeMetadataManager mutation observer", () => {
  test("emits detached sparse changes and replays exact order silently", () => {
    type Metadata = { label: string; nested: { value: number } };
    const batches: RangeMetadataDataChange<Metadata>[][] = [];
    const manager = new RangeMetadataManager<Metadata>((changes) =>
      batches.push([...changes])
    );
    const metadata = { label: "affected", nested: { value: 1 } };

    manager.addRangeMetadata({
      id: "affected",
      areas: [range("A1", "C3")],
      metadata,
    });
    manager.addRangeMetadata({
      id: "unrelated",
      areas: [range("A1", "C3", "OtherSheet")],
      metadata: { label: "unrelated", nested: { value: 2 } },
    });
    expect(batches[0]?.[0]?.after?.value.metadata).toEqual({
      label: "affected",
      nested: { value: 1 },
    });
    metadata.nested.value = 99;
    expect(batches[0]?.[0]?.after?.value.metadata.nested.value).toBe(1);

    const before = structuredClone(manager.getAllRangeMetadata());
    batches.length = 0;
    manager.clearRangeMetadataInRange(range("B2"));

    const after = structuredClone(manager.getAllRangeMetadata());
    const changes = batches.flat();
    expect(changes).toHaveLength(1);
    expect(changes[0]?.id).toBe("affected");
    expect(changes[0]?.before?.index).toBe(0);
    expect(changes[0]?.after?.index).toBe(0);

    const notificationCount = batches.length;
    manager.applyHistoryChanges(changes, "undo");
    expect(manager.getAllRangeMetadata()).toEqual(before);
    expect(batches).toHaveLength(notificationCount);

    manager.applyHistoryChanges(changes, "redo");
    expect(manager.getAllRangeMetadata()).toEqual(after);
    expect(batches).toHaveLength(notificationCount);
  });

  test("preserves insertion positions for add and remove replay", () => {
    const batches: RangeMetadataDataChange<{ label: string }>[][] = [];
    const manager = new RangeMetadataManager<{ label: string }>((changes) =>
      batches.push([...changes])
    );
    for (const id of ["first", "middle", "last"]) {
      manager.addRangeMetadata({
        id,
        areas: [range("A1")],
        metadata: { label: id },
      });
    }
    batches.length = 0;

    manager.removeRangeMetadata("middle");
    const changes = batches.flat();
    expect(changes[0]?.before?.index).toBe(1);
    manager.applyHistoryChanges(changes, "undo");
    expect(manager.getAllRangeMetadata().map(({ id }) => id)).toEqual([
      "first",
      "middle",
      "last",
    ]);
    manager.applyHistoryChanges(changes, "redo");
    expect(manager.getAllRangeMetadata().map(({ id }) => id)).toEqual([
      "first",
      "last",
    ]);
  });

  test("best-effort detaches metadata containing unsupported functions", () => {
    type Metadata = {
      callback: () => string;
      nested: { value: number };
    };
    const batches: RangeMetadataDataChange<Metadata>[][] = [];
    const manager = new RangeMetadataManager<Metadata>((changes) =>
      batches.push([...changes])
    );
    const metadata: Metadata = {
      callback: () => "ok",
      nested: { value: 1 },
    };

    expect(() =>
      manager.addRangeMetadata({
        id: "function-metadata",
        areas: [range("A1")],
        metadata,
      })
    ).not.toThrow();
    metadata.nested.value = 2;

    expect(batches[0]?.[0]?.after?.value.metadata.nested.value).toBe(1);
    expect(batches[0]?.[0]?.after?.value.metadata.callback()).toBe("ok");
  });
});

describe("ReferenceManager mutation observer", () => {
  test("emits detached rename/invalidation deltas for affected refs only", () => {
    const batches: ReferenceDataChange[][] = [];
    const manager = new ReferenceManager((changes) =>
      batches.push([...changes])
    );
    const firstId = manager.createRef(range("A1"));
    manager.createRef(range("A1", "A1", "OtherSheet"));
    batches.length = 0;

    manager.updateSheetName(workbookName, sheetName, "Renamed");

    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "reference",
      id: firstId,
      before: { index: 0 },
      after: { index: 0 },
    });
    expect(batches[0]?.[0]?.before?.value.address.sheetName).toBe(sheetName);
    expect(batches[0]?.[0]?.after?.value.address.sheetName).toBe("Renamed");

    const retainedAfter = batches[0]?.[0]?.after?.value;
    manager.updateSheetName(workbookName, "Renamed", "RenamedAgain");
    expect(retainedAfter?.address.sheetName).toBe("Renamed");

    batches.length = 0;
    manager.invalidateWorkbook(workbookName);
    expect(batches.flat()).toHaveLength(2);
    expect(
      batches.flat().every((change) => change.after?.value.isValid === false)
    ).toBe(true);
  });

  test("replays deletion with exact Map order without observer recursion", () => {
    const batches: ReferenceDataChange[][] = [];
    const manager = new ReferenceManager((changes) =>
      batches.push([...changes])
    );
    const ids = [
      manager.createRef(range("A1")),
      manager.createRef(range("B1")),
      manager.createRef(range("C1")),
    ];
    const before = structuredClone(manager.getAllReferences()) as Map<
      string,
      TrackedReference
    >;
    batches.length = 0;

    manager.deleteRef(ids[1]!);
    const changes = batches.flat();
    const after = structuredClone(manager.getAllReferences()) as Map<
      string,
      TrackedReference
    >;
    expect(changes[0]?.before?.index).toBe(1);

    const notificationCount = batches.length;
    manager.applyHistoryChanges(changes, "undo");
    expect(manager.getAllReferences()).toEqual(before);
    expect(batches).toHaveLength(notificationCount);

    manager.applyHistoryChanges(changes, "redo");
    expect(manager.getAllReferences()).toEqual(after);
    expect(batches).toHaveLength(notificationCount);
  });
});
