import { describe, expect, test } from "bun:test";
import { serialize } from "../../../../src/core/map-serializer";
import {
  TableManager,
  type TableMutation,
} from "../../../../src/core/managers/table-manager";
import { WorkbookManager } from "../../../../src/core/managers/workbook-manager";
import type { TableDefinition } from "../../../../src/core/types";

function table(
  name: string,
  workbookName = "Book",
  sheetName = "Sheet1",
  rowIndex = 1
): TableDefinition {
  return {
    name,
    workbookName,
    sheetName,
    start: { rowIndex, colIndex: 1 },
    endRow: { type: "number", value: rowIndex + 2 },
    headers: new Map([["Column", { name: "Column", index: 0 }]]),
  };
}

class CountingMap<TKey, TValue> extends Map<TKey, TValue> {
  clearCount = 0;

  override clear(): void {
    this.clearCount++;
    super.clear();
  }
}

describe("TableManager mutation observer", () => {
  test("skips delta construction while observation is inactive", () => {
    let active = false;
    const batches: Array<readonly TableMutation[]> = [];
    const manager = new TableManager(
      new WorkbookManager(),
      (changes) => batches.push(changes),
      () => active
    );

    manager.addWorkbook("Ignored");
    expect(batches).toHaveLength(0);

    active = true;
    manager.addWorkbook("Captured");
    expect(batches.flat()).toHaveLength(1);
  });

  test("add, update, rename, and remove emit affected detached entries", () => {
    const batches: Array<readonly TableMutation[]> = [];
    const manager = new TableManager(new WorkbookManager(), (changes) => {
      batches.push(changes);
    });
    manager.addWorkbook("Book");
    batches.length = 0;

    manager.addTable({
      tableName: "Data",
      workbookName: "Book",
      sheetName: "Sheet1",
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 1,
      getCellValue: () => "Column",
    });
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "table",
      before: undefined,
      after: { workbookName: "Book", tableName: "Data", index: 0 },
    });
    const addDelta = batches[0]?.[0];

    batches.length = 0;
    manager.updateTable({
      tableName: "Data",
      workbookName: "Book",
      start: "B2",
      getCellValue: () => "Updated",
    });
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "table",
      before: { table: { start: { rowIndex: 0, colIndex: 0 } } },
      after: { table: { start: { rowIndex: 1, colIndex: 1 } } },
    });
    expect(addDelta).toMatchObject({
      after: { table: { start: { rowIndex: 0, colIndex: 0 } } },
    });

    batches.length = 0;
    manager.renameTable("Book", { oldName: "Data", newName: "Renamed" });
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "table",
      before: { tableName: "Data" },
      after: { tableName: "Renamed", table: { name: "Renamed" } },
    });

    batches.length = 0;
    expect(
      manager.removeTable({ workbookName: "Book", tableName: "Renamed" })
    ).toBe(true);
    expect(batches.flat()).toHaveLength(1);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "table",
      before: { tableName: "Renamed" },
      after: undefined,
    });
  });

  test("header updates remain one atomic indexed patch", () => {
    let capture = false;
    const batchSizes: number[] = [];
    const manager = new TableManager(new WorkbookManager(), (changes) => {
      if (capture) {
        batchSizes.push(changes.length);
      }
    });
    const tables = new Map<string, TableDefinition>();
    const updates = [];
    for (let index = 0; index < 1_030; index++) {
      const definition = table(`Table${index}`, "Book", "Sheet1", index + 1);
      tables.set(definition.name, definition);
      updates.push({
        table: definition,
        index: 0,
        oldName: "Column",
        newName: `Column${index}`,
      });
    }
    manager.tables.set("Book", tables);

    capture = true;
    manager.applyHeaderUpdates(updates);

    expect(batchSizes).toEqual([1_030]);
  });

  test("bulk history replay rebuilds an affected table map once", () => {
    const patches: TableMutation[][] = [];
    let capture = false;
    const manager = new TableManager(new WorkbookManager(), (changes) => {
      if (capture) {
        patches.push([...changes]);
      }
    });
    const tables = new CountingMap<string, TableDefinition>();
    const updates = [];
    for (let index = 0; index < 5_000; index++) {
      const definition = table(`Table${index}`, "Book", "Sheet1", index + 1);
      tables.set(definition.name, definition);
      updates.push({
        table: definition,
        index: 0,
        oldName: "Column",
        newName: `Column${index}`,
      });
    }
    manager.tables.set("Book", tables);
    const expectedOrder = Array.from(tables.keys());

    capture = true;
    manager.applyHeaderUpdates(updates);
    capture = false;
    expect(patches).toHaveLength(1);
    expect(patches[0]).toHaveLength(5_000);

    manager.applyHistoryChanges(patches[0]!, "undo");
    expect(tables.clearCount).toBe(1);
    expect(Array.from(tables.keys())).toEqual(expectedOrder);
    expect(tables.get("Table4999")?.headers.has("Column")).toBe(true);

    manager.applyHistoryChanges(patches[0]!, "redo");
    expect(tables.clearCount).toBe(2);
    expect(Array.from(tables.keys())).toEqual(expectedOrder);
    expect(tables.get("Table4999")?.headers.has("Column4999")).toBe(true);
  });

  test("reset deltas replay exact ordering and empty workbook buckets", () => {
    let capture = false;
    const captured: TableMutation[] = [];
    const manager = new TableManager(new WorkbookManager(), (changes) => {
      if (capture) {
        captured.push(...changes);
      }
    });
    manager.resetTables(
      new Map([
        [
          "Book1",
          new Map([
            ["A", table("A", "Book1")],
            ["B", table("B", "Book1", "Sheet2")],
          ]),
        ],
        ["Book2", new Map<string, TableDefinition>()],
      ])
    );
    const before = serialize(manager.toSnapshot());

    const replacement = new Map([
      ["Book2", new Map<string, TableDefinition>()],
      [
        "Book1",
        new Map([
          ["B", table("B", "Book1", "RenamedSheet")],
          ["A", table("A", "Book1")],
        ]),
      ],
      ["Book3", new Map<string, TableDefinition>()],
    ]);
    capture = true;
    manager.resetTables(replacement);
    capture = false;
    const after = serialize(manager.toSnapshot());

    expect(captured.length).toBeGreaterThan(0);
    expect(captured.every((change) => !("tables" in change))).toBe(true);

    manager.applyHistoryChanges(captured, "undo");
    expect(serialize(manager.toSnapshot())).toBe(before);

    manager.applyHistoryChanges(captured, "redo");
    expect(serialize(manager.toSnapshot())).toBe(after);
    expect(Array.from(manager.tables.keys())).toEqual([
      "Book2",
      "Book1",
      "Book3",
    ]);
    expect(Array.from(manager.tables.get("Book1")!.keys())).toEqual(["B", "A"]);
    expect(manager.tables.get("Book3")?.size).toBe(0);
  });
});
