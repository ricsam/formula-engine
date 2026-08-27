import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../../src/core/engine";
import { WorkbookManager } from "../../../../src/core/managers/workbook-manager";
import type { SerializedCellValue } from "../../../../src/core/types";

describe("WorkbookManager bulk indexing", () => {
  test("builds a large sheet in one collection-and-sort pass", () => {
    const manager = new WorkbookManager();
    manager.addWorkbook("Book");
    manager.addSheet({ workbookName: "Book", sheetName: "Sheet" });

    const content = new Map<string, SerializedCellValue>();
    for (let row = 1; row <= 50_000; row++) {
      content.set(`A${row}`, row);
    }

    const startedAt = performance.now();
    manager.setSheetContent(
      { workbookName: "Book", sheetName: "Sheet" },
      content
    );
    const elapsedMs = performance.now() - startedAt;

    const indexes = manager.getSheetIndexes({
      workbookName: "Book",
      sheetName: "Sheet",
    });
    expect(indexes.cellsSortedByRow).toHaveLength(50_000);
    expect(indexes.cellsSortedByCol).toHaveLength(50_000);
    expect(indexes.rowGroups.size).toBe(50_000);
    expect(indexes.colGroups.get(0)).toHaveLength(50_000);
    expect(indexes.cellsSortedByRow[0]?.key).toBe("A1");
    expect(indexes.cellsSortedByRow.at(-1)?.key).toBe("A50000");

    // The former per-cell scan/splice implementation takes tens of seconds at
    // this size; leave ample headroom for slower CI while guarding complexity.
    expect(elapsedMs).toBeLessThan(2_000);
  });

  test("replays a large sheet patch with one index rebuild per direction", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook("Book");
    engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });
    engine.clearUndoRedoHistory();

    const content = new Map<string, SerializedCellValue>();
    for (let row = 1; row <= 20_000; row++) {
      content.set(`A${row}`, row);
    }
    engine.setSheetContent(
      { workbookName: "Book", sheetName: "Sheet" },
      content
    );

    const undoStartedAt = performance.now();
    expect(engine.undo()).toBe(true);
    const undoElapsedMs = performance.now() - undoStartedAt;
    expect(
      engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
        .size
    ).toBe(0);

    const redoStartedAt = performance.now();
    expect(engine.redo()).toBe(true);
    const redoElapsedMs = performance.now() - redoStartedAt;
    expect(
      engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
        .size
    ).toBe(20_000);

    // Per-cell flat-index splicing was quadratic and took seconds at this size.
    expect(undoElapsedMs).toBeLessThan(1_500);
    expect(redoElapsedMs).toBeLessThan(1_500);
  });

  test("groups append-only paste replay into one linear sheet rebuild", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook("Book");
    engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });

    const source: Array<{
      workbookName: string;
      sheetName: string;
      colIndex: number;
      rowIndex: number;
    }> = [];
    const content = new Map<string, SerializedCellValue>();
    for (let rowIndex = 0; rowIndex < 2_000; rowIndex++) {
      content.set(`A${rowIndex + 1}`, rowIndex);
      source.push({
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex,
      });
    }
    engine.setSheetContent(
      { workbookName: "Book", sheetName: "Sheet" },
      content
    );
    engine.clearUndoRedoHistory();
    engine.pasteCells(
      source,
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 1,
        rowIndex: 0,
      },
      { cut: false, type: "formula", include: ["content"] }
    );

    const undoStartedAt = performance.now();
    expect(engine.undo()).toBe(true);
    const undoElapsedMs = performance.now() - undoStartedAt;
    expect(
      engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
        .size
    ).toBe(2_000);
    expect(undoElapsedMs).toBeLessThan(1_000);

    expect(engine.redo()).toBe(true);
    expect(
      engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
        .size
    ).toBe(4_000);
  });

  test("replays sequential cut deletions with one ordered reconstruction", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook("Book");
    engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });

    const source: Array<{
      workbookName: string;
      sheetName: string;
      colIndex: number;
      rowIndex: number;
    }> = [];
    const content = new Map<string, SerializedCellValue>();
    for (let rowIndex = 0; rowIndex < 2_000; rowIndex++) {
      content.set(`A${rowIndex + 1}`, rowIndex);
      source.push({
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex,
      });
    }
    engine.setSheetContent(
      { workbookName: "Book", sheetName: "Sheet" },
      content
    );
    engine.clearUndoRedoHistory();
    engine.pasteCells(
      source,
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 1,
        rowIndex: 0,
      },
      { cut: true, type: "formula", include: ["content"] }
    );
    const afterCut = Array.from(
      engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
    );

    const undoStartedAt = performance.now();
    expect(engine.undo()).toBe(true);
    const undoElapsedMs = performance.now() - undoStartedAt;
    expect(
      Array.from(
        engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
      )
    ).toEqual(Array.from(content));
    expect(undoElapsedMs).toBeLessThan(1_000);

    expect(engine.redo()).toBe(true);
    expect(
      Array.from(
        engine.getSheetSerialized({ workbookName: "Book", sheetName: "Sheet" })
      )
    ).toEqual(afterCut);
  });

  test("sparse value and tail replay preserve collection instances", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook("Book");
    engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });
    engine.setSheetContent(
      { workbookName: "Book", sheetName: "Sheet" },
      new Map([
        ["A1", 1],
        ["A2", 2],
      ])
    );
    engine._workbookManager.setCellMetadata(
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex: 0,
      },
      { label: "before" }
    );
    engine.clearUndoRedoHistory();

    const indexes = engine._workbookManager.getSheetIndexes({
      workbookName: "Book",
      sheetName: "Sheet",
    });
    const metadata = engine.getSheetMetadataSerialized({
      workbookName: "Book",
      sheetName: "Sheet",
    });

    engine.setCellContent(
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex: 0,
      },
      10
    );
    expect(engine.undo()).toBe(true);
    expect(
      engine._workbookManager.getSheetIndexes({
        workbookName: "Book",
        sheetName: "Sheet",
      })
    ).toBe(indexes);

    engine.setCellMetadata(
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex: 0,
      },
      { label: "after" }
    );
    expect(engine.undo()).toBe(true);
    expect(
      engine.getSheetMetadataSerialized({
        workbookName: "Book",
        sheetName: "Sheet",
      })
    ).toBe(metadata);

    engine.setCellContent(
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex: 2,
      },
      3
    );
    expect(engine.undo()).toBe(true);
    expect(engine.redo()).toBe(true);
    expect(
      engine._workbookManager.getSheetIndexes({
        workbookName: "Book",
        sheetName: "Sheet",
      })
    ).toBe(indexes);

    engine.setCellContent(
      {
        workbookName: "Book",
        sheetName: "Sheet",
        colIndex: 0,
        rowIndex: 2,
      },
      undefined
    );
    expect(engine.undo()).toBe(true);
    expect(engine.redo()).toBe(true);
    expect(
      engine._workbookManager.getSheetIndexes({
        workbookName: "Book",
        sheetName: "Sheet",
      })
    ).toBe(indexes);
  });
});
