import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";

const workbookName = "Book";
const sheetName = "Data";

function buildEngine() {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  engine.clearUndoRedoHistory();
  return engine;
}

describe("resource events", () => {
  test("notifies matching subscriptions and supports unsubscribe", () => {
    const engine = buildEngine();
    engine.addSheet({ workbookName, sheetName: "Other" });
    const sheetNames: string[] = [];
    const workbookNames: string[] = [];

    const unsubscribeSheet = engine.onSheetRename(
      { workbookName, sheetName },
      (newSheetName) => sheetNames.push(newSheetName)
    );
    const unsubscribeWorkbook = engine.onWorkbookRename(
      workbookName,
      (newWorkbookName) => workbookNames.push(newWorkbookName)
    );

    engine.renameSheet({
      workbookName,
      sheetName: "Other",
      newSheetName: "Unrelated",
    });
    expect(sheetNames).toEqual([]);

    engine.renameSheet({
      workbookName,
      sheetName,
      newSheetName: "RenamedData",
    });
    engine.renameWorkbook({
      workbookName,
      newWorkbookName: "RenamedBook",
    });

    expect(sheetNames).toEqual(["RenamedData"]);
    expect(workbookNames).toEqual(["RenamedBook"]);

    unsubscribeSheet();
    unsubscribeWorkbook();
    engine.renameSheet({
      workbookName: "RenamedBook",
      sheetName: "RenamedData",
      newSheetName: "FinalData",
    });
    engine.renameWorkbook({
      workbookName: "RenamedBook",
      newWorkbookName: "FinalBook",
    });

    expect(sheetNames).toEqual(["RenamedData"]);
    expect(workbookNames).toEqual(["RenamedBook"]);
  });

  test("defers notifications until commit and follows identities", () => {
    const engine = buildEngine();
    const sheetNames: string[] = [];
    const workbookNames: string[] = [];

    engine.onSheetRename(
      { workbookName, sheetName },
      (newSheetName) => sheetNames.push(newSheetName)
    );
    engine.onWorkbookRename(workbookName, (newWorkbookName) =>
      workbookNames.push(newWorkbookName)
    );

    engine.transact(() => {
      engine.renameWorkbook({
        workbookName,
        newWorkbookName: "RenamedBook",
      });
      engine.renameSheet({
        workbookName: "RenamedBook",
        sheetName,
        newSheetName: "Summary",
      });
      engine.renameWorkbook({
        workbookName: "RenamedBook",
        newWorkbookName: "FinalBook",
      });
      engine.renameSheet({
        workbookName: "FinalBook",
        sheetName: "Summary",
        newSheetName: "Results",
      });

      expect(workbookNames).toEqual([]);
      expect(sheetNames).toEqual([]);
    });

    expect(workbookNames).toEqual(["RenamedBook", "FinalBook"]);
    expect(sheetNames).toEqual(["Summary", "Results"]);
  });

  test("notifies for undo and redo in replay order", () => {
    const engine = buildEngine();
    const events: string[] = [];

    engine.onSheetRename(
      { workbookName, sheetName },
      (newSheetName) => events.push(`sheet:${newSheetName}`)
    );
    engine.onWorkbookRename(workbookName, (newWorkbookName) =>
      events.push(`workbook:${newWorkbookName}`)
    );

    engine.transact(() => {
      engine.renameWorkbook({
        workbookName,
        newWorkbookName: "RenamedBook",
      });
      engine.renameSheet({
        workbookName: "RenamedBook",
        sheetName,
        newSheetName: "Summary",
      });
    });
    expect(events).toEqual(["workbook:RenamedBook", "sheet:Summary"]);

    events.length = 0;
    expect(engine.undo()).toBe(true);
    expect(events).toEqual(["sheet:Data", "workbook:Book"]);

    events.length = 0;
    expect(engine.redo()).toBe(true);
    expect(events).toEqual(["workbook:RenamedBook", "sheet:Summary"]);
  });

  test("does not notify when a transaction rolls back", () => {
    const engine = buildEngine();
    const sheetNames: string[] = [];

    engine.onSheetRename(
      { workbookName, sheetName },
      (newSheetName) => sheetNames.push(newSheetName)
    );

    expect(() =>
      engine.transact(() => {
        engine.renameSheet({
          workbookName,
          sheetName,
          newSheetName: "Temporary",
        });
        throw new Error("rollback");
      })
    ).toThrow("rollback");

    expect(sheetNames).toEqual([]);
    expect(engine.hasSheet({ workbookName, sheetName })).toBe(true);

    engine.renameSheet({
      workbookName,
      sheetName,
      newSheetName: "Committed",
    });
    expect(sheetNames).toEqual(["Committed"]);
  });

  test("notifies sheet and workbook deletion listeners after commit", () => {
    const engine = buildEngine();
    const events: string[] = [];

    engine.onSheetDelete({ workbookName, sheetName }, () =>
      events.push("sheet-delete")
    );
    engine.onWorkbookDelete(workbookName, () =>
      events.push("workbook-delete")
    );

    engine.transact(() => {
      engine.renameWorkbook({
        workbookName,
        newWorkbookName: "RenamedBook",
      });
      engine.renameSheet({
        workbookName: "RenamedBook",
        sheetName,
        newSheetName: "Summary",
      });
      engine.removeWorkbook("RenamedBook");
      expect(events).toEqual([]);
    });

    expect(events).toEqual(["sheet-delete", "workbook-delete"]);

    expect(engine.undo()).toBe(true);
    expect(events).toEqual(["sheet-delete", "workbook-delete"]);

    expect(engine.redo()).toBe(true);
    expect(events).toEqual([
      "sheet-delete",
      "workbook-delete",
      "sheet-delete",
      "workbook-delete",
    ]);
  });

  test("notifies when undo removes a newly added resource", () => {
    const engine = buildEngine();
    engine.addSheet({ workbookName, sheetName: "Temporary" });
    let deletes = 0;
    engine.onSheetDelete(
      { workbookName, sheetName: "Temporary" },
      () => deletes++
    );

    expect(engine.undo()).toBe(true);
    expect(deletes).toBe(1);

    expect(engine.redo()).toBe(true);
    expect(deletes).toBe(1);
  });

  test("does not notify deletion listeners for rolled-back deletes", () => {
    const engine = buildEngine();
    let deletes = 0;
    const unsubscribe = engine.onSheetDelete(
      { workbookName, sheetName },
      () => deletes++
    );

    expect(() =>
      engine.transact(() => {
        engine.removeSheet({ workbookName, sheetName });
        throw new Error("rollback");
      })
    ).toThrow("rollback");

    expect(deletes).toBe(0);
    expect(engine.hasSheet({ workbookName, sheetName })).toBe(true);

    engine.removeSheet({ workbookName, sheetName });
    expect(deletes).toBe(1);

    unsubscribe();
    expect(engine.undo()).toBe(true);
    expect(engine.redo()).toBe(true);
    expect(deletes).toBe(1);
  });
});
