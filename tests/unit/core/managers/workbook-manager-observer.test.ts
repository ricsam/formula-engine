import { describe, expect, test } from "bun:test";
import type { EvaluationManager } from "../../../../src/core/managers/evaluation-manager";
import { CopyManager } from "../../../../src/core/managers/copy-manager";
import { RangeMetadataManager } from "../../../../src/core/managers/range-metadata-manager";
import { StyleManager } from "../../../../src/core/managers/style-manager";
import {
  WorkbookManager,
  type WorkbookDataChange,
} from "../../../../src/core/managers/workbook-manager";
import type {
  CellAddress,
  RangeAddress,
  SerializedCellValue,
} from "../../../../src/core/types";
import { parseCellReference } from "../../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";

function address(reference: string): CellAddress {
  return {
    workbookName,
    sheetName,
    ...parseCellReference(reference),
  };
}

function range(start: string, end = start): RangeAddress {
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

function setupObservedWorkbook() {
  const batches: WorkbookDataChange[][] = [];
  const manager = new WorkbookManager((changes) => {
    batches.push([...changes]);
  });
  manager.addWorkbook(workbookName);
  manager.addSheet({ workbookName, sheetName });
  return { batches, manager };
}

describe("WorkbookManager mutation observer", () => {
  test("normalizes empty cell content and skips actual no-ops", () => {
    const { batches, manager } = setupObservedWorkbook();

    manager.setCellContent(address("A1"), "");
    manager.setCellContent(address("A1"), undefined);
    expect(batches).toHaveLength(0);

    manager.setCellContent(address("A1"), 1);
    expect(batches).toEqual([
      [
        {
          kind: "cell-content",
          address: address("A1"),
          before: undefined,
          after: 1,
        },
      ],
    ]);

    manager.setCellContent(address("A1"), 1);
    expect(batches).toHaveLength(1);

    manager.setCellContent(address("A1"), "");
    expect(batches[1]).toEqual([
      {
        kind: "cell-content",
        address: address("A1"),
        before: 1,
        after: undefined,
      },
    ]);
    expect(
      manager.getSheetSerialized({ workbookName, sheetName }).has("A1")
    ).toBe(false);
  });

  test("reports one sparse batch for a sheet rebuild", () => {
    const { batches, manager } = setupObservedWorkbook();
    manager.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["B1", 2],
        ["C1", 3],
      ])
    );
    batches.length = 0;

    const replacement = new Map<string, SerializedCellValue>([
      ["A1", 1],
      ["B1", 20],
      ["C1", ""],
      ["D1", 4],
    ]);
    manager.setSheetContent({ workbookName, sheetName }, replacement);

    expect(batches).toHaveLength(1);
    expect(batches[0]).toEqual([
      {
        kind: "cell-content",
        address: address("B1"),
        before: 2,
        after: 20,
      },
      {
        kind: "cell-content",
        address: address("C1"),
        before: 3,
        after: undefined,
      },
      {
        kind: "cell-content",
        address: address("D1"),
        before: undefined,
        after: 4,
      },
    ]);
    expect(
      manager
        .getSheetIndexes({ workbookName, sheetName })
        .rowGroups.get(0)
        ?.map((entry) => entry.key)
    ).toEqual(["A1", "B1", "D1"]);

    batches.length = 0;
    manager.setSheetContent({ workbookName, sheetName }, replacement);
    expect(batches).toHaveLength(0);
  });

  test("batches exact changes from every formula rewrite path", () => {
    const { batches, manager } = setupObservedWorkbook();
    const formulas = new Map<string, SerializedCellValue>([
      ["A1", "=Old+1"],
      ["B1", "=Old+2"],
      ["C1", 3],
    ]);
    manager.setSheetContent({ workbookName, sheetName }, formulas);
    batches.length = 0;

    expect(
      manager.updateAllFormulas((formula) => formula.replace("Old", "New"))
    ).toEqual([address("A1"), address("B1")]);
    expect(batches).toHaveLength(1);
    expect(batches[0]?.map((change) => change.kind)).toEqual([
      "cell-content",
      "cell-content",
    ]);
    expect(batches[0]?.[0]).toMatchObject({
      address: address("A1"),
      before: "=Old+1",
      after: "=New+1",
    });

    manager.setSheetContent({ workbookName, sheetName }, formulas);
    batches.length = 0;
    manager.updateFormulasExcluding(
      new Set([`${workbookName}:${sheetName}:0:0`]),
      (formula) => formula.replace("Old", "Other")
    );
    expect(batches).toHaveLength(1);
    expect(batches[0]).toEqual([
      {
        kind: "cell-content",
        address: address("B1"),
        before: "=Old+2",
        after: "=Other+2",
      },
    ]);

    manager.setSheetContent({ workbookName, sheetName }, formulas);
    batches.length = 0;
    manager.updateFormulasForWorkbook(workbookName, (formula) =>
      formula.replace("Old", "Scoped")
    );
    expect(batches).toHaveLength(1);
    expect(batches[0]).toHaveLength(2);
    expect(batches[0]?.[1]).toMatchObject({
      address: address("B1"),
      before: "=Old+2",
      after: "=Scoped+2",
    });

    batches.length = 0;
    manager.updateAllFormulas((formula) => formula);
    expect(batches).toHaveLength(0);
  });

  test("reports cell, sheet, and workbook metadata changes", () => {
    const { batches, manager } = setupObservedWorkbook();
    const cellMetadata = { comment: "hello" };
    const sheetMetadata = { frozenRows: 1 };
    const workbookMetadata = { theme: "dark" };

    manager.setCellMetadata(address("A1"), cellMetadata);
    manager.setCellMetadata(address("A1"), cellMetadata);
    manager.setSheetMetadata({ workbookName, sheetName }, sheetMetadata);
    manager.setSheetMetadata({ workbookName, sheetName }, sheetMetadata);
    manager.setWorkbookMetadata(workbookName, workbookMetadata);
    manager.setWorkbookMetadata(workbookName, workbookMetadata);

    expect(batches).toEqual([
      [
        {
          kind: "cell-metadata",
          address: address("A1"),
          before: undefined,
          after: cellMetadata,
        },
      ],
      [
        {
          kind: "sheet-metadata",
          workbookName,
          sheetName,
          before: undefined,
          after: sheetMetadata,
        },
      ],
      [
        {
          kind: "workbook-metadata",
          workbookName,
          before: undefined,
          after: workbookMetadata,
        },
      ],
    ]);

    batches.length = 0;
    manager.setCellMetadata(address("A1"), undefined);
    expect(batches[0]?.[0]).toMatchObject({
      kind: "cell-metadata",
      before: cellMetadata,
      after: undefined,
    });
  });

  test("copy, fill, and cut report batches and maintain indexes", () => {
    const { batches, manager } = setupObservedWorkbook();
    const evaluationManager = {
      getCellEvaluationResult: () => undefined,
    } as unknown as EvaluationManager;
    const copyManager = new CopyManager(
      manager,
      evaluationManager,
      new StyleManager(evaluationManager),
      new RangeMetadataManager()
    );

    manager.setCellContent(address("A1"), 7);
    manager.setCellMetadata(address("A1"), { comment: "source" });
    batches.length = 0;
    copyManager.pasteCells([address("A1")], address("C1"), {
      include: ["content", "cellMetadata"],
      type: "formula",
    });
    expect(batches).toHaveLength(1);
    expect(batches[0]?.map((change) => change.kind)).toEqual([
      "cell-content",
      "cell-metadata",
    ]);

    batches.length = 0;
    copyManager.fillAreas(range("A1"), [range("D1")], {
      include: ["content"],
      type: "formula",
    });
    expect(batches).toHaveLength(1);
    expect(batches[0]).toEqual([
      {
        kind: "cell-content",
        address: address("D1"),
        before: undefined,
        after: 7,
      },
    ]);

    manager.setCellContent(address("B1"), "=A1");
    batches.length = 0;
    copyManager.pasteCells([address("A1")], address("E1"), {
      cut: true,
      include: ["content"],
      type: "formula",
    });
    expect(batches).toHaveLength(1);
    expect(batches[0]).toEqual([
      {
        kind: "cell-content",
        address: address("A1"),
        before: 7,
        after: undefined,
      },
      {
        kind: "cell-content",
        address: address("B1"),
        before: "=A1",
        after: "=E1",
      },
      {
        kind: "cell-content",
        address: address("E1"),
        before: undefined,
        after: 7,
      },
    ]);

    const rowKeys = manager
      .getSheetIndexes({ workbookName, sheetName })
      .rowGroups.get(0)
      ?.map((entry) => entry.key);
    expect(rowKeys).toEqual(["B1", "C1", "D1", "E1"]);
  });
});
