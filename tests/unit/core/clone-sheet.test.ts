import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type {
  CellAddress,
  RangeAddress,
  SerializedCellValue,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sourceSheetName = "Source";
const targetSheetName = "Source Copy";

type TestMetadata = {
  cell: { label: string; nested: { value: number } };
  sheet: { title: string; nested: { frozenRows: number } };
  range: { label: string };
};

function cell(ref: string, sheetName = sourceSheetName): CellAddress {
  return {
    workbookName,
    sheetName,
    ...parseCellReference(ref),
  };
}

function range(
  start: string,
  end = start,
  sheetName = sourceSheetName
): RangeAddress {
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

function buildEngine(): FormulaEngine<TestMetadata> {
  const engine = FormulaEngine.buildEmpty<TestMetadata>();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName: sourceSheetName });
  engine.addSheet({ workbookName, sheetName: "Other" });
  return engine;
}

describe("cloneSheet", () => {
  test("clones complete sheet state and rewrites self references", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName: sourceSheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "Amount"],
        ["A2", 5],
        ["B1", "=Source!A2"],
        ["C1", "=SUM(SourceTable[Amount])"],
        ["D1", "=LOCAL_TOTAL"],
      ])
    );
    engine.setCellMetadata(cell("A2"), {
      label: "source cell",
      nested: { value: 1 },
    });
    engine.setSheetMetadata(
      { workbookName, sheetName: sourceSheetName },
      { title: "Source sheet", nested: { frozenRows: 1 } }
    );
    engine.addNamedExpression({
      workbookName,
      sheetName: sourceSheetName,
      expressionName: "LOCAL_TOTAL",
      expression: "Source!A2+SUM(SourceTable[Amount])",
    });
    engine.addTable({
      workbookName,
      sheetName: sourceSheetName,
      tableName: "SourceTable",
      start: "A1",
      numRows: { type: "number", value: 1 },
      numCols: 1,
    });
    engine.addCellStyle({
      areas: [range("A1", "D2"), range("A1", "A1", "Other")],
      style: { bold: true },
    });
    engine.addConditionalStyle({
      areas: [range("A2"), range("B2", "B2", "Other")],
      condition: {
        type: "formula",
        formula: "Source!A2>0",
        color: { l: 70, c: 30, h: 120 },
      },
    });
    engine.addCellDataType({
      areas: [range("A2"), range("C3", "C3", "Other")],
      dataType: "number",
    });
    engine.addRangeMetadata({
      id: "source-range-metadata",
      areas: [range("A1", "A2"), range("A1", "A1", "Other")],
      metadata: { label: "source range" },
    });

    const cloned = engine.cloneSheet({
      workbookName,
      sheetName: sourceSheetName,
      newSheetName: targetSheetName,
    });

    expect(cloned.name).toBe(targetSheetName);
    expect(engine.getOrderedSheetNames(workbookName)).toEqual([
      sourceSheetName,
      "Other",
      targetSheetName,
    ]);
    expect(
      engine.getSheetSerialized({
        workbookName,
        sheetName: targetSheetName,
      })
    ).toEqual(
      new Map<string, SerializedCellValue>([
        ["A1", "Amount"],
        ["A2", 5],
        ["B1", "='Source Copy'!A2"],
        ["C1", "=SUM(SourceTable_2[Amount])"],
        ["D1", "=LOCAL_TOTAL"],
      ])
    );
    expect(engine.getCellValue(cell("B1", targetSheetName))).toBe(5);
    expect(engine.getCellValue(cell("C1", targetSheetName))).toBe(5);
    expect(engine.getCellValue(cell("D1", targetSheetName))).toBe(10);

    expect(engine.getCellMetadata(cell("A2", targetSheetName))).toEqual({
      label: "source cell",
      nested: { value: 1 },
    });
    expect(
      engine.getSheetMetadata({
        workbookName,
        sheetName: targetSheetName,
      })
    ).toEqual({ title: "Source sheet", nested: { frozenRows: 1 } });
    expect(
      engine._namedExpressionManager
        .getNamedExpressions()
        .sheetExpressions.get(workbookName)
        ?.get(targetSheetName)
        ?.get("LOCAL_TOTAL")?.expression
    ).toBe("'Source Copy'!A2+SUM(SourceTable_2[Amount])");

    expect(engine.getTable({ workbookName, tableName: "SourceTable_2" })).toMatchObject({
      name: "SourceTable_2",
      sheetName: targetSheetName,
      workbookName,
    });
    expect(engine.getCellStyle(cell("A2", targetSheetName))).toMatchObject({
      bold: true,
    });
    expect(engine.getCellDataType(cell("A2", targetSheetName))).toBe("number");
    expect(
      engine
        .getAllConditionalStyles()
        .find((style) =>
          style.areas.some((area) => area.sheetName === targetSheetName)
        )
    ).toEqual({
      areas: [range("A2", "A2", targetSheetName)],
      condition: {
        type: "formula",
        formula: "'Source Copy'!A2>0",
        color: { l: 70, c: 30, h: 120 },
      },
    });

    const clonedRangeMetadata = engine
      .getAllRangeMetadata()
      .find((entry) =>
        entry.areas.some((area) => area.sheetName === targetSheetName)
      );
    expect(clonedRangeMetadata).toMatchObject({
      areas: [range("A1", "A2", targetSheetName)],
      metadata: { label: "source range" },
    });
    expect(clonedRangeMetadata?.id).not.toBe("source-range-metadata");

    const targetCellMetadata = engine.getCellMetadata(
      cell("A2", targetSheetName)
    );
    targetCellMetadata!.nested.value = 2;
    expect(engine.getCellMetadata(cell("A2"))).toEqual({
      label: "source cell",
      nested: { value: 1 },
    });

    const targetSheetMetadata = engine.getSheetMetadata({
      workbookName,
      sheetName: targetSheetName,
    });
    targetSheetMetadata!.nested.frozenRows = 2;
    expect(
      engine.getSheetMetadata({
        workbookName,
        sheetName: sourceSheetName,
      })
    ).toEqual({ title: "Source sheet", nested: { frozenRows: 1 } });
  });

  test("uses the next available table suffix", () => {
    const engine = buildEngine();
    engine.setSheetContent(
      { workbookName, sheetName: sourceSheetName },
      new Map([["A1", "Value"]])
    );
    engine.addTable({
      workbookName,
      sheetName: sourceSheetName,
      tableName: "Table",
      start: "A1",
      numRows: { type: "number", value: 0 },
      numCols: 1,
    });
    engine.setSheetContent(
      { workbookName, sheetName: "Other" },
      new Map([["A1", "Value"]])
    );
    engine.addTable({
      workbookName,
      sheetName: "Other",
      tableName: "Table_2",
      start: "A1",
      numRows: { type: "number", value: 0 },
      numCols: 1,
    });

    engine.cloneSheet({
      workbookName,
      sheetName: sourceSheetName,
      newSheetName: targetSheetName,
    });

    expect(engine.getTable({ workbookName, tableName: "Table_3" })).toMatchObject({
      name: "Table_3",
      sheetName: targetSheetName,
    });
  });

  test("rejects a missing source or an existing target without changing state", () => {
    const engine = buildEngine();
    engine.clearUndoRedoHistory();

    expect(() =>
      engine.cloneSheet({
        workbookName,
        sheetName: "Missing",
        newSheetName: targetSheetName,
      })
    ).toThrow('Source sheet "Missing" not found');
    expect(() =>
      engine.cloneSheet({
        workbookName,
        sheetName: sourceSheetName,
        newSheetName: "Other",
      })
    ).toThrow('Target sheet "Other" already exists');

    expect(engine.hasSheet({ workbookName, sheetName: targetSheetName })).toBe(
      false
    );
    expect(engine.getUndoRedoState().canUndo).toBe(false);
  });
});
