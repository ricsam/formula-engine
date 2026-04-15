import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type {
  CellAddress,
  RangeAddress,
  SerializedCellValue,
  SpreadsheetRange,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

describe("Direct mutator invalidation", () => {
  const workbookName = "Book1";
  const sheetName = "Sheet1";
  let engine: FormulaEngine;

  const address = (
    ref: string,
    targetSheetName = sheetName,
    targetWorkbookName = workbookName
  ): CellAddress => ({
    workbookName: targetWorkbookName,
    sheetName: targetSheetName,
    ...parseCellReference(ref),
  });

  const spreadsheetRange = (
    startRef: string,
    endRef: string
  ): SpreadsheetRange => {
    const start = parseCellReference(startRef);
    const end = parseCellReference(endRef);
    return {
      start: { col: start.colIndex, row: start.rowIndex },
      end: {
        col: { type: "number", value: end.colIndex },
        row: { type: "number", value: end.rowIndex },
      },
    };
  };

  const range = (
    startRef: string,
    endRef: string,
    targetSheetName = sheetName,
    targetWorkbookName = workbookName
  ): RangeAddress => ({
    workbookName: targetWorkbookName,
    sheetName: targetSheetName,
    range: spreadsheetRange(startRef, endRef),
  });

  const setCell = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(address(ref), content);
  };

  const cell = (
    ref: string,
    targetSheetName = sheetName,
    targetWorkbookName = workbookName
  ) => engine.getCellValue(address(ref, targetSheetName, targetWorkbookName));

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  test("set cell, set sheet, and clear range invalidate dependent formulas", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 1],
        ["A2", 2],
        ["B1", "=A1+1"],
        ["B2", "=SUM(A1:A2)"],
      ])
    );

    expect(cell("B1")).toBe(2);
    expect(cell("B2")).toBe(3);

    setCell("A1", 5);
    expect(cell("B1")).toBe(6);
    expect(cell("B2")).toBe(7);

    engine.clearSpreadsheetRange(range("A2", "A2"));
    expect(cell("B2")).toBe(5);

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", 10],
        ["A2", 20],
        ["B1", "=A1+A2"],
      ])
    );

    expect(cell("B1")).toBe(30);
  });

  test("paste, fill, move, and autofill invalidate dependent formulas", () => {
    setCell("A1", 1);
    setCell("D1", "=SUM(C1)+1");
    expect(cell("D1")).toBe(1);

    engine.pasteCells([address("A1")], address("C1"), {
      include: ["content"],
      type: "formula",
    });
    expect(cell("D1")).toBe(2);

    setCell("A2", 2);
    setCell("D2", "=SUM(C2)+1");
    expect(cell("D2")).toBe(1);

    engine.fillAreas(range("A2", "A2"), [range("C2", "C2")], {
      include: ["content"],
      type: "formula",
    });
    expect(cell("D2")).toBe(3);

    setCell("A3", 3);
    setCell("B3", "=A3+1");
    expect(cell("B3")).toBe(4);

    engine.moveCell(address("A3"), address("C3"));
    expect(cell("A3")).toBe("");
    expect(cell("B3")).toBe(4);
    expect(cell("C3")).toBe(3);

    setCell("A4", 4);
    setCell("B4", 5);
    setCell("C4", "=SUM(A4:B4)");
    expect(cell("C4")).toBe(9);

    engine.moveRange(range("A4", "B4"), address("D4"));
    expect(cell("C4")).toBe(9);
    expect(cell("D4")).toBe(4);
    expect(cell("E4")).toBe(5);

    setCell("A5", 1);
    setCell("A6", 2);
    setCell("B5", "=SUM(A5:A8)");
    expect(cell("B5")).toBe(3);

    engine.autoFill(
      { workbookName, sheetName },
      spreadsheetRange("A5", "A6"),
      [spreadsheetRange("A7", "A8")],
      "down"
    );
    expect(cell("B5")).toBe(10);
  });

  test("table add, rename, update, and remove invalidate table formulas", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "Name"],
        ["B1", "Amount"],
        ["A2", "A"],
        ["B2", 10],
        ["A3", "B"],
        ["B3", 20],
        ["D1", "=SUM(Items[Amount])"],
      ])
    );

    expect(cell("D1")).toBe(0);

    engine.addTable({
      workbookName,
      sheetName,
      tableName: "Items",
      start: "A1",
      numRows: { type: "number", value: 2 },
      numCols: 2,
    });
    expect(cell("D1")).toBe(30);

    engine.renameTable(workbookName, {
      oldName: "Items",
      newName: "Products",
    });
    expect(cell("D1")).toBe(30);

    engine.updateTable({
      workbookName,
      tableName: "Products",
      numRows: { type: "number", value: 1 },
    });
    expect(cell("D1")).toBe(10);

    engine.removeTable({ workbookName, tableName: "Products" });
    expect(cell("D1")).toBe(0);
  });

  test("named expression add, update, rename, and remove invalidate formulas", () => {
    engine.addNamedExpression({
      expressionName: "Rate",
      expression: "2",
    });
    setCell("A1", "=Rate+1");
    expect(cell("A1")).toBe(3);

    engine.updateNamedExpression({
      expressionName: "Rate",
      expression: "5",
    });
    expect(cell("A1")).toBe(6);

    engine.renameNamedExpression({
      expressionName: "Rate",
      newName: "Tax",
    });
    expect(cell("A1")).toBe(6);

    engine.removeNamedExpression({
      expressionName: "Tax",
    });
    expect(cell("A1")).toBe("#NAME?");
  });

  test("sheet and workbook lifecycle mutations invalidate external references", () => {
    setCell("A1", "=Sheet2!A1+1");
    expect(typeof cell("A1")).toBe("string");

    engine.addSheet({ workbookName, sheetName: "Sheet2" });
    engine.setCellContent(address("A1", "Sheet2"), 4);
    expect(cell("A1")).toBe(5);

    engine.renameSheet({
      workbookName,
      sheetName: "Sheet2",
      newSheetName: "Data",
    });
    expect(cell("A1")).toBe(5);

    engine.removeSheet({ workbookName, sheetName: "Data" });
    expect(typeof cell("A1")).toBe("string");

    setCell("B1", "=[Book2]Sheet1!A1+1");
    setCell("C1", "=[Clone]Sheet1!A1+1");
    expect(typeof cell("B1")).toBe("string");
    expect(typeof cell("C1")).toBe("string");

    engine.addWorkbook("Book2");
    engine.addSheet({ workbookName: "Book2", sheetName: "Sheet1" });
    engine.setCellContent(address("A1", "Sheet1", "Book2"), 10);
    expect(cell("B1")).toBe(11);

    engine.renameWorkbook({
      workbookName: "Book2",
      newWorkbookName: "RenamedBook",
    });
    expect(cell("B1")).toBe(11);

    engine.cloneWorkbook("RenamedBook", "Clone");
    expect(cell("C1")).toBe(11);

    engine.removeWorkbook("Clone");
    expect(typeof cell("C1")).toBe("string");
  });
});
