import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRange,
} from "../../../src/core/types";
import type { WorkbookData } from "../../../src/core/workbook-data";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Imported";

function cell(sheetName: string, ref: string): CellAddress {
  const { rowIndex, colIndex } = parseCellReference(ref);
  return { workbookName, sheetName, rowIndex, colIndex };
}

function range(start: string, end: string): SpreadsheetRange {
  const from = parseCellReference(start);
  const to = parseCellReference(end);
  return {
    start: { col: from.colIndex, row: from.rowIndex },
    end: {
      col: { type: "number", value: to.colIndex },
      row: { type: "number", value: to.rowIndex },
    },
  };
}

describe("addWorkbook with WorkbookData", () => {
  test("adds an empty workbook when no data is supplied", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });

    expect(engine.hasWorkbook(workbookName)).toBe(true);
    expect(engine.getOrderedSheetNames(workbookName)).toEqual([]);
  });

  test("creates sheets in array order and evaluates imported formulas", () => {
    const engine = FormulaEngine.buildEmpty();
    const data: WorkbookData = {
      sheets: [
        {
          name: "Data",
          content: new Map<string, SerializedCellValue>([
            ["A1", 10],
            ["A2", 32],
            ["A3", "=SUM(A1:A2)"],
          ]),
        },
        {
          name: "Report",
          content: { B1: "=Data!A3 * 2" },
        },
      ],
    };

    engine.addWorkbook({ workbookName, data });

    expect(engine.getOrderedSheetNames(workbookName)).toEqual([
      "Data",
      "Report",
    ]);
    expect(engine.getCellValue(cell("Data", "A3"))).toBe(42);
    expect(engine.getCellValue(cell("Report", "B1"))).toBe(84);
  });

  test("accepts plain-object content so data survives a JSON round trip", () => {
    const engine = FormulaEngine.buildEmpty();
    const data: WorkbookData = {
      sheets: [{ name: "Sheet1", content: { A1: 5, B1: "=A1+1" } }],
    };

    const roundTripped = JSON.parse(JSON.stringify(data)) as WorkbookData;
    engine.addWorkbook({ workbookName, data: roundTripped });

    expect(engine.getCellValue(cell("Sheet1", "B1"))).toBe(6);
  });

  test("creates tables whose headers resolve against imported content", () => {
    const engine = FormulaEngine.buildEmpty();
    const data: WorkbookData = {
      sheets: [
        {
          name: "Sheet1",
          content: new Map<string, SerializedCellValue>([
            ["A1", "Item"],
            ["B1", "Amount"],
            ["A2", "pens"],
            ["B2", 3],
            ["A3", "pads"],
            ["B3", 4],
            ["D1", "=SUM(Sales[Amount])"],
          ]),
        },
      ],
      tables: [
        {
          name: "Sales",
          sheetName: "Sheet1",
          start: "A1",
          numRows: { type: "number", value: 2 },
          numCols: 2,
        },
      ],
    };

    engine.addWorkbook({ workbookName, data });

    const table = engine.getTable({ workbookName, tableName: "Sales" });
    expect(table).toBeDefined();
    expect(Array.from(table!.headers.keys())).toEqual(["Item", "Amount"]);
    expect(engine.getCellValue(cell("Sheet1", "D1"))).toBe(7);
  });

  test("imports workbook- and sheet-scoped named expressions", () => {
    const engine = FormulaEngine.buildEmpty();
    const data: WorkbookData = {
      sheets: [{ name: "Sheet1", content: { A1: "=TaxRate", A2: "=Local" } }],
      namedExpressions: [
        { name: "TaxRate", expression: "0.25" },
        { name: "Local", expression: "7", sheetName: "Sheet1" },
      ],
    };

    engine.addWorkbook({ workbookName, data });

    expect(engine.getCellValue(cell("Sheet1", "A1"))).toBe(0.25);
    expect(engine.getCellValue(cell("Sheet1", "A2"))).toBe(7);
  });

  test("imports styles, conditional styles and data types", () => {
    const engine = FormulaEngine.buildEmpty();
    const data: WorkbookData = {
      sheets: [{ name: "Sheet1", content: { A1: 1, A2: 200 } }],
      cellStyles: [
        {
          areas: [{ sheetName: "Sheet1", range: range("A1", "A2") }],
          style: { bold: true, backgroundColor: "#ffeeaa" },
        },
      ],
      conditionalStyles: [
        {
          areas: [{ sheetName: "Sheet1", range: range("A1", "A2") }],
          condition: {
            type: "formula",
            formula: "=A1>100",
            color: { l: 50, c: 60, h: 20 },
          },
        },
      ],
      cellDataTypes: [
        {
          areas: [{ sheetName: "Sheet1", range: range("A1", "A2") }],
          dataType: "number",
        },
      ],
    };

    engine.addWorkbook({ workbookName, data });

    expect(engine.getCellStyle(cell("Sheet1", "A1"))).toMatchObject({
      bold: true,
      backgroundColor: "#ffeeaa",
    });
    expect(engine.getAllConditionalStyles()).toHaveLength(1);
    expect(engine.getCellDataType(cell("Sheet1", "A1"))).toBe("number");
  });

  test("imports cell, sheet and workbook metadata", () => {
    const engine = FormulaEngine.buildEmpty<{
      cell: { note: string };
      sheet: { frozenRows: number };
      workbook: { source: string };
    }>();
    engine.addWorkbook({
      workbookName,
      data: {
        sheets: [
          {
            name: "Sheet1",
            content: { A1: 1 },
            cellMetadata: { A1: { note: "from excel" } },
            sheetMetadata: { frozenRows: 1 },
          },
        ],
        workbookMetadata: { source: "book.xlsx" },
      },
    });

    expect(engine.getCellMetadata(cell("Sheet1", "A1"))).toEqual({
      note: "from excel",
    });
    expect(
      engine.getSheetMetadata({ workbookName, sheetName: "Sheet1" })
    ).toEqual({ frozenRows: 1 });
    expect(engine.getWorkbookMetadata(workbookName)).toEqual({
      source: "book.xlsx",
    });
  });

  test("imports the whole workbook as a single undo step", () => {
    const engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({
      workbookName,
      data: {
        sheets: [
          { name: "Sheet1", content: { A1: 1 } },
          { name: "Sheet2", content: { A1: 2 } },
        ],
        namedExpressions: [{ name: "Rate", expression: "0.1" }],
      },
    });

    expect(engine.hasWorkbook(workbookName)).toBe(true);
    expect(engine.canUndo()).toBe(true);

    engine.undo();

    expect(engine.hasWorkbook(workbookName)).toBe(false);
  });
});
