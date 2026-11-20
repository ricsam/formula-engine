import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import {
  type SerializedCellValue,
  type TableColumnFilter,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet1";
const tableName = "Orders";

describe("Table sorting and filtering", () => {
  let engine: FormulaEngine;

  const address = (reference: string) => ({
    workbookName,
    sheetName,
    ...parseCellReference(reference),
  });
  const value = (reference: string) => engine.getCellValue(address(reference));
  const content = (reference: string) =>
    engine.getSheet({ workbookName, sheetName })?.content.get(reference);
  const column = (letter: string, from: number, to: number) =>
    Array.from({ length: to - from + 1 }, (_, index) =>
      value(`${letter}${from + index}`)
    );
  const table = () => engine.getTable({ workbookName, tableName })!;
  const hiddenRows = () =>
    [...engine.getHiddenRows({ workbookName, sheetName })].sort((a, b) => a - b);
  const filter = (columnIndex: number, columnFilter: TableColumnFilter | undefined) =>
    engine.setTableColumnFilter({
      workbookName,
      tableName,
      column: columnIndex,
      filter: columnFilter,
    });

  /**
   * Orders table at A1:D7:
   *   Region | Product | Units | Revenue (=[@Units]*10)
   */
  const setUpOrders = () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "Region"],
        ["B1", "Product"],
        ["C1", "Units"],
        ["D1", "Revenue"],
        ["A2", "North"],
        ["B2", "pear"],
        ["C2", 30],
        ["D2", "=[@Units]*10"],
        ["A3", "south"],
        ["B3", "Apple"],
        ["C3", 10],
        ["D3", "=[@Units]*10"],
        ["A4", "North"],
        ["B4", "banana"],
        ["C4", 20],
        ["D4", "=[@Units]*10"],
        ["A5", "East"],
        ["B5", "apple"],
        ["C5", 50],
        ["D5", "=[@Units]*10"],
        ["A6", "South"],
        ["B6", "Cherry"],
        ["C6", 40],
        ["D6", "=[@Units]*10"],
        ["A7", "North"],
        ["B7", "date"],
        ["C7", 60],
        ["D7", "=[@Units]*10"],
      ])
    );
    engine.addTable({
      workbookName,
      sheetName,
      tableName,
      start: "A1",
      numRows: { type: "number", value: 6 },
      numCols: 4,
    });
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("sortTable", () => {
    test("sorts text without regard to case and keeps structured formulas working", () => {
      setUpOrders();
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 1, direction: "ascending" }],
      });

      expect(column("B", 2, 7)).toEqual([
        "Apple",
        "apple",
        "banana",
        "Cherry",
        "date",
        "pear",
      ]);
      expect(column("C", 2, 7)).toEqual([10, 50, 20, 40, 60, 30]);
      expect(column("D", 2, 7)).toEqual([100, 500, 200, 400, 600, 300]);
      expect(content("B1")).toBe("Product");
      expect(table().sort).toEqual([{ column: 1, direction: "ascending" }]);
    });

    test("sorts numbers descending and by several keys, keeping ties stable", () => {
      setUpOrders();
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 2, direction: "descending" }],
      });
      expect(column("C", 2, 7)).toEqual([60, 50, 40, 30, 20, 10]);

      engine.sortTable({
        workbookName,
        tableName,
        keys: [
          { column: 0, direction: "ascending" },
          { column: 2, direction: "ascending" },
        ],
      });
      expect(column("A", 2, 7)).toEqual([
        "East",
        "North",
        "North",
        "North",
        "south",
        "South",
      ]);
      expect(column("C", 2, 7)).toEqual([50, 20, 30, 60, 10, 40]);

      // Equal keys keep their current order.
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 0, direction: "descending" }],
      });
      expect(column("C", 2, 7)).toEqual([10, 40, 20, 30, 60, 50]);
    });

    test("orders mixed types like Excel with blanks last in both directions", () => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "Value"],
          ["A2", "b"],
          ["A3", true],
          ["A4", 2],
          ["A6", "=#N/A"],
          ["A7", false],
          ["A8", "A"],
          ["A9", -1],
        ])
      );
      engine.addTable({
        workbookName,
        sheetName,
        tableName,
        start: "A1",
        numRows: { type: "number", value: 8 },
        numCols: 1,
      });

      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 0, direction: "ascending" }],
      });
      expect(column("A", 2, 9)).toEqual([
        -1,
        2,
        "A",
        "b",
        false,
        true,
        "#N/A",
        "",
      ]);

      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 0, direction: "descending" }],
      });
      expect(column("A", 2, 9)).toEqual([
        "#N/A",
        true,
        false,
        "b",
        "A",
        2,
        -1,
        "",
      ]);
    });

    test("adjusts moved formulas like a paste and leaves outside cells alone", () => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "Item"],
          ["B1", "Qty"],
          ["C1", "Total"],
          ["A2", "c"],
          ["B2", 3],
          ["C2", "=B2*$F$1"],
          ["A3", "a"],
          ["B3", 1],
          ["C3", "=B3*$F$1+COUNTIF(A:A,A3)"],
          ["A4", "b"],
          ["B4", 2],
          ["C4", "=B4*$F$1"],
          ["E2", "outside 2"],
          ["E3", "outside 3"],
          ["F1", 10],
          ["G1", "=B2"],
        ])
      );
      engine.addTable({
        workbookName,
        sheetName,
        tableName,
        start: "A1",
        numRows: { type: "number", value: 3 },
        numCols: 3,
      });

      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 0, direction: "ascending" }],
      });

      expect(column("A", 2, 4)).toEqual(["a", "b", "c"]);
      // The engine stores `A:A` as `A1:A`; the whole column stays whole.
      expect(content("C2")).toBe("=B2*$F$1+COUNTIF(A1:A, A2)");
      expect(content("C3")).toBe("=B3*$F$1");
      expect(content("C4")).toBe("=B4*$F$1");
      expect(column("C", 2, 4)).toEqual([11, 20, 30]);
      // Cells outside the table's columns stay where they are.
      expect(column("E", 2, 3)).toEqual(["outside 2", "outside 3"]);
      // References from outside the table are not rewritten.
      expect(content("G1")).toBe("=B2");
      expect(value("G1")).toBe(1);
    });

    test("turns a reference moved off the sheet into #REF!", () => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "Key"],
          ["B1", "Prev"],
          ["A2", 2],
          ["B2", 0],
          ["A3", 1],
          ["B3", "=A1"],
        ])
      );
      engine.addTable({
        workbookName,
        sheetName,
        tableName,
        start: "A1",
        numRows: { type: "number", value: 2 },
        numCols: 2,
      });
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 0, direction: "ascending" }],
      });
      expect(content("B2")).toBe("=#REF!");
      expect(value("B2")).toBe("#REF!");
    });

    test("moves cell metadata, highlighted styles and data types with their rows", () => {
      setUpOrders();
      engine.setCellMetadata(address("B4"), { note: "banana note" });
      // A uniform column style and one highlighted row.
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { row: 1, col: 2 },
              end: {
                row: { type: "number", value: 6 },
                col: { type: "number", value: 2 },
              },
            },
          },
        ],
        style: { bold: true },
      });
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { row: 3, col: 0 },
              end: {
                row: { type: "number", value: 3 },
                col: { type: "number", value: 3 },
              },
            },
          },
        ],
        style: { backgroundColor: "#ffff00" },
      });
      engine.addCellDataType({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { row: 3, col: 1 },
              end: {
                row: { type: "number", value: 3 },
                col: { type: "number", value: 1 },
              },
            },
          },
        ],
        dataType: "text",
      });

      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 2, direction: "ascending" }],
      });

      // banana (20 units) moved from row 4 to row 3.
      expect(value("B3")).toBe("banana");
      expect(engine.getCellMetadata(address("B3"))).toEqual({
        note: "banana note",
      });
      expect(engine.getCellMetadata(address("B4"))).toBeUndefined();
      expect(engine.getCellStyle(address("A3"))?.backgroundColor).toBe("#ffff00");
      expect(engine.getCellStyle(address("C3"))).toEqual({
        bold: true,
        backgroundColor: "#ffff00",
      });
      expect(engine.getCellStyle(address("A4"))).toBeUndefined();
      expect(engine.getCellStyle(address("C4"))).toEqual({ bold: true });
      expect(engine.getCellDataType(address("B3"))).toBe("text");
      expect(engine.getCellDataType(address("B4"))).toBe("general");
    });

    test("does not rewrite uniform formatting", () => {
      setUpOrders();
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { row: 1, col: 0 },
              end: {
                row: { type: "number", value: 6 },
                col: { type: "number", value: 3 },
              },
            },
          },
        ],
        style: { italic: true },
      });
      const before = engine.getAllCellStyles();
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 2, direction: "descending" }],
      });
      expect(engine.getAllCellStyles()).toEqual(before);
    });

    test("undoes and redoes a sort as one step", () => {
      setUpOrders();
      engine.setCellMetadata(address("A2"), { tag: 1 });
      const original = column("B", 2, 7);
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 2, direction: "descending" }],
      });
      const sorted = column("B", 2, 7);
      expect(sorted).not.toEqual(original);

      expect(engine.undo()).toBe(true);
      expect(column("B", 2, 7)).toEqual(original);
      expect(column("D", 2, 7)).toEqual([300, 100, 200, 500, 400, 600]);
      expect(engine.getCellMetadata(address("A2"))).toEqual({ tag: 1 });
      expect(table().sort).toBeUndefined();

      expect(engine.redo()).toBe(true);
      expect(column("B", 2, 7)).toEqual(sorted);
      expect(table().sort).toEqual([{ column: 2, direction: "descending" }]);
    });

    test("sorts the populated rows of an unbounded table", () => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "N"],
          ["A2", 3],
          ["A3", 1],
          ["A4", 2],
        ])
      );
      engine.addTable({
        workbookName,
        sheetName,
        tableName,
        start: "A1",
        numRows: { type: "infinity", sign: "positive" },
        numCols: 1,
      });
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 0, direction: "ascending" }],
      });
      expect(column("A", 2, 5)).toEqual([1, 2, 3, ""]);
    });

    test("rejects invalid keys", () => {
      setUpOrders();
      expect(() =>
        engine.sortTable({ workbookName, tableName, keys: [] })
      ).toThrow("at least one sort key");
      expect(() =>
        engine.sortTable({
          workbookName,
          tableName,
          keys: [{ column: 4, direction: "ascending" }],
        })
      ).toThrow("out of range");
      expect(() =>
        engine.sortTable({
          workbookName,
          tableName: "Missing",
          keys: [{ column: 0, direction: "ascending" }],
        })
      ).toThrow("Table not found");
      expect(engine.canUndo()).toBe(true);
      expect(column("C", 2, 7)).toEqual([30, 10, 20, 50, 40, 60]);
    });
  });

  describe("filters", () => {
    test("hides the rows a values filter leaves out, text without regard to case", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["north"] });

      expect(hiddenRows()).toEqual([2, 4, 5]);
      expect(table().filter).toEqual({
        columns: [{ column: 0, filter: { type: "values", values: ["north"] } }],
        hiddenRows: [2, 4, 5],
      });
    });

    test("combines filters of several columns", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["North", "South"] });
      filter(2, {
        type: "condition",
        conditions: [{ operator: "greaterThan", value: 25 }],
      });
      // Rows 2 (North 30), 6 (South 40), 7 (North 60) stay.
      expect(hiddenRows()).toEqual([2, 3, 4]);

      filter(2, undefined);
      expect(hiddenRows()).toEqual([4]);
      expect(table().filter?.columns.map((entry) => entry.column)).toEqual([0]);

      engine.clearTableFilters({ workbookName, tableName });
      expect(hiddenRows()).toEqual([]);
      expect(table().filter).toBeUndefined();
    });

    test("supports text conditions, any-of matching, blanks and errors", () => {
      engine.setSheetContent(
        { workbookName, sheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "Name"],
          ["A2", "Alpha"],
          ["A3", "beta"],
          ["A5", "=#N/A"],
          ["A6", "Gamma ray"],
          ["A7", 12],
        ])
      );
      engine.addTable({
        workbookName,
        sheetName,
        tableName,
        start: "A1",
        numRows: { type: "number", value: 6 },
        numCols: 1,
      });

      filter(0, {
        type: "condition",
        conditions: [
          { operator: "beginsWith", value: "AL" },
          { operator: "contains", value: "RAY" },
        ],
        match: "any",
      });
      expect(hiddenRows()).toEqual([2, 3, 4, 6]);

      filter(0, {
        type: "condition",
        conditions: [{ operator: "notEquals", value: "beta" }],
      });
      // Blank, error and number rows differ from "beta" too.
      expect(hiddenRows()).toEqual([2]);

      filter(0, {
        type: "values",
        values: [12],
        blanks: true,
        errors: ["#N/A"],
      });
      expect(hiddenRows()).toEqual([1, 2, 5]);
    });

    test("keeps edited rows visible until the filters are re-applied", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["North"] });
      expect(hiddenRows()).toEqual([2, 4, 5]);

      engine.setCellContent(address("A2"), "West");
      expect(hiddenRows()).toEqual([2, 4, 5]);

      engine.reapplyTableFilters({ workbookName, tableName });
      expect(hiddenRows()).toEqual([1, 2, 4, 5]);
    });

    test("re-applies filters after a sort so the same rows stay visible", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["North"] });
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 2, direction: "ascending" }],
      });
      // Units ascending: 10 south, 20 North, 30 North, 40 South, 50 East, 60 North.
      expect(hiddenRows()).toEqual([1, 4, 5]);
      expect(table().sort).toEqual([{ column: 2, direction: "ascending" }]);
    });

    test("remaps filters by column name when the table is resized", () => {
      setUpOrders();
      engine.setCellContent(address("A8"), "North");
      engine.setCellContent(address("B8"), "fig");
      engine.setCellContent(address("C8"), 70);
      filter(1, { type: "values", values: ["apple", "fig"] });
      expect(hiddenRows()).toEqual([1, 3, 5, 6]);

      engine.updateTable({
        workbookName,
        tableName,
        start: "B1",
        numRows: { type: "number", value: 7 },
        numCols: 3,
      });
      // Product is column 0 now; the new row 8 is filtered too.
      expect(table().filter?.columns).toEqual([
        { column: 0, filter: { type: "values", values: ["apple", "fig"] } },
      ]);
      expect(hiddenRows()).toEqual([1, 3, 5, 6]);

      engine.updateTable({
        workbookName,
        tableName,
        start: "C1",
        numRows: { type: "number", value: 7 },
        numCols: 2,
      });
      expect(table().filter).toBeUndefined();
      expect(hiddenRows()).toEqual([]);
    });

    test("moves hidden rows with a moved table", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["East"] });
      expect(hiddenRows()).toEqual([1, 2, 3, 5, 6]);
      engine.moveRange(
        {
          workbookName,
          sheetName,
          range: {
            start: { row: 0, col: 0 },
            end: {
              row: { type: "number", value: 6 },
              col: { type: "number", value: 3 },
            },
          },
        },
        { workbookName, sheetName, rowIndex: 10, colIndex: 0 }
      );
      expect(hiddenRows()).toEqual([11, 12, 13, 15, 16]);
      expect(value("A15")).toBe("East");
    });

    test("keeps filters when a header is renamed", () => {
      setUpOrders();
      filter(2, {
        type: "condition",
        conditions: [{ operator: "lessThan", value: 30 }],
      });
      engine.setCellContent(address("C1"), "Quantity");
      expect(table().filter?.columns[0]?.column).toBe(2);
      expect(hiddenRows()).toEqual([1, 4, 5, 6]);
    });

    test("undoes and redoes filter changes", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["North"] });
      filter(2, {
        type: "condition",
        conditions: [{ operator: "greaterThanOrEqual", value: 60 }],
      });
      expect(hiddenRows()).toEqual([1, 2, 3, 4, 5]);

      engine.undo();
      expect(hiddenRows()).toEqual([2, 4, 5]);
      engine.undo();
      expect(hiddenRows()).toEqual([]);
      expect(table().filter).toBeUndefined();
      engine.redo();
      expect(hiddenRows()).toEqual([2, 4, 5]);
    });

    test("validates filters", () => {
      setUpOrders();
      expect(() => filter(9, { type: "values", values: [] })).toThrow(
        "out of range"
      );
      expect(() =>
        filter(0, { type: "condition", conditions: [] })
      ).toThrow("at least one condition");
      expect(() =>
        filter(0, {
          type: "condition",
          conditions: [{ operator: "like" as never, value: 1 }],
        })
      ).toThrow("Invalid filter condition");
      expect(() =>
        filter(0, { type: "values", values: [{}] as never })
      ).toThrow("values filter");
    });
  });

  describe("getTableColumnValues", () => {
    test("lists distinct values in sort order with counts", () => {
      setUpOrders();
      engine.setCellContent(address("A3"), undefined);
      expect(
        engine.getTableColumnValues({ workbookName, tableName, column: 0 })
      ).toEqual([
        { kind: "value", value: "East", count: 1, selected: true },
        { kind: "value", value: "North", count: 3, selected: true },
        { kind: "value", value: "South", count: 1, selected: true },
        { kind: "blank", count: 1, selected: true },
      ]);
    });

    test("narrows to rows the other filters show and marks the own selection", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["North"] });
      filter(1, { type: "values", values: ["pear"] });
      expect(
        engine.getTableColumnValues({ workbookName, tableName, column: 1 })
      ).toEqual([
        { kind: "value", value: "banana", count: 1, selected: false },
        { kind: "value", value: "date", count: 1, selected: false },
        { kind: "value", value: "pear", count: 1, selected: true },
      ]);
      expect(
        engine
          .getTableColumnValues({ workbookName, tableName, column: 0 })
          .map((entry) => [
            entry.kind === "value" ? entry.value : entry.kind,
            entry.selected,
          ])
      ).toEqual([["North", true]]);
    });
  });

  describe("persistence", () => {
    test("round-trips filters and sort keys through serialization", () => {
      setUpOrders();
      engine.sortTable({
        workbookName,
        tableName,
        keys: [{ column: 2, direction: "descending" }],
      });
      filter(0, { type: "values", values: ["North"] });
      const expected = table();

      for (const includeEvaluationState of [true, false]) {
        const restored = FormulaEngine.buildEmpty();
        restored.resetToSerializedEngine(
          engine.serializeEngine({ includeEvaluationState })
        );
        const restoredTable = restored.getTable({ workbookName, tableName })!;
        expect({
          filter: restoredTable.filter,
          sort: restoredTable.sort,
        }).toEqual({ filter: expected.filter, sort: expected.sort });
        expect([
          ...restored.getHiddenRows({ workbookName, sheetName }),
        ]).toEqual(expected.filter!.hiddenRows);
      }
    });

    test("applies filters and sort keys from WorkbookData", () => {
      engine.addWorkbook({
        workbookName: "Imported",
        data: {
          sheets: [
            {
              name: "Data",
              content: {
                A1: "Name",
                B1: "Score",
                A2: "a",
                B2: "=10*2",
                A3: "b",
                B3: 5,
                A4: "c",
                B4: 30,
              },
            },
          ],
          tables: [
            {
              name: "Scores",
              sheetName: "Data",
              start: "A1",
              numRows: { type: "number", value: 3 },
              numCols: 2,
              filters: [
                {
                  column: 1,
                  filter: {
                    type: "condition",
                    conditions: [{ operator: "greaterThan", value: 10 }],
                  },
                },
              ],
              sort: [{ column: 0, direction: "ascending" }],
            },
          ],
        },
      });
      expect([
        ...engine.getHiddenRows({ workbookName: "Imported", sheetName: "Data" }),
      ]).toEqual([2]);
      expect(
        engine.getTable({ workbookName: "Imported", tableName: "Scores" })?.sort
      ).toEqual([{ column: 0, direction: "ascending" }]);
    });

    test("copies filters with a cloned sheet", () => {
      setUpOrders();
      filter(0, { type: "values", values: ["South"] });
      engine.cloneSheet({ workbookName, sheetName, newSheetName: "Copy" });
      expect([
        ...engine.getHiddenRows({ workbookName, sheetName: "Copy" }),
      ].sort((a, b) => a - b)).toEqual(hiddenRows());
    });
  });
});
