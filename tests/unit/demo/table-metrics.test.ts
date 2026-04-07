import { describe, expect, test } from "bun:test";
import type { TableDefinition } from "../../../src/core/types";
import {
  getFiniteTableLastRowIndex,
  getFiniteTableRowCount,
  isTableLastRow,
} from "../../../demo/lib/table-metrics";

function makeTable({
  startRowIndex,
  endRowValue,
}: {
  startRowIndex: number;
  endRowValue?: number;
}): TableDefinition {
  return {
    name: "Table1",
    sheetName: "Sheet1",
    workbookName: "Workbook1",
    start: {
      rowIndex: startRowIndex,
      colIndex: 0,
    },
    headers: new Map([
      ["Name", { name: "Name", index: 0 }],
      ["Value", { name: "Value", index: 1 }],
    ]),
    endRow:
      endRowValue === undefined
        ? { type: "infinity", sign: "positive" }
        : { type: "number", value: endRowValue },
  };
}

describe("demo table metrics", () => {
  test("uses the absolute table end row for tables starting on row 1", () => {
    const table = makeTable({ startRowIndex: 0, endRowValue: 4 });

    expect(getFiniteTableLastRowIndex(table)).toBe(4);
    expect(getFiniteTableRowCount(table)).toBe(5);
    expect(isTableLastRow(table, 4)).toBe(true);
    expect(isTableLastRow(table, 3)).toBe(false);
  });

  test("uses the same absolute end row for tables starting below row 1", () => {
    const table = makeTable({ startRowIndex: 6, endRowValue: 10 });

    expect(getFiniteTableLastRowIndex(table)).toBe(10);
    expect(getFiniteTableRowCount(table)).toBe(5);
    expect(isTableLastRow(table, 10)).toBe(true);
    expect(isTableLastRow(table, 9)).toBe(false);
  });

  test("does not mark a last row for infinite tables", () => {
    const table = makeTable({ startRowIndex: 3 });

    expect(getFiniteTableLastRowIndex(table)).toBeUndefined();
    expect(getFiniteTableRowCount(table)).toBeUndefined();
    expect(isTableLastRow(table, 100)).toBe(false);
  });
});
