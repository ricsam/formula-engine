import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import {
  buildRangeEvalOrder,
  type RangeEvalOrderEntry,
  type RangeEvalOrderEntryDict,
} from "../../../src/core/managers/range-eval-order-builder";
import { WorkbookManager } from "../../../src/core/managers/workbook-manager";
import type {
  CellAddress,
  RangeAddress,
  SerializedCellValue,
} from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";
import { visualizeSpreadsheet } from "../../../src/core/utils/spreadsheet-visualizer";

function assertType<T extends keyof RangeEvalOrderEntryDict>(
  value: RangeEvalOrderEntry | undefined,
  type: T
): asserts value is RangeEvalOrderEntryDict[T] {
  if (value?.type !== type) {
    throw new Error(`Expected ${type}, got ${value}`);
  }
}

describe("buildRangeEvalOrder", () => {
  let manager: WorkbookManager;
  const workbookName = "TestWorkbook";
  const sheetName = "Sheet1";

  let engine: FormulaEngine;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
    manager = engine._workbookManager;
  });

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  function setCell(col: string, row: number, value: string): void {
    const colIndex = col.charCodeAt(0) - "A".charCodeAt(0);
    manager.setCellContent(
      {
        workbookName,
        sheetName,
        rowIndex: row - 1,
        colIndex,
      },
      value
    );
  }

  function makeRange(startCell: string, endCell: string): RangeAddress {
    const parseCell = (cell: string) => {
      const col = cell.charCodeAt(0) - "A".charCodeAt(0);
      const row = parseInt(cell.slice(1)) - 1;
      return { col, row };
    };

    const start = parseCell(startCell);
    const end = parseCell(endCell);

    return {
      workbookName,
      sheetName,
      range: {
        start: { row: start.row, col: start.col },
        end: {
          row: { type: "number", value: end.row },
          col: { type: "number", value: end.col },
        },
      },
    };
  }

  function makeCell(cell: string): CellAddress {
    const col = cell.charCodeAt(0) - "A".charCodeAt(0);
    const row = parseInt(cell.slice(1)) - 1;
    return {
      workbookName,
      sheetName,
      rowIndex: row,
      colIndex: col,
    };
  }

  function cellToString(addr: CellAddress): string {
    const col = String.fromCharCode(65 + addr.colIndex);
    const row = addr.rowIndex + 1;
    return `${col}${row}`;
  }

  function rangeToString(addr: RangeAddress): string {
    const startCol = String.fromCharCode(65 + addr.range.start.col);
    const startRow = addr.range.start.row + 1;
    const endCol =
      addr.range.end.col.type === "number"
        ? String.fromCharCode(65 + addr.range.end.col.value)
        : "∞";
    const endRow =
      addr.range.end.row.type === "number" ? addr.range.end.row.value + 1 : "∞";
    return `${startCol}${startRow}:${endCol}${endRow}`;
  }

  /**
   * Serialize a RangeEvalOrderEntry to a readable string format for snapshot testing
   * Examples:
   * - "value:A1" - occupied cell
   * - "empty:C4[]" - empty cell with no candidates
   * - "empty:C4[A2,B1]" - empty cell with left and above candidates
   * - "range:C5:D5[A2]" - empty range with candidate
   */
  function serializeEntry(entry: RangeEvalOrderEntry): string {
    if (entry.type === "value") {
      return `value:${cellToString(entry.address)}`;
    } else if (entry.type === "empty_cell") {
      const candidates = entry.candidates.map(cellToString).join(",");
      return `empty:${cellToString(entry.address)}[${candidates}]`;
    } else {
      // empty_range
      const candidates = entry.candidates.map(cellToString).join(",");
      return `range:${rangeToString(entry.address)}[${candidates}]`;
    }
  }

  test("empty range with no occupied cells - row-major", () => {
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    // Should have one entry for the entire empty range per row
    expect(result.length).toBe(3);

    // Row 1: A1:C1
    expect(result[0]!.type).toBe("empty_range");
    expect(rangeToString((result[0] as any).address)).toBe("A1:C1");
    expect((result[0] as any).candidates).toEqual([]);

    // Row 2: A2:C2
    expect(result[1]!.type).toBe("empty_range");
    expect(rangeToString((result[1] as any).address)).toBe("A2:C2");
    expect((result[1] as any).candidates).toEqual([]);

    // Row 3: A3:C3
    expect(result[2]!.type).toBe("empty_range");
    expect(rangeToString((result[2] as any).address)).toBe("A3:C3");
    expect((result[2] as any).candidates).toEqual([]);
  });

  test("empty range with no occupied cells - col-major", () => {
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    // Should have one entry for the entire empty range per column
    expect(result.length).toBe(3);

    // Col A: A1:A3
    expect(result[0]!.type).toBe("empty_range");
    expect(rangeToString((result[0] as any).address)).toBe("A1:A3");
    expect((result[0] as any).candidates).toEqual([]);

    // Col B: B1:B3
    expect(result[1]!.type).toBe("empty_range");
    expect(rangeToString((result[1] as any).address)).toBe("B1:B3");
    expect((result[1] as any).candidates).toEqual([]);

    // Col C: C1:C3
    expect(result[2]!.type).toBe("empty_range");
    expect(rangeToString((result[2] as any).address)).toBe("C1:C3");
    expect((result[2] as any).candidates).toEqual([]);
  });

  test("single occupied cell - row-major", () => {
    setCell("B", 2, "=10");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 |     |     |    
         2 |     | 10  |    
         3 |     |     |    
      "
    `);

    // Snapshot test the entire result
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:A1:C1[]",
        "empty:A2[]",
        "value:B2",
        "empty:C2[B2]",
        "range:A3:C3[B2]",
      ]
    `);
  });

  test("single occupied cell - col-major", () => {
    setCell("B", 2, "=10");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 |     |     |    
         2 |     | 10  |    
         3 |     |     |    
      "
    `);

    // Snapshot test the entire result
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:A1:A3[]",
        "empty:B1[]",
        "value:B2",
        "empty:B3[B2]",
        "range:C1:C3[B2]",
      ]
    `);
  });

  test("two candidates - row-major ordering", () => {
    setCell("A", 2, "=1");
    setCell("B", 1, "=2");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 |     | 2   |    
         2 | 1   |     |    
         3 |     |     |    
      "
    `);

    // B2 should have both left (A2) and above (B1) candidates
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:A1[]",
        "value:B1",
        "empty:C1[B1]",
        "value:A2",
        "range:B2:C2[A2,B1]",
        "range:A3:C3[A2,B1]",
      ]
    `);
  });

  test("two candidates - col-major ordering", () => {
    setCell("A", 2, "=1");
    setCell("B", 1, "=2");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 |     | 2   |    
         2 | 1   |     |    
         3 |     |     |    
      "
    `);

    // B2 should have both above (B1) and left (A2) candidates
    // In col-major, above comes first
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:A1[]",
        "value:A2",
        "empty:A3[A2]",
        "value:B1",
        "range:B2:B3[B1,A2]",
        "range:C1:C3[B1,A2]",
      ]
    `);
  });

  test("complex grid with multiple values - row-major", () => {
    // Set up a grid like:
    //   A  B  C  D
    // 1    =2
    // 2 =1    =3
    // 3
    setCell("B", 1, "=2");
    setCell("A", 2, "=1");
    setCell("C", 2, "=3");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 4,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C   | D  
      -----+-----+-----+-----+----
         1 |     | 2   |     |    
         2 | 1   |     | 3   |    
         3 |     |     |     |    
      "
    `);

    const range = makeRange("A1", "D3");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    // Snapshot test the entire result
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:A1[]",
        "value:B1",
        "range:C1:D1[B1]",
        "value:A2",
        "empty:B2[A2,B1]",
        "value:C2",
        "empty:D2[C2]",
        "range:A3:D3[A2,B1,C2]",
      ]
    `);
  });

  test("complex grid with multiple values - col-major", () => {
    // Set up a grid like:
    //   A  B  C  D
    // 1    =2
    // 2 =1    =3
    // 3
    setCell("B", 1, "=2");
    setCell("A", 2, "=1");
    setCell("C", 2, "=3");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 4,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C   | D  
      -----+-----+-----+-----+----
         1 |     | 2   |     |    
         2 | 1   |     | 3   |    
         3 |     |     |     |    
      "
    `);

    const range = makeRange("A1", "D3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:A1[]",
        "value:A2",
        "empty:A3[A2]",
        "value:B1",
        "range:B2:B3[B1,A2]",
        "empty:C1[B1]",
        "value:C2",
        "empty:C3[C2]",
        "range:D1:D3[B1,C2]",
      ]
    `);
  });

  test("fully occupied range", () => {
    setCell("A", 1, "=1");
    setCell("B", 1, "=2");
    setCell("A", 2, "=3");
    setCell("B", 2, "=4");

    const range = makeRange("A1", "B2");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 2,
        numCols: 2,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B  
      -----+-----+----
         1 | 1   | 2  
         2 | 3   | 4  
      "
    `);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:A1",
        "value:B1",
        "value:A2",
        "value:B2",
      ]
    `);
  });

  test("single cell range - occupied", () => {
    setCell("B", 2, "=42");
    const range = makeRange("B2", "B2");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:B2",
      ]
    `);
  });

  test("single cell range - empty", () => {
    const range = makeRange("B2", "B2");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:B2[]",
      ]
    `);
  });

  test("single cell range - empty with left candidate", () => {
    setCell("A", 2, "=1");
    const range = makeRange("B2", "B2");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:B2[A2]",
      ]
    `);
  });

  test("diagonal candidates can be found outside lookup range", () => {
    // Place a formula outside the lookup range (top-left diagonal)
    setCell("A", 1, "=999"); // Outside range, but diagonal to B2

    // Lookup range is B2:D4
    const range = makeRange("B2", "D4");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    // A1 is found as a diagonal candidate for all cells in the range
    // since there are no direct left/above formula candidates within the range
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:B2:D2[A1]",
        "range:B3:D3[A1]",
        "range:B4:D4[A1]",
      ]
    `);
  });

  test("empty cells between two occupied cells in a row", () => {
    setCell("A", 1, "=1");
    setCell("D", 1, "=2");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 1,
        numCols: 4,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C   | D  
      -----+-----+-----+-----+----
         1 | 1   |     |     | 2  
      "
    `);

    const range = makeRange("A1", "D1");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:A1",
        "range:B1:C1[A1]",
        "value:D1",
      ]
    `);
  });

  test("empty cells between two occupied cells in a column", () => {
    setCell("A", 1, "=1");
    setCell("A", 4, "=2");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 4,
        numCols: 1,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A  
      -----+----
         1 | 1  
         2 |    
         3 |    
         4 | 2  
      "
    `);

    const range = makeRange("A1", "A4");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    // A1 is occupied
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:A1",
        "range:A2:A3[A1]",
        "value:A4",
      ]
    `);
  });

  test("row-major preserves row-by-row order", () => {
    setCell("B", 2, "=1");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 |     |     |    
         2 |     | 1   |    
         3 |     |     |    
      "
    `);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:A1:C1[]",
        "empty:A2[]",
        "value:B2",
        "empty:C2[B2]",
        "range:A3:C3[B2]",
      ]
    `);
  });

  test("col-major preserves col-by-col order", () => {
    setCell("B", 2, "=1");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 |     |     |    
         2 |     | 1   |    
         3 |     |     |    
      "
    `);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:A1:A3[]",
        "empty:B1[]",
        "value:B2",
        "empty:B3[B2]",
        "range:C1:C3[B2]",
      ]
    `);
  });

  test("complex spilling scenario with multiple formulas - row-major", () => {
    // Set up a complex grid with multiple spilling formulas
    // G1 spills down to G7
    setCell("G", 1, "=SEQUENCE(7,1)");
    // C2 spills to C2:E4
    setCell("C", 2, "=SEQUENCE(3,3)");
    // C5 would spill but gets blocked (outside lookup range anyway)
    setCell("C", 5, "=SEQUENCE(1,4)");
    // E5 spills to E5:F5
    setCell("E", 5, "=SEQUENCE(1,2)");
    // F4 and F6 are regular values
    setCell("F", 4, "42");
    setCell("F", 6, "99");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 7,
        numCols: 7,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
        startCol: 2,
      })
    ).toMatchInlineSnapshot(`
      "    | C       | D   | E   | F   | G   | H   | I  
      -----+---------+-----+-----+-----+-----+-----+----
         1 |         |     |     |     | 1   |     |    
         2 | 1       | 2   | 3   |     | 2   |     |    
         3 | 4       | 5   | 6   |     | 3   |     |    
         4 | 7       | 8   | 9   | 42  | 4   |     |    
         5 | #SPILL! |     | 1   | 2   | 5   |     |    
         6 |         |     |     | 99  | 6   |     |    
         7 |         |     |     |     | 7   |     |    
      "
    `);

    // Lookup range is E4:G6
    const range = makeRange("E4", "G6");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:E4[C2]",
        "value:F4",
        "empty:G4[G1]",
        "value:E5",
        "range:F5:G5[E5,G1]",
        "empty:E6[E5]",
        "value:F6",
        "empty:G6[G1]",
      ]
    `);
  });

  test("infinite range with open-ended columns - row-major", () => {
    // Set up a simple grid with formulas
    setCell("A", 1, "=1");
    setCell("C", 1, "=2");
    setCell("A", 2, "=3");

    // Lookup range is A1:∞1 (entire first row)
    const range: RangeAddress = {
      workbookName,
      sheetName,
      range: {
        start: { row: 0, col: 0 }, // A1
        end: {
          row: { type: "number", value: 0 }, // Row 1
          col: { type: "infinity", sign: "positive" }, // All columns
        },
      },
    };

    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    // Should only include cells up to the maximum occupied cell (C1)
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:A1",
        "empty:B1[A1]",
        "value:C1",
        "range:D1:∞1[C1]",
      ]
    `);
  });

  test("infinite range with open-ended rows - col-major", () => {
    // Set up a simple grid with formulas
    setCell("A", 1, "=1");
    setCell("A", 3, "=2");
    setCell("B", 1, "=3");

    // Lookup range is A1:A∞ (entire first column)
    const range: RangeAddress = {
      workbookName,
      sheetName,
      range: {
        start: { row: 0, col: 0 }, // A1
        end: {
          row: { type: "infinity", sign: "positive" }, // All rows
          col: { type: "number", value: 0 }, // Column A
        },
      },
    };

    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    // Should include cells up to A3 and then an infinite range A4:A∞
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:A1",
        "empty:A2[A1]",
        "value:A3",
        "range:A4:A∞[A3]",
      ]
    `);
  });

  test("infinite sorts candidates - col-major", () => {
    // Set up formulas at different rows in columns A and B
    // B6 at row 5, col 1
    setCell("B", 6, '="Formula"');
    // A11 at row 10, col 0
    setCell("A", 11, '="Formula"');

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 15,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A       | B       | C  
      -----+---------+---------+----
         1 |         |         |    
         2 |         |         |    
         3 |         |         |    
         4 |         |         |    
         5 |         |         |    
         6 |         | Formula |    
         7 |         |         |    
         8 |         |         |    
         9 |         |         |    
        10 |         |         |    
        11 | Formula |         |    
        12 |         |         |    
        13 |         |         |    
        14 |         |         |    
        15 |         |         |    
      "
    `);

    // Lookup range is C:C (entire column C)
    const range: RangeAddress = {
      workbookName,
      sheetName,
      range: {
        start: { row: 0, col: 2 }, // C1
        end: {
          row: { type: "infinity", sign: "positive" }, // All rows
          col: { type: "number", value: 2 }, // Column C
        },
      },
    };

    const result = buildRangeEvalOrder.call(manager, "col-major", range);
    // with col-major, we sort candidates by row first, then col
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:C1:C∞[B6,A11]",
      ]
    `);
  });

  test("debug eval order for infinite ranges", () => {
    setCell("A", 1, "=SUM(B10:INFINITY)");
    setCell("A", 10, "=SEQUENCE(3,3)");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 15,
        numCols: 3,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C  
      -----+-----+-----+----
         1 | 33  |     |    
         2 |     |     |    
         3 |     |     |    
         4 |     |     |    
         5 |     |     |    
         6 |     |     |    
         7 |     |     |    
         8 |     |     |    
         9 |     |     |    
        10 | 1   | 2   | 3  
        11 | 4   | 5   | 6  
        12 | 7   | 8   | 9  
        13 |     |     |    
        14 |     |     |    
        15 |     |     |    
      "
    `);

    // B10:∞∞
    const result = buildRangeEvalOrder.call(manager, "col-major", {
      workbookName,
      sheetName,
      range: {
        end: {
          col: {
            type: "infinity",
            sign: "positive",
          },
          row: {
            type: "infinity",
            sign: "positive",
          },
        },
        start: {
          col: 1,
          row: 9,
        },
      },
    });

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:B10:∞∞[A10]",
      ]
    `);
  });

  test("debug eval order for infinite ranges /2", () => {
    setCell("A", 1, "=SUM(B1:D)");
    setCell("A", 2, "=SEQUENCE(1,INFINITY)");

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 3,
        numCols: 5,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A   | B   | C   | D   | E  
      -----+-----+-----+-----+-----+----
         1 | 9   |     |     |     |    
         2 | 1   | 2   | 3   | 4   | 5  
         3 |     |     |     |     |    
      "
    `);

    // B1:D
    const result = buildRangeEvalOrder.call(manager, "col-major", {
      workbookName,
      sheetName,
      range: {
        end: {
          col: {
            type: "number",
            value: 3,
          },
          row: {
            type: "infinity",
            sign: "positive",
          },
        },
        start: {
          col: 1,
          row: 0,
        },
      },
    });

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:B1:D∞[A1,A2]",
      ]
    `);
  });

  test("debug eval order for infinite ranges /3", () => {
    setCell("A", 1, "=SUM(B10:D)");
    setCell("C", 8, "=SEQUENCE(3,1)");
    setCell("A", 10, "=SEQUENCE(1,3)");

    // should spill from C8
    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 10,
        numCols: 5,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A       | B   | C   | D   | E  
      -----+---------+-----+-----+-----+----
         1 | 3       |     |     |     |    
         2 |         |     |     |     |    
         3 |         |     |     |     |    
         4 |         |     |     |     |    
         5 |         |     |     |     |    
         6 |         |     |     |     |    
         7 |         |     |     |     |    
         8 |         |     | 1   |     |    
         9 |         |     | 2   |     |    
        10 | #SPILL! |     | 3   |     |    
      "
    `);

    // B10:D
    const result = buildRangeEvalOrder.call(manager, "col-major", {
      workbookName,
      sheetName,
      range: {
        end: {
          col: {
            type: "number",
            value: 3,
          },
          row: {
            type: "infinity",
            sign: "positive",
          },
        },
        start: {
          col: 1,
          row: 9,
        },
      },
    });

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:B10:D∞[C8,A10]",
      ]
    `);
  });

  test("debug eval order for infinite ranges /4", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["X1", "=SEQUENCE(2,2)"],
        ["A10", "=X1:Y2*10"],
        ["A1", "=SUM(A10:B)"],
        ["A2", "shouldn't be evaluated"],
      ])
    );

    expect(
      visualizeSpreadsheet(engine, {
        workbookName,
        numRows: 12,
        numCols: 2,
        sheetName,
        emptyCellChar: "",
        snapshot: true,
      })
    ).toMatchInlineSnapshot(`
      "    | A                    | B  
      -----+----------------------+----
         1 | 100                  |    
         2 | shouldn't be eval... |    
         3 |                      |    
         4 |                      |    
         5 |                      |    
         6 |                      |    
         7 |                      |    
         8 |                      |    
         9 |                      |    
        10 | 10                   | 20 
        11 | 30                   | 40 
        12 |                      |    
      "
    `);

    // A10:B
    const result = buildRangeEvalOrder.call(manager, "col-major", {
      workbookName,
      sheetName,
      range: {
        end: {
          col: {
            type: "number",
            value: 1,
          },
          row: {
            type: "infinity",
            sign: "positive",
          },
        },
        start: {
          col: 0,
          row: 9,
        },
      },
    });

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "value:A10",
        "range:A11:A∞[A10]",
        "range:B10:B∞[A10]",
      ]
    `);

    // X1:Y2 is [1,2;3,4], multiplied by 10 = [10,20;30,40]
    // So A10:B11 = 10+20+30+40 = 100
    expect(cell("A1", true)).toBe(100);
  });

  test("performance: large sparse range should not loop excessively", () => {
    // Set up a few formulas scattered across a large range
    setCell("A", 1, "=1");
    setCell("Z", 1000, "=2");

    // Lookup range G1:Z1000 - large range but sparse (only 2 cells, neither in range)
    const range: RangeAddress = {
      workbookName,
      sheetName,
      range: {
        start: { row: 0, col: 6 }, // G1
        end: {
          row: { type: "number", value: 999 }, // Row 1000
          col: { type: "number", value: 25 }, // Column Z
        },
      },
    };

    const startTime = performance.now();
    const result = buildRangeEvalOrder.call(manager, "col-major", range);
    const duration = performance.now() - startTime;

    // Performance regression check: Should complete in < 5ms
    // (Was ~27ms before optimization, now ~0.1ms)
    expect(duration).toBeLessThan(5);

    // Check the result structure
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:G1:G1000[A1]",
        "range:H1:H1000[A1]",
        "range:I1:I1000[A1]",
        "range:J1:J1000[A1]",
        "range:K1:K1000[A1]",
        "range:L1:L1000[A1]",
        "range:M1:M1000[A1]",
        "range:N1:N1000[A1]",
        "range:O1:O1000[A1]",
        "range:P1:P1000[A1]",
        "range:Q1:Q1000[A1]",
        "range:R1:R1000[A1]",
        "range:S1:S1000[A1]",
        "range:T1:T1000[A1]",
        "range:U1:U1000[A1]",
        "range:V1:V1000[A1]",
        "range:W1:W1000[A1]",
        "range:X1:X1000[A1]",
        "range:Y1:Y1000[A1]",
        "range:Z1:Z999[A1]",
        "value:Z1000",
      ]
    `);
  });

  test("performance: infinite range should not loop excessively", () => {
    // Set up a few formulas at large indices
    setCell("A", 1, "=1");
    setCell("Z", 1000, "=2");

    // Lookup range B1001:∞∞
    const range: RangeAddress = {
      workbookName,
      sheetName,
      range: {
        start: { row: 1000, col: 1 }, // B1001
        end: {
          row: { type: "infinity", sign: "positive" },
          col: { type: "infinity", sign: "positive" },
        },
      },
    };

    const startTime = performance.now();
    const result = buildRangeEvalOrder.call(manager, "col-major", range);
    const duration = performance.now() - startTime;

    // Should complete very quickly (< 50ms)
    expect(duration).toBeLessThan(50);

    // Should just have one infinite range with candidates
    expect(result.length).toBe(1);
    expect(result[0]?.type).toBe("empty_range");
  });

  test("performance: diagonal candidates with many formulas should use O(n log n) algorithm", () => {
    // Create a grid of formulas in the top-left quadrant
    // This simulates the worst case where we have many diagonal candidates
    // With the old O(n²) algorithm, this would take ~500ms for 25k candidates
    // With O(n log n), should complete in < 10ms

    const gridSize = 100; // 100x100 = 10,000 formulas
    for (let row = 1; row <= gridSize; row++) {
      for (let col = 0; col < gridSize; col++) {
        const colLetter = String.fromCharCode(65 + (col % 26)); // A-Z repeating
        setCell(colLetter, row, `=${row * 100 + col}`);
      }
    }

    // Target cell far to the right and below
    const targetCell: CellAddress = {
      workbookName,
      sheetName,
      rowIndex: 200,
      colIndex: 200,
    };

    const startTime = performance.now();

    // This will trigger findAllDiagonalStepCandidates with ~10k candidates
    const range: RangeAddress = {
      workbookName,
      sheetName,
      range: {
        start: { row: targetCell.rowIndex, col: targetCell.colIndex },
        end: {
          row: { type: "number", value: targetCell.rowIndex },
          col: { type: "number", value: targetCell.colIndex },
        },
      },
    };

    buildRangeEvalOrder.call(manager, "col-major", range);

    const duration = performance.now() - startTime;

    // Performance regression check: Should complete in < 50ms
    // (Would be ~500ms with O(n²) algorithm for 25k candidates)
    // With O(n log n), typically ~10ms for 10k candidates
    expect(duration).toBeLessThan(50);
  });
});
