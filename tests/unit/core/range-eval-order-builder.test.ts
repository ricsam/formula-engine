import { describe, expect, test, beforeEach } from "bun:test";
import { WorkbookManager } from "../../../src/core/managers/workbook-manager";
import {
  buildRangeEvalOrder,
  type RangeEvalOrderEntry,
  type LookupOrder,
  type RangeEvalOrderEntryDict,
} from "../../../src/core/managers/range-eval-order-builder";
import type {
  RangeAddress,
  CellAddress,
  SerializedCellValue,
} from "../../../src/core/types";
import { getCellReference, parseCellReference } from "../../../src/core/utils";
import { visualizeSpreadsheet } from "src/core/utils/spreadsheet-visualizer";
import { FormulaEngine } from "src/core/engine";

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

    // Snapshot test the entire result
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:A1:C1[]",
        "empty:A2[]",
        "value:B2",
        "empty:C2[B2]",
        "empty:A3[]",
        "range:B3:C3[B2]",
      ]
    `);
  });

  test("single occupied cell - col-major", () => {
    setCell("B", 2, "=10");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    // Snapshot test the entire result
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:A1:A3[]",
        "empty:B1[]",
        "value:B2",
        "empty:B3[B2]",
        "empty:C1[]",
        "range:C2:C3[B2]",
      ]
    `);
  });

  test("two candidates - row-major ordering", () => {
    setCell("A", 2, "=1");
    setCell("B", 1, "=2");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "row-major", range);

    // Find B2 which should have both left (A2) and above (B1) candidates
    const b2Entry = result.find(
      (e) => e.type === "empty_cell" && cellToString(e.address) === "B2"
    );
    expect(b2Entry).toBeDefined();
    if (b2Entry && b2Entry.type === "empty_cell") {
      expect(b2Entry.candidates.map(cellToString)).toEqual(["A2", "B1"]);
    }
  });

  test("two candidates - col-major ordering", () => {
    setCell("A", 2, "=1");
    setCell("B", 1, "=2");
    const range = makeRange("A1", "C3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    // Find B2 which should have both above (B1) and left (A2) candidates
    const b2Entry = result.find(
      (e) => e.type === "empty_cell" && cellToString(e.address) === "B2"
    );
    expect(b2Entry).toBeDefined();
    // In col-major, above comes first
    if (b2Entry && b2Entry.type === "empty_cell") {
      expect(b2Entry.candidates.map(cellToString)).toEqual(["B1", "A2"]);
    }
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
        "empty:A3[A2]",
        "empty:B3[B1]",
        "range:C3:D3[C2]",
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

    const range = makeRange("A1", "D3");
    const result = buildRangeEvalOrder.call(manager, "col-major", range);

    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "empty:A1[]",
        "value:A2",
        "empty:A3[A2]",
        "value:B1",
        "empty:B2[B1,A2]",
        "empty:B3[B1]",
        "empty:C1[B1]",
        "value:C2",
        "empty:C3[C2]",
        "empty:D1[B1]",
        "range:D2:D3[C2]",
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
        "empty:A3[]",
        "range:B3:C3[B2]",
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
        "empty:C1[]",
        "range:C2:C3[B2]",
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
        "empty:F5[E5]",
        "empty:G5[E5,G1]",
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

  test("infinite range splits on candidate changes to avoid ambiguity - col-major", () => {
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

    // The range should be split based on when candidates change:
    // - C1:C5 - no diagonal candidates (before B6)
    // - C6:C10 - B6 is a candidate (B6 is at row 5, so cells from row 5 onwards could be influenced)
    // - C11:C13 - Both B6 and A11 are candidates (A11 is at row 10, so from row 10 onwards)
    expect(result.map(serializeEntry)).toMatchInlineSnapshot(`
      [
        "range:C1:C5[]",
        "range:C6:C10[B6]",
        "empty:C11[A11]",
        "range:C12:C∞[A11,B6]",
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
        "empty:B1[A1]",
        "empty:B2[A2]",
        "range:B3:B∞[A2]",
        "empty:C1[A1]",
        "empty:C2[A2]",
        "range:C3:C∞[A2]",
        "empty:D1[A1]",
        "empty:D2[A2]",
        "range:D3:D∞[A2]",
      ]
    `);
  });

  test("debug eval order for infinite ranges /3", () => {
    setCell("A", 1, "=SUM(B10:D)");
    setCell("C", 8, "=SEQUENCE(3,1)");
    setCell("A", 10, "=SEQUENCE(1,3)");

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
      "    | A   | B   | C       | D   | E  
      -----+-----+-----+---------+-----+----
         1 | 5   |     |         |     |    
         2 |     |     |         |     |    
         3 |     |     |         |     |    
         4 |     |     |         |     |    
         5 |     |     |         |     |    
         6 |     |     |         |     |    
         7 |     |     |         |     |    
         8 |     |     | #SPILL! |     |    
         9 |     |     |         |     |    
        10 | 1   | 2   | 3       |     |    
      "
    `);

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
        "empty:B10[A10]",
        "range:B11:B∞[A10]",
        "empty:C10[C8,A10]",
        "range:C11:C∞[C8]",
        "empty:D10[A10]",
        "range:D11:D∞[A10,C8]",
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
        "empty:B10[A10]",
        "range:B11:B∞[A10]",
      ]
    `);

    // X1:Y2 is [1,2;3,4], multiplied by 10 = [10,20;30,40]
    // So A10:B11 = 10+20+30+40 = 100
    expect(cell("A1", true)).toBe(100);
  });
});
