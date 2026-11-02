import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../core/engine";
import { FormulaError, type SerializedCellValue } from "../core/types";
import { parseCellReference } from "../core/utils";
import {
  getBroadcastShape,
  getBroadcastSpillArea,
  mapBroadcastOffset,
} from "./broadcast";

describe("broadcast helpers", () => {
  test("the result is as large as the largest operand in each dimension", () => {
    expect(
      getBroadcastShape([
        { rows: 1, cols: 3 },
        { rows: 2, cols: 1 },
      ])
    ).toEqual({ rows: 2, cols: 3 });
    expect(
      getBroadcastShape([
        { rows: Infinity, cols: 1 },
        { rows: 1, cols: 4 },
        { rows: 1, cols: 1 },
      ])
    ).toEqual({ rows: Infinity, cols: 4 });
  });

  test("a dimension of size 1 repeats and a shorter one ends in #N/A", () => {
    expect(mapBroadcastOffset({ rows: 1, cols: 3 }, { x: 2, y: 5 })).toEqual({
      x: 2,
      y: 0,
    });
    expect(mapBroadcastOffset({ rows: 2, cols: 1 }, { x: 7, y: 1 })).toEqual({
      x: 0,
      y: 1,
    });
    expect(
      mapBroadcastOffset({ rows: 2, cols: 2 }, { x: 0, y: 2 })
    ).toBeUndefined();
    expect(
      mapBroadcastOffset({ rows: Infinity, cols: 1 }, { x: 3, y: 1_000_000 })
    ).toEqual({ x: 0, y: 1_000_000 });
  });

  test("the spill area starts at the origin", () => {
    const origin = {
      workbookName: "W",
      sheetName: "S",
      colIndex: 2,
      rowIndex: 4,
    };
    expect(getBroadcastSpillArea({ rows: 3, cols: 4 }, origin)).toEqual({
      start: { col: 2, row: 4 },
      end: {
        col: { type: "number", value: 5 },
        row: { type: "number", value: 6 },
      },
    });
    expect(
      getBroadcastSpillArea({ rows: Infinity, cols: 2 }, origin).end.row
    ).toEqual({ type: "infinity", sign: "positive" });
  });
});

describe("array broadcasting in binary operators", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const setContent = (entries: [string, SerializedCellValue][]) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(entries)
    );
  };

  /** Reads a block of cells row by row, e.g. block("C1", 3, 4). */
  const block = (topLeft: string, rows: number, cols: number) => {
    const start = parseCellReference(topLeft);
    const values: SerializedCellValue[][] = [];
    for (let y = 0; y < rows; y++) {
      const row: SerializedCellValue[] = [];
      for (let x = 0; x < cols; x++) {
        row.push(
          engine.getCellValue({
            sheetName,
            workbookName,
            colIndex: start.colIndex + x,
            rowIndex: start.rowIndex + y,
          })
        );
      }
      values.push(row);
    }
    return values;
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("row and column operands", () => {
    test("a column times a row is a multiplication table", () => {
      setContent([["A1", "=SEQUENCE(3)*SEQUENCE(1,4)"]]);

      expect(block("A1", 3, 4)).toEqual([
        [1, 2, 3, 4],
        [2, 4, 6, 8],
        [3, 6, 9, 12],
      ]);
      // The spill ends at the broadcast shape.
      expect(cell("E1")).toBe("");
      expect(cell("A4")).toBe("");
    });

    test("SEQUENCE with an omitted rows argument broadcasts as a row", () => {
      setContent([["A1", "=SEQUENCE(3)*SEQUENCE(,4)"]]);

      expect(block("A1", 3, 4)).toEqual([
        [1, 2, 3, 4],
        [2, 4, 6, 8],
        [3, 6, 9, 12],
      ]);
    });

    test("a row plus a column of ranges", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", 10],
        ["C1", 20],
        ["E1", "=B1:C1+A1:A3"],
      ]);

      expect(block("E1", 3, 2)).toEqual([
        [11, 21],
        [12, 22],
        [13, 23],
      ]);
    });

    test("comparison operators broadcast", () => {
      setContent([["A1", "=SEQUENCE(3)=SEQUENCE(1,3)"]]);

      expect(block("A1", 3, 3)).toEqual([
        [true, false, false],
        [false, true, false],
        [false, false, true],
      ]);
    });

    test("concatenation broadcasts into a cartesian product", () => {
      setContent([
        ["A1", "a"],
        ["A2", "b"],
        ["B1", "x"],
        ["C1", "y"],
        ["D1", "z"],
        ["F1", '=A1:A2&"-"&B1:D1'],
      ]);

      expect(block("F1", 2, 3)).toEqual([
        ["a-x", "a-y", "a-z"],
        ["b-x", "b-y", "b-z"],
      ]);
    });

    test("array literals broadcast", () => {
      setContent([
        ["A1", "={1,2,3}*{10;20}"],
        ["A4", "={1;2;3}+{10;20}"],
        ["C4", "=SUM({1,2,3}*{10;20})"],
      ]);

      expect(block("A1", 2, 3)).toEqual([
        [10, 20, 30],
        [20, 40, 60],
      ]);
      expect(block("A4", 3, 1)).toEqual([[11], [22], [FormulaError.NA]]);
      expect(cell("C4")).toBe(180);
    });

    test("an operand of one cell in both directions is a single value", () => {
      setContent([
        ["A1", 5],
        ["C1", "=SEQUENCE(2,2)*A1:A1"],
      ]);

      expect(block("C1", 2, 2)).toEqual([
        [5, 10],
        [15, 20],
      ]);
    });
  });

  describe("arrays of mismatched sizes", () => {
    test("cells beyond the shorter array are #N/A", () => {
      setContent([["A1", "=SEQUENCE(3)+SEQUENCE(2)"]]);

      expect(cell("A1")).toBe(2);
      expect(cell("A2")).toBe(4);
      expect(cell("A3")).toBe(FormulaError.NA);
      expect(cell("A4")).toBe("");
    });

    test("both dimensions are padded independently", () => {
      setContent([["A1", "=SEQUENCE(2,3)+SEQUENCE(3,2)"]]);

      expect(block("A1", 3, 3)).toEqual([
        [2, 4, FormulaError.NA],
        [7, 9, FormulaError.NA],
        [FormulaError.NA, FormulaError.NA, FormulaError.NA],
      ]);
    });

    test("cells below a shorter range are not read", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", 10],
        ["B2", 20],
        // Outside B1:B2: it must not be paired with A3.
        ["B3", 30],
        ["C1", "=A1:A3+B1:B2"],
      ]);

      expect(cell("C1")).toBe(11);
      expect(cell("C2")).toBe(22);
      expect(cell("C3")).toBe(FormulaError.NA);
    });

    test("consumers see the #N/A cells", () => {
      setContent([
        // UNIQUE reads every cell and returns the first error it finds.
        ["A1", "=UNIQUE(SEQUENCE(3)+SEQUENCE(2))"],
        ["B1", "=UNIQUE(SEQUENCE(2)+SEQUENCE(2))"],
        ["C1", "=IFERROR(SEQUENCE(3)+SEQUENCE(2), 100)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.NA);
      expect(block("B1", 2, 1)).toEqual([[2], [4]]);
      expect(block("C1", 3, 1)).toEqual([[2], [4], [100]]);
    });
  });

  describe("arrays and single values", () => {
    test("an array with a single value keeps its shape", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", "=A1:A3*10"],
        ["C1", "=10-A1:A3"],
        ["D1", "=SEQUENCE(2,2)+1"],
      ]);

      expect(block("B1", 3, 2)).toEqual([
        [10, 9],
        [20, 8],
        [30, 7],
      ]);
      expect(block("D1", 2, 2)).toEqual([
        [2, 3],
        [4, 5],
      ]);
      expect(cell("B4")).toBe("");
    });

    test("an error single value is the whole result", () => {
      setContent([["A1", "=SEQUENCE(3)+CEILING(5)"]]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe("");
    });
  });

  describe("blank cells inside operands", () => {
    test("a blank cell reads like it does in a single-cell formula", () => {
      setContent([
        ["A1", 1],
        // A2 is blank
        ["A3", "x"],
        ["B1", 2],
        ["B2", 3],
        // B3 is blank
        ["D1", "=A1:A3+B1:B3"],
        ["E1", "=A1+B1"],
        ["E2", "=A2+B2"],
        ["E3", "=A3+B3"],
        ["F1", "=A1:A3&B1:B3"],
        ["G1", "=A1:A3=B1:B3"],
      ]);

      // Arithmetic treats a blank cell exactly as =A2+B2 does.
      expect(cell("D1")).toBe(3);
      expect([cell("D2"), cell("D3")]).toEqual([cell("E2"), cell("E3")]);

      // Concatenation reads a blank cell as empty text.
      expect(block("F1", 3, 1)).toEqual([["12"], ["3"], ["x"]]);
      expect(block("G1", 3, 1)).toEqual([[false], [false], [false]]);
    });

    test("cells blank in every operand are skipped by aggregates", () => {
      setContent([
        ["A1", 1],
        ["A3", 3],
        ["B1", 10],
        ["B3", 30],
        // Row 2 is blank in both ranges
        ["D1", "=SUM(A1:A3*B1:B3)"],
        ["D2", "=COUNT(A1:A3*B1:B3)"],
      ]);

      expect(cell("D1")).toBe(100);
      expect(cell("D2")).toBe(2);
    });
  });

  describe("aggregates over broadcast results", () => {
    test("each operator applies once per cell", () => {
      setContent([
        ["A1", "=SUM(SEQUENCE(4)/2)"],
        ["A2", "=SUM(SEQUENCE(4)*SEQUENCE(4))"],
        ["A3", "=SUM(SEQUENCE(3)*2)"],
        ["A4", "=SUM(2*SEQUENCE(3))"],
        ["A5", "=SUM((SEQUENCE(3)+1)*10)"],
        ["B1", 1],
        ["B2", 2],
        ["B3", 2],
        ["A6", "=SUM(UNIQUE(B1:B3)*10)"],
      ]);

      expect(cell("A1")).toBe(5);
      expect(cell("A2")).toBe(30);
      expect(cell("A3")).toBe(12);
      expect(cell("A4")).toBe(12);
      expect(cell("A5")).toBe(90);
      expect(cell("A6")).toBe(30);
    });

    test("SUM and COUNT read every cell of a broadcast table", () => {
      setContent([
        ["A1", "=SUM(SEQUENCE(3)*SEQUENCE(1,4))"],
        ["A2", "=COUNT(SEQUENCE(3)*SEQUENCE(1,4))"],
        ["A3", "=SUM(SEQUENCE(3)=SEQUENCE(1,3))"],
        ["A4", "=UNIQUE(SEQUENCE(2)&SEQUENCE(1,2))"],
      ]);

      expect(cell("A1")).toBe(60);
      expect(cell("A2")).toBe(12);
      // Booleans are not summed.
      expect(cell("A3")).toBe(0);
      expect(block("A4", 2, 2)).toEqual([
        ["11", "12"],
        ["21", "22"],
      ]);
    });

    test("a row of a range times a column of a range", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", 10],
        ["C1", 100],
        ["E1", "=SUM(A1:A3*B1:C1)"],
        ["E2", "=COUNT(A1:A3*B1:C1)"],
      ]);

      expect(cell("E1")).toBe(660);
      expect(cell("E2")).toBe(6);
    });
  });

  describe("open-ended operands", () => {
    test("open-ended ranges of the same shape pair their populated cells", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["B1", 4],
        ["B2", 5],
        ["B3", 6],
        ["D1", "=SUM(A:A*B:B)"],
        ["D2", "=SUM(A:A*2)"],
        ["D3", "=SUM(A2:A*B2:B)"],
      ]);

      expect(cell("D1")).toBe(32);
      expect(cell("D2")).toBe(12);
      expect(cell("D3")).toBe(28);
    });

    test("a row is repeated along the populated cells of an open-ended column", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["A3", 3],
        ["C1", 10],
        ["D1", 100],
        ["F1", "=SUM(A:A*C1:D1)"],
        ["F2", "=COUNT(A:A*C1:D1)"],
      ]);

      expect(cell("F1")).toBe(660);
      expect(cell("F2")).toBe(6);
    });

    test("an open-ended result spills lazily", () => {
      setContent([["B1", "=SEQUENCE(INFINITY)*SEQUENCE(1,2)"]]);

      expect(cell("B1")).toBe(1);
      expect(cell("C1")).toBe(2);
      expect(cell("B500")).toBe(500);
      expect(cell("C500")).toBe(1000);
      expect(cell("D1")).toBe("");
    });

    test("an open range reads the part of a broadcast spill it overlaps", () => {
      setContent([
        ["A10", "=SEQUENCE(3)*SEQUENCE(1,3)"],
        ["A1", "=SUM(B10:B)"],
        ["A2", "=SUM(B10:INFINITY)"],
        ["A3", "=SUM(A11:C11)"],
      ]);

      // A10:C12 holds [1,2,3; 2,4,6; 3,6,9]
      expect(cell("A1")).toBe(12);
      expect(cell("A2")).toBe(30);
      expect(cell("A3")).toBe(12);
    });

    test("an open range reads an overlapped broadcast of ranges", () => {
      setContent([
        ["X1", 1],
        ["X2", 2],
        ["X3", 3],
        ["Y1", 10],
        ["Z1", 100],
        ["A10", "=X1:X3*Y1:Z1"],
        ["A1", "=SUM(B10:B)"],
        ["A2", "=SUM(A11:B)"],
      ]);

      // A10:B12 holds [10,100; 20,200; 30,300]
      expect(cell("A1")).toBe(600);
      expect(cell("A2")).toBe(550);
    });

    test("an open range bounds an unbounded broadcast spill", () => {
      setContent([
        ["A1", "=SUM(B10:D)"],
        ["A100", "=SEQUENCE(1,INFINITY)*2"],
      ]);

      // B100:D100 holds 4, 6, 8
      expect(cell("A1")).toBe(18);
    });
  });

  describe("recalculation", () => {
    test("spilled cells and aggregates follow edits of either operand", () => {
      setContent([
        ["A1", 1],
        ["A2", 2],
        ["B1", 10],
        ["C1", 100],
        ["E1", "=A1:A2*B1:C1"],
        ["H1", "=SUM(A1:A2*B1:C1)"],
      ]);

      expect(block("E1", 2, 2)).toEqual([
        [10, 100],
        [20, 200],
      ]);
      expect(cell("H1")).toBe(330);

      setCellContent("A2", 3);
      expect(block("E1", 2, 2)).toEqual([
        [10, 100],
        [30, 300],
      ]);
      expect(cell("H1")).toBe(440);

      setCellContent("C1", 1000);
      expect(block("E1", 2, 2)).toEqual([
        [10, 1000],
        [30, 3000],
      ]);
      expect(cell("H1")).toBe(4040);
    });

    test("the shape follows the operands", () => {
      setContent([
        ["A1", 2],
        ["B1", 3],
        ["D1", "=SEQUENCE(A1)*SEQUENCE(1,B1)"],
      ]);

      expect(cell("F2")).toBe(6);
      expect(cell("D3")).toBe("");

      setCellContent("A1", 3);
      expect(cell("D3")).toBe(3);
      expect(cell("F3")).toBe(9);

      setCellContent("B1", 1);
      expect(cell("E1")).toBe("");
      expect(cell("D3")).toBe(3);
    });

    test("a blocked broadcast spill is #SPILL!", () => {
      setContent([
        ["A1", "=SEQUENCE(2)*SEQUENCE(1,2)"],
        ["B2", "blocker"],
      ]);

      expect(cell("A1")).toBe(FormulaError.SPILL);

      setCellContent("B2", "");
      expect(cell("A1")).toBe(1);
      expect(cell("B2")).toBe(4);
    });
  });
});
