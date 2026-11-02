import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SerializedCellValue } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet";
const address = (ref: string) => ({
  workbookName,
  sheetName,
  ...parseCellReference(ref),
});

/** The user's formula: the cartesian product of the distinct values in A and B. */
const CARTESIAN =
  "=LET(a, UNIQUE(TOCOL(A2:A1000, 1)), b, UNIQUE(TOCOL(B2:B1000, 1)), HSTACK(TOCOL(IF(SEQUENCE(,ROWS(b)), a)), TOCOL(IF(SEQUENCE(ROWS(a)), TOROW(b)))))";

function build(cells: Record<string, SerializedCellValue>) {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName });
  engine.addSheet({ workbookName, sheetName });
  engine.setSheetContent(
    { workbookName, sheetName },
    new Map(Object.entries(cells))
  );
  return engine;
}

function column(values: SerializedCellValue[], col: string, firstRow = 2) {
  return Object.fromEntries(
    values.map((value, index) => [`${col}${firstRow + index}`, value])
  );
}

/** Reads `rows` x `cols` cells starting at `topLeft`, row by row. */
function readBlock(
  engine: FormulaEngine,
  topLeft: string,
  rows: number,
  cols: number
): SerializedCellValue[][] {
  const start = parseCellReference(topLeft);
  const block: SerializedCellValue[][] = [];
  for (let row = 0; row < rows; row++) {
    const values: SerializedCellValue[] = [];
    for (let col = 0; col < cols; col++) {
      values.push(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: start.colIndex + col,
          rowIndex: start.rowIndex + row,
        })
      );
    }
    block.push(values);
  }
  return block;
}

function cartesian(a: SerializedCellValue[], b: SerializedCellValue[]) {
  return a.flatMap((left) => b.map((right) => [left, right]));
}

describe("combinatorial sets with LET, TOCOL, HSTACK and broadcasting", () => {
  test("spills the cartesian product of A and B into C:D", () => {
    const engine = build({
      A1: "A",
      B1: "B",
      ...column([1, 2, 3], "A"),
      ...column([1, 2, 3], "B"),
      C2: CARTESIAN,
    });

    expect(readBlock(engine, "C2", 9, 2)).toEqual(
      cartesian([1, 2, 3], [1, 2, 3])
    );
    expect(engine.getCellValue(address("C11"))).toBe("");
  });

  test("uses distinct non-blank values, in first-seen order", () => {
    const engine = build({
      ...column(["x", "", "y", "X", "z"], "A"),
      A9: "y",
      ...column([10, 10, 20], "B"),
      B12: 30,
      C2: CARTESIAN,
    });

    expect(readBlock(engine, "C2", 9, 2)).toEqual(
      cartesian(["x", "y", "z"], [10, 20, 30])
    );
  });

  test("follows edits to the source columns", () => {
    const engine = build({
      ...column([1, 2], "A"),
      ...column(["p", "q"], "B"),
      C2: CARTESIAN,
    });
    expect(readBlock(engine, "C2", 4, 2)).toEqual(
      cartesian([1, 2], ["p", "q"])
    );

    engine.setCellContent(address("A4"), 3);
    engine.setCellContent(address("B3"), "r");
    expect(readBlock(engine, "C2", 6, 2)).toEqual(
      cartesian([1, 2, 3], ["p", "r"])
    );

    engine.setCellContent(address("A2"), "");
    expect(readBlock(engine, "C2", 4, 2)).toEqual(
      cartesian([2, 3], ["p", "r"])
    );
    expect(engine.getCellValue(address("C6"))).toBe("");
  });

  test("handles a single value on either side", () => {
    const engine = build({
      A2: "only",
      ...column([1, 2, 3], "B"),
      C2: CARTESIAN,
    });
    expect(readBlock(engine, "C2", 3, 2)).toEqual(
      cartesian(["only"], [1, 2, 3])
    );

    engine.setCellContent(address("B3"), "");
    engine.setCellContent(address("B4"), "");
    expect(readBlock(engine, "C2", 1, 2)).toEqual([["only", 1]]);
  });

  test("reports #SPILL! when the result area is occupied", () => {
    const engine = build({
      ...column([1, 2], "A"),
      ...column([1, 2], "B"),
      C2: CARTESIAN,
      D4: "blocker",
    });
    expect(engine.getCellValue(address("C2"))).toBe("#SPILL!");

    engine.setCellContent(address("D4"), "");
    expect(readBlock(engine, "C2", 4, 2)).toEqual(cartesian([1, 2], [1, 2]));
  });

  test("survives a warm serialization round trip", () => {
    const engine = build({
      ...column(["a", "b"], "A"),
      ...column([1, 2, 3], "B"),
      C2: CARTESIAN,
    });
    const expected = cartesian(["a", "b"], [1, 2, 3]);
    expect(readBlock(engine, "C2", 6, 2)).toEqual(expected);

    const restored = FormulaEngine.buildEmpty();
    restored.resetToSerializedEngine(engine.serializeEngine());
    expect(readBlock(restored, "C2", 6, 2)).toEqual(expected);

    restored.setCellContent(address("B5"), 4);
    expect(readBlock(restored, "C2", 8, 2)).toEqual(
      cartesian(["a", "b"], [1, 2, 3, 4])
    );
  });

  test("the new functions compose with the result", () => {
    const engine = build({
      ...column([1, 2, 3], "A"),
      ...column(["p", "q"], "B"),
      C2: CARTESIAN,
      F1: "=ROWS(C2)",
      F2: "=COLUMNS(C2)",
      F3: "=COLS(C2)",
      F4: "=COUNTA(C2)",
      F5: "=COUNTA(UNIQUE(TOCOL(A2:A1000, 1)))",
      // Row index into `a` for each pair, built from INT.
      G2: "=INT((SEQUENCE(6)-1)/2)+1",
    });

    expect(engine.getCellValue(address("F1"))).toBe(6);
    expect(engine.getCellValue(address("F2"))).toBe(2);
    expect(engine.getCellValue(address("F3"))).toBe(2);
    expect(engine.getCellValue(address("F4"))).toBe(12);
    expect(engine.getCellValue(address("F5"))).toBe(3);
    expect(readBlock(engine, "G2", 6, 1).flat()).toEqual([1, 1, 2, 2, 3, 3]);
  });

  test("stays fast for a 60 x 60 product", () => {
    const size = 60;
    const values = Array.from({ length: size }, (_, index) => index + 1);
    const engine = build({
      ...column(values, "A"),
      ...column(values.map((value) => `v${value}`), "B"),
      C2: CARTESIAN,
    });

    const started = performance.now();
    const block = readBlock(engine, "C2", size * size, 2);
    const elapsed = performance.now() - started;

    expect(block[0]).toEqual([1, "v1"]);
    expect(block[size]).toEqual([2, "v1"]);
    expect(block[size * size - 1]).toEqual([size, `v${size}`]);
    expect(
      engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 2,
        rowIndex: size * size + 1,
      })
    ).toBe("");
    expect(elapsed).toBeLessThan(10_000);
  });
});
