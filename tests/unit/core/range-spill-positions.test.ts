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

const read = (engine: FormulaEngine, refs: string[]) =>
  refs.map((ref) => engine.getCellValue(address(ref)));

/**
 * Cells that a spill places inside a range must be reported at their position
 * in the range, whichever cell the spill is anchored at.
 */
describe("ranges that contain spilled cells", () => {
  test("keep the position of a SEQUENCE spill below the range start", () => {
    const engine = build({
      A1: "x",
      A2: "=SEQUENCE(3)",
      C1: "=UNIQUE(A1:A5)",
      D1: "=MATCH(2, A1:A5, 0)",
    });
    expect(read(engine, ["C1", "C2", "C3", "C4", "C5"])).toEqual([
      "x",
      1,
      2,
      3,
      "",
    ]);
    expect(read(engine, ["D1"])).toEqual([3]);
  });

  test("keep the position of a spilled range reference", () => {
    const engine = build({
      A1: "x",
      A2: "=B1:B3",
      B1: 1,
      B2: 2,
      B3: 3,
      C1: "=UNIQUE(A1:A5)",
      D1: "=MATCH(3, A1:A5, 0)",
    });
    expect(read(engine, ["C1", "C2", "C3", "C4", "C5"])).toEqual([
      "x",
      1,
      2,
      3,
      "",
    ]);
    expect(read(engine, ["D1"])).toEqual([4]);
  });

  test("keep positions in open-ended ranges", () => {
    const engine = build({
      A1: "x",
      A2: "=SEQUENCE(3)",
      C1: "=MATCH(3, A:A, 0)",
      D1: "=UNIQUE(A:A)",
    });
    expect(read(engine, ["C1"])).toEqual([4]);
    expect(read(engine, ["D1", "D2", "D3", "D4", "D5"])).toEqual([
      "x",
      1,
      2,
      3,
      "",
    ]);
  });

  test("keep positions for a range starting inside a spill", () => {
    const engine = build({
      A2: "=SEQUENCE(3)",
      C1: "=MATCH(3, A3:A6, 0)",
      D1: "=MATCH(3, F3:F6, 0)",
      F2: "=G1:G3",
      G1: 1,
      G2: 2,
      G3: 3,
    });
    expect(read(engine, ["C1", "D1"])).toEqual([2, 2]);
  });

  test("keep positions when re-reading a restored warm spill", () => {
    const engine = build({
      F2: "=G1:G3",
      G1: 1,
      G2: 2,
      G3: 3,
      D1: "=MATCH(3, F3:F6, 0)",
    });
    expect(read(engine, ["D1"])).toEqual([2]);

    const restored = FormulaEngine.buildEmpty();
    restored.resetToSerializedEngine(engine.serializeEngine());
    // Editing F5 re-reads F3:F6 while F2 keeps its restored spill result.
    restored.setCellContent(address("F5"), 3);
    expect(read(restored, ["D1", "F2", "F3", "F4"])).toEqual([2, 1, 2, 3]);
  });
});
