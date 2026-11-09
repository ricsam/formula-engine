import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SerializedCellValue } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

/**
 * Warm snapshots store spills as fixed values together with the cells they
 * were read from. These tests check that the spills of the functions added for
 * the Operations workbooks (FILTER, TRANSPOSE, IFS, ISBLANK, ROUND, VLOOKUP)
 * survive a warm restore and still follow edits afterwards.
 */

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

function warmRestore(engine: FormulaEngine): FormulaEngine {
  const restored = FormulaEngine.buildEmpty();
  restored.resetToSerializedEngine(engine.serializeEngine());
  return restored;
}

/**
 * Warms `refs`, restores a warm snapshot, applies `edit` to both engines and
 * returns the restored engine's values, checked against the original's.
 */
function readAfterWarmRestoreAndEdit(
  cells: Record<string, SerializedCellValue>,
  refs: string[],
  edit: (engine: FormulaEngine) => void
) {
  const engine = build(cells);
  const before = read(engine, refs);

  const restored = warmRestore(engine);
  expect(read(restored, refs)).toEqual(before);

  edit(engine);
  edit(restored);
  const after = read(restored, refs);
  expect(after).toEqual(read(engine, refs));
  return after;
}

const setCell =
  (ref: string, content: SerializedCellValue) => (engine: FormulaEngine) =>
    engine.setCellContent(address(ref), content);

const data: Record<string, SerializedCellValue> = {
  A1: "North",
  B1: 10.25,
  A2: "South",
  B2: 20.75,
  A3: "North",
  B3: 30.5,
};

describe("warm snapshots of the new array functions", () => {
  test("FILTER follows an edit of a selected cell and of the condition", () => {
    const cells = { ...data, D1: '=FILTER(B1:B3, A1:A3="North")' };

    expect(
      readAfterWarmRestoreAndEdit(cells, ["D1", "D2"], setCell("B3", 99))
    ).toEqual([10.25, 99]);
    expect(
      readAfterWarmRestoreAndEdit(cells, ["D1", "D2", "D3"], setCell("A2", "North"))
    ).toEqual([10.25, 20.75, 30.5]);
  });

  test("TRANSPOSE follows an edit of a source cell", () => {
    const cells = { ...data, D1: "=TRANSPOSE(A1:B3)" };

    expect(
      readAfterWarmRestoreAndEdit(cells, ["D1", "E1", "F2"], setCell("B3", 7))
    ).toEqual(["North", "South", 7]);
  });

  test("IFS with an array condition follows an edit", () => {
    const cells = {
      ...data,
      D1: '=IFS(B1:B3>20, "high", TRUE, "low")',
    };

    expect(
      readAfterWarmRestoreAndEdit(cells, ["D1", "D2", "D3"], setCell("B1", 50))
    ).toEqual(["high", "high", "high"]);
  });

  test("ISBLANK and ROUND over a range follow edits", () => {
    const cells = {
      ...data,
      C2: 1,
      D1: "=ISBLANK(C1:C3)",
      E1: "=ROUND(B1:B3, 0)",
    };

    expect(
      readAfterWarmRestoreAndEdit(cells, ["D1", "D2", "D3"], setCell("C3", "x"))
    ).toEqual([true, false, false]);
    expect(
      readAfterWarmRestoreAndEdit(cells, ["E1", "E2", "E3"], setCell("B2", 2.5))
    ).toEqual([10, 3, 31]);
  });

  test("VLOOKUP over an array of lookup values follows an edit", () => {
    const cells = {
      ...data,
      D1: '=VLOOKUP({"South";"North"}, A1:B3, 2, FALSE)',
    };

    expect(
      readAfterWarmRestoreAndEdit(cells, ["D1", "D2"], setCell("B1", 1))
    ).toEqual([20.75, 1]);
  });
});
