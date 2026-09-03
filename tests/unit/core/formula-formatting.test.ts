import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SerializedCellValue } from "../../../src/core/types";

const workbookName = "Book";
const sheetName = "Sheet1";
const address = {
  workbookName,
  sheetName,
  rowIndex: 0,
  colIndex: 0,
};

function buildEngine(): FormulaEngine {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });
  engine.clearUndoRedoHistory();
  return engine;
}

function stored(engine: FormulaEngine, cell = "A1"): SerializedCellValue {
  return engine.getSheetSerialized({ workbookName, sheetName }).get(cell);
}

describe("formula content formatting", () => {
  test("stores valid formulas in compact form", () => {
    const engine = buildEngine();

    engine.setCellContent(address, "= IF( TRUE,\n SUM(1,2), 0 )");

    expect(stored(engine)).toBe("=IF(TRUE, SUM(1, 2), 0)");
    expect(engine.getCellValue(address)).toBe(3);
  });

  test("an equivalent layout-only edit creates no update or undo entry", () => {
    const engine = buildEngine();
    let updates = 0;
    const unsubscribe = engine.onUpdate(() => {
      updates++;
    });

    engine.setCellContent(address, "=SUM(1,2)");
    expect(updates).toBe(1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);

    engine.setCellContent(address, "= SUM(\n  1,\n  2\n)");
    unsubscribe();

    expect(stored(engine)).toBe("=SUM(1, 2)");
    expect(updates).toBe(1);
    expect(engine.getUndoRedoState().undoDepth).toBe(1);
  });

  test("normalizes formulas supplied through whole-sheet content", () => {
    const engine = buildEngine();

    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, SerializedCellValue>([
        ["A1", "=SUM(\n  1,2\n)"],
        ["A2", "ordinary  text\nwith a newline"],
        ["A3", "=SUM(A1,,B1)"],
      ])
    );

    expect(stored(engine, "A1")).toBe("=SUM(1, 2)");
    expect(stored(engine, "A2")).toBe("ordinary  text\nwith a newline");
    expect(stored(engine, "A3")).toBe("=SUM(A1,,B1)");
  });

  test("preserves ordinary text and invalid formula text exactly", () => {
    const engine = buildEngine();
    const plainText = " SUM( 1,2 )\nnot a formula ";
    const invalidFormula = "=SUM( 1,, 2 )\n";

    engine.setCellContent(address, plainText);
    expect(stored(engine)).toBe(plainText);

    engine.setCellContent(address, invalidFormula);
    expect(stored(engine)).toBe(invalidFormula);
  });
});
