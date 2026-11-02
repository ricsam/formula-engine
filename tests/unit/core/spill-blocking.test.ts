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

describe("blocked spills", () => {
  test("spill once the blocking value is cleared", () => {
    const engine = build({ A1: "=SEQUENCE(3)", A2: "x" });
    expect(read(engine, ["A1", "A2"])).toEqual(["#SPILL!", "x"]);

    engine.setCellContent(address("A2"), "");
    expect(read(engine, ["A1", "A2", "A3"])).toEqual([1, 2, 3]);
  });

  test("stay blocked while a blocker remains in the area", () => {
    const engine = build({ A1: "=SEQUENCE(3)", A2: "x", A3: "y" });
    expect(read(engine, ["A1"])).toEqual(["#SPILL!"]);

    engine.setCellContent(address("A2"), "");
    expect(read(engine, ["A1", "A3"])).toEqual(["#SPILL!", "y"]);

    engine.setCellContent(address("A3"), "");
    expect(read(engine, ["A1", "A2", "A3"])).toEqual([1, 2, 3]);
  });

  test("recover from a blocker that depends on the spilling cell", () => {
    const engine = build({ A1: "=SEQUENCE(3)", A2: "=A1*10" });
    expect(read(engine, ["A1", "A2"])).toEqual(["#SPILL!", "#SPILL!"]);

    engine.setCellContent(address("A2"), "");
    expect(read(engine, ["A1", "A2", "A3"])).toEqual([1, 2, 3]);
  });

  test("take the area when the spill that blocked them goes away", () => {
    const engine = build({ A2: "=SEQUENCE(1, 3)", B1: "=SEQUENCE(3)" });
    expect(read(engine, ["A2", "B2", "C2"])).toEqual([1, 2, 3]);
    expect(read(engine, ["B1"])).toEqual(["#SPILL!"]);

    engine.setCellContent(address("A2"), 5);
    expect(read(engine, ["A2", "B1", "B2", "B3", "C2"])).toEqual([
      5,
      1,
      2,
      3,
      "",
    ]);
  });

  test("recover after a warm serialization round trip", () => {
    const engine = build({ A1: "=SEQUENCE(3)", A3: "x" });
    expect(read(engine, ["A1"])).toEqual(["#SPILL!"]);

    const restored = FormulaEngine.buildEmpty();
    restored.resetToSerializedEngine(engine.serializeEngine());
    expect(read(restored, ["A1", "A3"])).toEqual(["#SPILL!", "x"]);

    restored.setCellContent(address("A3"), "");
    expect(read(restored, ["A1", "A2", "A3"])).toEqual([1, 2, 3]);
  });
});
