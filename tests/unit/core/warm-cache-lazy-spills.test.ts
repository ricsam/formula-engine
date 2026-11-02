import { describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { deserialize } from "../../../src/core/map-serializer";
import type { SerializedCellValue } from "../../../src/core/types";
import { cellAddressToKey, parseCellReference } from "../../../src/core/utils";

const workbookName = "Book";
const sheetName = "Sheet";
const otherSheetName = "Other";

const address = (ref: string, sheet = sheetName) => ({
  workbookName,
  sheetName: sheet,
  ...parseCellReference(ref),
});

function build(
  cells: Record<string, SerializedCellValue>,
  otherCells: Record<string, SerializedCellValue> = {}
) {
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({ workbookName });
  engine.addSheet({ workbookName, sheetName });
  engine.addSheet({ workbookName, sheetName: otherSheetName });
  engine.setSheetContent(
    { workbookName, sheetName: otherSheetName },
    new Map(Object.entries(otherCells))
  );
  engine.setSheetContent(
    { workbookName, sheetName },
    new Map(Object.entries(cells))
  );
  return engine;
}

const read = (engine: FormulaEngine, refs: string[]) =>
  refs.map((ref) => engine.getCellValue(address(ref)));

/** Restores a new engine from a snapshot that keeps evaluation state. */
function warmRestore(engine: FormulaEngine): FormulaEngine {
  const restored = FormulaEngine.buildEmpty();
  restored.resetToSerializedEngine(engine.serializeEngine());
  return restored;
}

const cellKey = (ref: string) => cellAddressToKey(address(ref));

/**
 * Reads `refs` so that the formulas and their spills are warm, restores a warm
 * snapshot, applies `edit` and returns the values of `refs` in the restored
 * engine. They are also checked against the original engine after the same
 * edit, which never went through a snapshot.
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

/**
 * A spill over a range is lazy: its cells read the range when they are
 * displayed, so the node holding the spill never depends on those cells. A warm
 * snapshot stores such a spill as fixed values; it must also store that the
 * values depend on the cells they were read from, or an edit after restoring
 * would keep serving the old values.
 */
describe("warm snapshots of lazy spills", () => {
  const data: Record<string, SerializedCellValue> = {
    A1: 1,
    A2: 2,
    B1: 3,
    B2: 4,
  };

  test("operator spills over a range recalculate after an edit", () => {
    expect(
      readAfterWarmRestoreAndEdit(
        { ...data, D1: "=A1:A2*1", E1: "=-A1:A2" },
        ["D1", "D2", "E1", "E2"],
        setCell("A1", 10)
      )
    ).toEqual([10, 2, -10, -2]);
  });

  test("TOCOL and HSTACK spills recalculate after an edit", () => {
    expect(
      readAfterWarmRestoreAndEdit(
        { ...data, D1: "=TOCOL(A1:B2)", F1: "=HSTACK(A1:A2,B1:B2)" },
        ["D1", "D2", "D3", "D4", "F1", "G1", "F2", "G2"],
        (engine) => {
          setCell("A1", 10)(engine);
          setCell("B2", 40)(engine);
        }
      )
    ).toEqual([10, 3, 2, 40, 10, 3, 2, 40]);
  });

  test("UNIQUE and SEQUENCE based spills recalculate after an edit", () => {
    expect(
      readAfterWarmRestoreAndEdit(
        {
          ...data,
          A3: 1,
          D1: "=UNIQUE(A1:A3)",
          E1: "=UNIQUE(A1:A3)*10",
          F1: "=SEQUENCE(A2)",
          G1: "=SEQUENCE(2)*A1",
        },
        ["D1", "D2", "D3", "E1", "E2", "E3", "F1", "F2", "F3", "G1", "G2"],
        (engine) => {
          setCell("A1", 3)(engine);
          setCell("A2", 3)(engine);
        }
      )
    ).toEqual([3, 1, "", 30, 10, "", 1, 2, 3, 3, 6]);
  });

  test("formulas that consume a lazy spill recalculate after an edit", () => {
    expect(
      readAfterWarmRestoreAndEdit(
        {
          ...data,
          D1: "=A1:A2*1",
          E1: "=SUM(A1:A2*1)",
          E2: "=INDEX(A1:A2*1,2)",
          E3: "=SUM(TOCOL(A1:B2))",
          E4: "=SUM(D1:D2)",
          E5: "=D2",
          E6: "=INDEX(D1,2)",
        },
        ["D1", "D2", "E1", "E2", "E3", "E4", "E5", "E6"],
        setCell("A2", 20)
      )
    ).toEqual([1, 20, 21, 20, 28, 21, 20, 20]);
  });

  test("a spill read through several formulas recalculates after an edit", () => {
    // E1 and G1 share one formula node.
    expect(
      readAfterWarmRestoreAndEdit(
        {
          ...data,
          D1: "=TOCOL(A1:B2)*2",
          E1: "=A1:A2*1",
          G1: "=A1:A2*1",
        },
        ["D1", "D2", "D3", "D4", "E1", "E2", "G1", "G2"],
        (engine) => {
          setCell("A1", 10)(engine);
          setCell("B1", 30)(engine);
        }
      )
    ).toEqual([20, 60, 4, 8, 10, 2, 10, 2]);
  });

  test("a spill over cells of another sheet recalculates after an edit", () => {
    const engine = build(
      { D1: `=${otherSheetName}!A1:A2*1` },
      { A1: 1, A2: 2 }
    );
    expect(read(engine, ["D1", "D2"])).toEqual([1, 2]);

    const restored = warmRestore(engine);
    restored.setCellContent(address("A2", otherSheetName), 20);

    expect(read(restored, ["D1", "D2"])).toEqual([1, 20]);
  });

  test("a spill that read a blank cell recalculates once the cell is filled", () => {
    expect(
      readAfterWarmRestoreAndEdit(
        { ...data, A3: 5, D1: "=TOCOL(A1:B3)" },
        ["D1", "D2", "D3", "D4", "D5", "D6"],
        setCell("B3", 9)
      )
    ).toEqual([1, 3, 2, 4, 5, 9]);
  });

  test("a partly displayed spill over blank cells restores correctly", () => {
    // Only part of the spill was read, so storing it would read blank cells
    // that were never evaluated: it is left out with the formulas built on it.
    const engine = build({
      A1: 1,
      A2: 2,
      A10: 10,
      C1: "=TOCOL(A1:A10)",
      D1: "=SUM(C1:C10)",
      E1: "=C2",
      F1: "=UNIQUE(TOCOL(A1:A10))",
      G1: "=A10+1",
    });
    expect(read(engine, ["C1", "C2", "D1", "E1", "F1", "G1"])).toEqual([
      1, 2, 13, 2, 1, 11,
    ]);

    const restored = warmRestore(engine);
    expect(
      restored._dependencyManager.getCellValueNode(cellKey("G1")).resolved
    ).toBe(true);
    expect(
      read(restored, ["C1", "C2", "C5", "C10", "D1", "E1", "F1", "F2", "G1"])
    ).toEqual([1, 2, "", 10, 13, 2, 1, 2, 11]);

    restored.setCellContent(address("A2"), 20);
    restored.setCellContent(address("A5"), 5);
    expect(
      read(restored, ["C1", "C2", "C5", "C10", "D1", "E1", "F1", "F2", "F3"])
    ).toEqual([1, 20, 5, 10, 36, 20, 1, 20, ""]);
  });

  test("a restored engine serializes again into a snapshot that recalculates", () => {
    const engine = build({ ...data, D1: "=A1:A2*1", E1: "=SUM(D1:D2)" });
    expect(read(engine, ["D1", "D2", "E1"])).toEqual([1, 2, 3]);

    const first = warmRestore(engine);
    first.setCellContent(address("A1", sheetName), 10);
    expect(read(first, ["D1", "D2", "E1"])).toEqual([10, 2, 12]);

    const second = warmRestore(first);
    expect(read(second, ["D1", "D2", "E1"])).toEqual([10, 2, 12]);
    second.setCellContent(address("A2", sheetName), 20);
    expect(read(second, ["D1", "D2", "E1"])).toEqual([10, 20, 30]);
  });

  test("stores a lazy spill warm, depending on the cells it read", () => {
    const engine = build({ ...data, D1: "=A1:A2*1" });
    expect(read(engine, ["D1", "D2"])).toEqual([1, 2]);

    const snapshot = deserialize(engine.serializeEngine()) as any;
    const nodes: any[] = snapshot.managers.dependency.nodes;
    const ast = nodes.find(
      (node) => node.kind === "ast" && node.key === "ast:A1:A2*1"
    );
    const spillMeta = nodes.find(
      (node) =>
        node.kind === "spill-meta" &&
        node.key === cellKey("D1").replace("cell-value:", "spill-meta:")
    );
    for (const node of [ast, spillMeta]) {
      expect(node?.evaluationResult.spill.kind).toBe("materialized");
      expect(node.dependencies).toContain(cellKey("A1"));
      expect(node.dependencies).toContain(cellKey("A2"));
    }

    // The spill is restored as computed, not re-evaluated.
    const restored = warmRestore(engine);
    expect(
      restored._dependencyManager.getCellValueNode(cellKey("D1")).resolved
    ).toBe(true);
    expect(read(restored, ["D1", "D2"])).toEqual([1, 2]);
  });

  test("leaves out the nodes that depend on a spill it cannot store", () => {
    // An unbounded spill cannot be stored as values. Everything computed from
    // it is left out too, rather than restored without that dependency.
    const engine = build({
      A1: 1,
      F1: "=SEQUENCE(INFINITY,1,A1)",
      G1: "=F3+1",
      H1: "=INDEX(SEQUENCE(INFINITY,1,A1),2)",
      I1: "=A1+1",
    });
    expect(read(engine, ["F1", "F3", "G1", "H1", "I1"])).toEqual([
      1, 3, 4, 2, 2,
    ]);

    const snapshot = deserialize(engine.serializeEngine()) as any;
    const keys = new Set(
      (snapshot.managers.dependency.nodes as any[]).map((node) => node.key)
    );
    for (const ref of ["F1", "G1", "H1"]) {
      expect(keys.has(cellKey(ref))).toBe(false);
    }
    expect(keys.has(cellKey("I1"))).toBe(true);

    const restored = warmRestore(engine);
    expect(read(restored, ["F1", "F3", "G1", "H1", "I1"])).toEqual([
      1, 3, 4, 2, 2,
    ]);

    restored.setCellContent(address("A1"), 10);
    expect(read(restored, ["F1", "F3", "G1", "H1", "I1"])).toEqual([
      10, 12, 13, 11, 11,
    ]);
  });
});
