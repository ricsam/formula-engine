import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { SerializedCellValue } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";

describe("Incremental invalidation", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const address = (ref: string, targetSheetName = sheetName) => ({
    workbookName,
    sheetName: targetSheetName,
    ...parseCellReference(ref),
  });

  const setCellContent = (
    ref: string,
    content: SerializedCellValue,
    targetSheetName = sheetName
  ) => {
    engine.setCellContent(address(ref, targetSheetName), content);
  };

  const cell = (ref: string, targetSheetName = sheetName) =>
    engine.getCellValue(address(ref, targetSheetName));

  const cellNode = (ref: string, targetSheetName = sheetName) =>
    engine._dependencyManager.getCellValueNode(
      `cell-value:${workbookName}:${targetSheetName}:${ref}`
    );

  const spillMetaNode = (ref: string, targetSheetName = sheetName) =>
    engine._dependencyManager.getSpillMetaNode(
      `spill-meta:${workbookName}:${targetSheetName}:${ref}`
    );

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  test("editing a cell only invalidates its dependency chain", () => {
    setCellContent("A1", 1);
    setCellContent("B1", "=A1+1");
    setCellContent("C1", "=10+1");

    expect(cell("B1")).toBe(2);
    expect(cell("C1")).toBe(11);

    const dependentNode = cellNode("B1");
    const unrelatedNode = cellNode("C1");

    expect(dependentNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);

    setCellContent("A1", 2);

    expect(dependentNode.resolved).toBe(false);
    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("B1")).toBe(3);
    expect(cell("C1")).toBe(11);
  });

  test("empty to value edits invalidate open range consumers without clearing unrelated caches", () => {
    setCellContent("A1", "=SUM(B1:B)");
    setCellContent("D1", "=10+5");

    expect(cell("A1")).toBe(0);
    expect(cell("D1")).toBe(15);

    const unrelatedNode = cellNode("D1");
    expect(unrelatedNode.resolved).toBe(true);

    setCellContent("B5", 5);

    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("A1")).toBe(5);
    expect(cell("D1")).toBe(15);
  });

  test("named expression updates only invalidate dependent formulas", () => {
    engine.addNamedExpression({
      expressionName: "X",
      expression: "1+1",
    });

    setCellContent("A1", "=X");
    setCellContent("B1", "=A1+1");
    setCellContent("C1", "=10+5");

    expect(cell("B1")).toBe(3);
    expect(cell("C1")).toBe(15);

    const dependentNode = cellNode("B1");
    const unrelatedNode = cellNode("C1");

    expect(dependentNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);

    engine.updateNamedExpression({
      expressionName: "X",
      expression: "5",
    });

    expect(dependentNode.resolved).toBe(false);
    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("B1")).toBe(6);
    expect(cell("C1")).toBe(15);
  });

  test("clearing a formula prunes orphaned ASTs without invalidating sibling table formulas", () => {
    engine.setSheetContent(
      { workbookName, sheetName },
      new Map<string, string | number>([
        ["A1", "key"],
        ["B1", "amount"],
        ["C1", "total"],
        ["D1", "=10+1"],
        ["A2", "x"],
        ["A3", "y"],
        ["A4", "x"],
        ["B2", 1],
        ["B3", 2],
        ["B4", 3],
        ["C2", "=SUMIFS(Table1[amount],Table1[key],[@key])"],
        ["C3", "=SUMIFS(Table1[amount],Table1[key],[@key])"],
        ["C4", "=SUMIFS(Table1[amount],Table1[key],[@key])"],
      ])
    );

    engine.addTable({
      tableName: "Table1",
      sheetName,
      workbookName,
      start: "A1",
      numRows: { type: "number", value: 3 },
      numCols: 3,
    });

    expect(cell("C2")).toBe(4);
    expect(cell("C3")).toBe(2);
    expect(cell("C4")).toBe(4);
    expect(cell("D1")).toBe(11);

    const siblingNode = cellNode("C3");
    const otherSiblingNode = cellNode("C4");
    const unrelatedNode = cellNode("D1");

    expect(siblingNode.resolved).toBe(true);
    expect(otherSiblingNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);

    setCellContent("C2", undefined);

    expect(siblingNode.resolved).toBe(true);
    expect(otherSiblingNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("C3")).toBe(2);
    expect(cell("C4")).toBe(4);
    expect(cell("D1")).toBe(11);

    const restoredEngine = FormulaEngine.buildEmpty();
    restoredEngine.addWorkbook({ workbookName: workbookName });
    restoredEngine.addSheet({ workbookName, sheetName });
    restoredEngine.resetToSerializedEngine(engine.serializeEngine());

    expect(restoredEngine.getCellValue(address("C3"))).toBe(2);
    expect(restoredEngine.getCellValue(address("C4"))).toBe(4);
  });

  test("spill shape changes invalidate intersecting consumers and keep unrelated caches warm", () => {
    setCellContent("C1", 1);
    setCellContent("A1", "=IF(C1=1,SEQUENCE(2,2),SEQUENCE(1,1))");
    setCellContent("D1", "=SUM(A1:B2)");
    setCellContent("F1", "=10+5");

    expect(cell("D1")).toBe(10);
    expect(cell("F1")).toBe(15);

    const unrelatedNode = cellNode("F1");
    expect(unrelatedNode.resolved).toBe(true);

    setCellContent("C1", 0);

    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("D1")).toBe(1);
    expect(cell("A2")).toBe("");
    expect(cell("B2")).toBe("");
    expect(cell("F1")).toBe(15);
  });

  test("snapshot restore keeps incremental invalidation working", () => {
    setCellContent("A1", 1);
    setCellContent("B1", "=A1+1");
    setCellContent("C1", "=10+1");

    expect(cell("B1")).toBe(2);
    expect(cell("C1")).toBe(11);

    const snapshot = engine.serializeEngine();

    const restoredEngine = FormulaEngine.buildEmpty();
    restoredEngine.addWorkbook({ workbookName: workbookName });
    restoredEngine.addSheet({ workbookName, sheetName });
    restoredEngine.resetToSerializedEngine(snapshot);
    engine = restoredEngine;

    const dependentNode = cellNode("B1");
    const unrelatedNode = cellNode("C1");

    expect(dependentNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);

    setCellContent("A1", 2);

    expect(dependentNode.resolved).toBe(false);
    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("B1")).toBe(3);
    expect(cell("C1")).toBe(11);
  });

  test("sheet rename invalidates only formulas that explicitly depend on that sheet", () => {
    engine.addSheet({ workbookName, sheetName: "Sheet2" });
    engine.addSheet({ workbookName, sheetName: "Sheet3" });

    setCellContent("A1", 5, "TestSheet");
    setCellContent("A1", "=TestSheet!A1+1", "Sheet2");
    setCellContent("A1", "=10+1", "Sheet3");

    expect(cell("A1", "Sheet2")).toBe(6);
    expect(cell("A1", "Sheet3")).toBe(11);

    const dependentNode = cellNode("A1", "Sheet2");
    const unrelatedNode = cellNode("A1", "Sheet3");

    expect(dependentNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);

    engine.renameSheet({
      workbookName,
      sheetName: "TestSheet",
      newSheetName: "RenamedSheet",
    });

    expect(dependentNode.resolved).toBe(false);
    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("A1", "Sheet2")).toBe(6);
    expect(cell("A1", "Sheet3")).toBe(11);
  });

  test("sheet creation invalidates only formulas that explicitly depend on the new sheet", () => {
    setCellContent("A1", 1);
    setCellContent("B1", `=[${workbookName}]${sheetName}!A1+1`);
    setCellContent("C1", "=Sheet2!A1");
    setCellContent("D1", "=10+1");

    expect(cell("B1")).toBe(2);
    expect(typeof cell("C1")).toBe("string");
    expect(cell("D1")).toBe(11);

    const explicitWorkbookNode = cellNode("B1");
    const newSheetDependentNode = cellNode("C1");
    const unrelatedNode = cellNode("D1");

    expect(explicitWorkbookNode.resolved).toBe(true);
    expect(newSheetDependentNode.resolved).toBe(true);
    expect(unrelatedNode.resolved).toBe(true);

    engine.addSheet({ workbookName, sheetName: "Sheet2" });

    expect(explicitWorkbookNode.resolved).toBe(true);
    expect(newSheetDependentNode.resolved).toBe(false);
    expect(unrelatedNode.resolved).toBe(true);
    expect(cell("B1")).toBe(2);
    expect(cell("C1")).toBe("");
    expect(cell("D1")).toBe(11);
  });

  describe("spill anchors evaluated before their inputs change", () => {
    // A spilling formula stores its real result on a spill-meta node, and copies the
    // top-left value onto the anchor's own cell-value node. Reading the anchor first
    // resolves the anchor without recording the AST dependency that the spill-meta
    // node collects, so editing an input never invalidates the anchor and every
    // subsequent read is served from a stale cache.
    //
    // Reading a spill member first happens to build the edge, which is why the same
    // formula recalculates correctly depending only on read order.

    test("editing an input invalidates a spill anchor that was already read", () => {
      setCellContent("A1", 2);
      setCellContent("C1", "=SEQUENCE(3,1,A1)");

      // Read the anchor, not a member.
      expect(cell("C1")).toBe(2);

      const anchorNode = cellNode("C1");
      const metaNode = spillMetaNode("C1");
      expect(anchorNode.resolved).toBe(true);
      expect(metaNode.resolved).toBe(true);

      setCellContent("A1", 7);

      // The anchor depends on A1 through the spilling formula, so both nodes must be
      // invalidated by the edit.
      expect(metaNode.resolved).toBe(false);
      expect(anchorNode.resolved).toBe(false);

      expect(cell("C1")).toBe(7);
      expect(cell("C2")).toBe(8);
      expect(cell("C3")).toBe(9);
    });

    test("a spilling formula recalculates the same way regardless of read order", () => {
      setCellContent("A1", 2);
      setCellContent("C1", "=SEQUENCE(3,1,A1)");
      // Read a member first, which is the ordering that already works.
      expect(cell("C2")).toBe(3);
      setCellContent("A1", 7);
      const memberFirst = [cell("C1"), cell("C2"), cell("C3")];

      engine = FormulaEngine.buildEmpty();
      engine.addWorkbook({ workbookName: workbookName });
      engine.addSheet({ workbookName, sheetName });

      setCellContent("A1", 2);
      setCellContent("C1", "=SEQUENCE(3,1,A1)");
      // Read the anchor first.
      expect(cell("C1")).toBe(2);
      setCellContent("A1", 7);
      const anchorFirst = [cell("C1"), cell("C2"), cell("C3")];

      expect(anchorFirst).toEqual(memberFirst);
      expect(anchorFirst).toEqual([7, 8, 9]);
    });

    test("repeated edits refresh a spill anchor that was already read", () => {
      setCellContent("A1", 2);
      setCellContent("C1", "=SEQUENCE(3,1,A1)");

      expect(cell("C1")).toBe(2);

      // Neither a second edit nor an unrelated write recovers the cached value, so
      // once an anchor goes stale it stays stale for the lifetime of the engine.
      setCellContent("A1", 7);
      expect(cell("C1")).toBe(7);

      setCellContent("A1", 8);
      expect(cell("C1")).toBe(8);

      setCellContent("Z9", "unrelated");
      expect(cell("C1")).toBe(8);
    });

    test("editing an input invalidates a UNIQUE anchor that was already read", () => {
      setCellContent("A1", "a");
      setCellContent("A2", "b");
      setCellContent("C1", "=UNIQUE(A1:A2)");

      expect(cell("C1")).toBe("a");

      setCellContent("A1", "z");

      expect(cell("C1")).toBe("z");
      expect(cell("C2")).toBe("b");
    });
  });
});
