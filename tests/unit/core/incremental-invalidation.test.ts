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

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
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
    restoredEngine.addWorkbook(workbookName);
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
});
