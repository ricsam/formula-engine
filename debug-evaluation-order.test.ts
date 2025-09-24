import { test } from "bun:test";
import { FormulaEngine } from "./src/core/engine";
import { cellAddressToKey, parseCellReference } from "./src/core/utils";
import type { SerializedCellValue, SpreadsheetRange } from "./src/core/types";

test("debug evaluation order for multiplication", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "TestSheet";
  const sheetAddress = { workbookName, sheetName };

  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook(workbookName);
  engine.addSheet({ workbookName, sheetName });

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  // Set up the multiplication of ranges test scenario
  engine.setSheetContent(
    sheetAddress,
    new Map<string, SerializedCellValue>([
      ["A1", 1],
      ["A2", 2],
      ["A3", 3],
      ["B1", "=D11 * 0.5"], // This depends on D11 which is created by D10's spill
      ["B2", 8],
      ["B3", 7],
      ["C1", "=A1:A3 * B1:B3"], // This should evaluate to spilled values
      ["D10", "=A1:A2 * (B2 + A1)"], // This spills to D10:D11
    ])
  );

  console.log("\n=== Debug Evaluation Order ===");

  // Let's trace what happens when we evaluate D11
  console.log("\nStep 1: Evaluating D11");
  const d11Result = cell("D11", true);
  console.log("D11 result:", d11Result);

  // Check evaluated nodes after D11
  const storeManager = engine._storeManager;
  const depManager = engine._dependencyManager;

  console.log("\n=== Evaluated Nodes after D11 ===");
  const evaluatedAfterD11 = new Set<string>();
  for (const [key, node] of storeManager.getEvaluatedNodes()) {
    evaluatedAfterD11.add(key);
    console.log(
      `${key}: deps=${node.deps ? [...node.deps] : []}, frontierDeps=${node.frontierDependencies ? [...node.frontierDependencies] : []}`
    );
  }

  console.log("\nStep 2: Evaluating C1");

  // Get the node key for C1
  const c1NodeKey = cellAddressToKey({
    ...sheetAddress,
    colIndex: 2,
    rowIndex: 0,
    sheetName,
    workbookName,
  });

  // Build evaluation order for C1
  console.log("\nBuilding evaluation order for C1...");
  const evaluationPlan = depManager.buildEvaluationOrder(c1NodeKey);
  console.log("Has cycle:", evaluationPlan.hasCycle);
  console.log("Evaluation order:", evaluationPlan.evaluationOrder);

  // Now evaluate C1
  const c1Result = cell("C1", true);
  console.log("\nC1 result:", c1Result);

  // Check what new nodes were evaluated
  console.log("\n=== New Evaluated Nodes after C1 ===");
  for (const [key, node] of storeManager.getEvaluatedNodes()) {
    if (!evaluatedAfterD11.has(key)) {
      console.log(
        `${key}: deps=${node.deps ? [...node.deps] : []}, frontierDeps=${node.frontierDependencies ? [...node.frontierDependencies] : []}, discardedFrontierDeps=${node.discardedFrontierDependencies ? [...node.discardedFrontierDependencies] : []}`
      );
    }
  }

  // Check all evaluated nodes after C1
  console.log("\n=== All Evaluated Nodes after C1 ===");
  for (const [key, node] of storeManager.getEvaluatedNodes()) {
    const resultType = node.evaluationResult?.type;
    const resultValue =
      resultType === "value" ? node.evaluationResult?.result : resultType;
    console.log(
      `${key}: result=${resultValue}, deps=${node.deps ? [...node.deps].length : 0}, frontierDeps=${node.frontierDependencies ? [...node.frontierDependencies].length : 0}`
    );
  }

  // Check B1 specifically
  const b1NodeKey = cellAddressToKey({
    ...sheetAddress,
    colIndex: 1,
    rowIndex: 0,
    sheetName,
    workbookName,
  });
  const b1Node = storeManager.getEvaluatedNode(b1NodeKey);
  console.log("\n=== B1 Node Details ===");
  console.log("B1 evaluated:", !!b1Node);
  if (b1Node) {
    console.log("B1 result:", b1Node.evaluationResult);
    console.log("B1 deps:", b1Node.deps ? [...b1Node.deps] : []);
  }

  // Check the actual frontier candidates for C1:C3
  const workbookManager = engine._workbookManager;
  const c1c3Range: SpreadsheetRange = {
    start: { row: 0, col: 2 },
    end: {
      row: { type: "number", value: 2 },
      col: { type: "number", value: 2 },
    },
  };

  const frontierCandidates = workbookManager.getFrontierCandidates(
    c1c3Range,
    sheetAddress
  );
  console.log("\nFrontier candidates for C1:C3:");
  frontierCandidates.forEach((candidate) => {
    const ref = `${String.fromCharCode(65 + candidate.colIndex)}${candidate.rowIndex + 1}`;
    console.log(`  ${ref}`);
  });
});
