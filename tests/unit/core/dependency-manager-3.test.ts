import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import {
  FormulaError,
  type DependencyNode,
  type SerializedCellValue,
} from "src/core/types";
import { parseCellReference } from "src/core/utils";
import { dependencyNodeToKey } from "src/core/utils/dependency-node-key";

describe("DependencyManager", () => {
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

  const address = (ref: string) => ({
    sheetName,
    workbookName,
    ...parseCellReference(ref),
  });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    const cellToDepKey = (cell: string) =>
      `cell:TestWorkbook:TestSheet:${cell}`;

    const evalCell = (cellRef: string) => {
      return engine._evaluationManager.evaluateDependencyNode(
        cellToDepKey(cellRef)
      );
    };

    const directDeps = (cell: string) => {
      const deps: string[] = [];
      const frontierDependencies: Record<string, string[]> = {};
      const discardedFrontierDependencies: Record<string, string[]> = {};
      const node = engine._storeManager.getEvaluatedNode(cellToDepKey(cell));
      node?.deps?.forEach((dep) => deps.push(dep.split(":")[3]!));

      // Handle Map-based frontier dependencies
      if (node?.frontierDependencies) {
        for (const [range, rangeDeps] of node.frontierDependencies.entries()) {
          frontierDependencies[range] = [];
          rangeDeps.forEach((dep) =>
            frontierDependencies[range]!.push(dep.split(":")[3]!)
          );
        }
      }

      // Handle Map-based discarded frontier dependencies
      if (node?.discardedFrontierDependencies) {
        for (const [
          range,
          rangeDeps,
        ] of node.discardedFrontierDependencies.entries()) {
          discardedFrontierDependencies[range] = [];
          rangeDeps.forEach((dep) =>
            discardedFrontierDependencies[range]!.push(dep.split(":")[3]!)
          );
        }
      }

      return {
        deps,
        frontierDependencies,
        discardedFrontierDependencies,
      };
    };

    const evalOrder = (cell: string) => {
      return engine._dependencyManager
        .buildEvaluationOrder(cellToDepKey(cell))
        .evaluationOrder.map((depKey) => depKey.split(":")[3]!);
    };

    const dependencyTree = (cell: string) => {
      return engine._dependencyManager.getDependencyTree(cellToDepKey(cell));
    };

    test("Should handle resolved correctly", () => {
      const evaluate = () => {
        for (const cell of evalOrder("A1")) {
          evalCell(cell);
        }
        // evalCell("A1");
      };

      // Setup: Reproduce the exact scenario from the spec
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(C3:D4)"],
          ["B2", "=I12:K14"],
          ["H10", "=SEQUENCE(10, 10)"],
        ])
      );

      evaluate();

      expect(directDeps("A1")).toMatchInlineSnapshot(`
        {
          "deps": [],
          "discardedFrontierDependencies": {},
          "frontierDependencies": {
            "C3:D4": [
              "B2",
              "A1",
            ],
          },
        }
      `);

      expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "B2",
          "A1",
        ]
      `);
      expect(dependencyTree("A1")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "C3:D4": [
                "B2",
                "A1",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "C3:D4": [
                "B2",
                "A1",
              ],
            },
          },
          "frontierDependencies": {
            "C3:D4": [
              {
                "key": "B2",
                "resolved": false,
              },
              {
                "key": "A1",
                "resolved": false,
                "self": true,
              },
            ],
          },
          "key": "A1",
          "resolved": false,
        }
      `);
    });
  });
});
