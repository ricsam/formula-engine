import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

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
      const node = engine._dependencyManager.getEvaluatedNode(cellToDepKey(cell));
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

    test("Should handle resolved correctly when evaluating a spilled cell first", () => {
      const evaluate = () => {
        for (const cell of evalOrder("D11")) {
          evalCell(cell);
        }
        engine._dependencyManager.markResolvedNodes(cellToDepKey("D11"));
      };

      // Setup: Reproduce the exact scenario from the spec
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 2],
          ["A3", 3],
          ["B1", "=D11 * 0.5"],
          ["B2", 8],
          ["B3", 7],
          ["C1", "=A1:A3 * B1:B3"],
          ["D10", "=A1:A2 * (B2 + A1)"],
        ])
      );

      //#region step 1
      evaluate();

      expect(directDeps("D11")).toMatchInlineSnapshot(`
        {
          "deps": [],
          "discardedFrontierDependencies": {},
          "frontierDependencies": {
            "D11:D11": [
              "D10",
              "C1",
            ],
          },
        }
      `);
      expect(evalOrder("D11")).toMatchInlineSnapshot(`
        [
          "D10",
          "C1",
          "D11",
        ]
      `);
      expect(dependencyTree("D11")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "D11:D11": [
                "D10",
                "C1",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "D11:D11": [
                "D10",
                "C1",
              ],
            },
          },
          "directDepsUpdated": true,
          "frontierDependencies": {
            "D11:D11": [
              {
                "directDepsUpdated": false,
                "key": "D10",
                "resolved": false,
              },
              {
                "directDepsUpdated": false,
                "key": "C1",
                "resolved": false,
              },
            ],
          },
          "key": "D11",
          "resolved": false,
        }
      `);
      //#endregion

      //#region step 2
      evaluate();
      expect(evalOrder("D11")).toMatchInlineSnapshot(`
        [
          "B2",
          "A1",
          "D10",
          "C1",
          "D11",
        ]
      `);
      expect(dependencyTree("D11")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "D11:D11": [
                "D10",
                "C1",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "D11:D11": [
                "D10",
                "C1",
              ],
            },
          },
          "directDepsUpdated": false,
          "frontierDependencies": {
            "D11:D11": [
              {
                "deps": [
                  {
                    "directDepsUpdated": false,
                    "key": "B2",
                    "resolved": false,
                  },
                  {
                    "directDepsUpdated": false,
                    "key": "A1",
                    "resolved": false,
                  },
                ],
                "directDepsUpdated": true,
                "key": "D10",
                "resolved": false,
              },
              {
                "deps": [
                  {
                    "directDepsUpdated": false,
                    "key": "A1",
                    "resolved": false,
                  },
                ],
                "directDepsUpdated": true,
                "key": "C1",
                "resolved": false,
              },
            ],
          },
          "key": "D11",
          "resolved": false,
        }
      `);
      //#endregion
      const evaluationManager = engine._evaluationManager;
      console.log(evaluationManager.getEvaluatedNodes());
    });
  });
});
