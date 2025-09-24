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

    test("Should reproduce DEPENDENCY_RESOLUTION_SPEC.md SUM example - tracking eval order at each step", () => {
      const evaluate = () => {
        for (const cell of evalOrder("C1")) {
          evalCell(cell);
        }
      };

      // Setup: Reproduce the exact scenario from the spec
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", "=SEQUENCE(F1, 2)"],
          ["A3", 3],
          ["B1", "=D11 * 0.5"],
          ["B2", ""],
          ["B3", 7],
          ["C1", "=SUM(A1:A3 * B1:B3)"],
          ["D10", "=A1:A2 * (B2 + A1)"], // Spills to D10:D11
          ["F1", 1],
        ])
      );

      evaluate();

      expect(directDeps("C1")).toEqual({
        deps: ["B1", "A1"],
        frontierDependencies: { "B1:B3": ["A2"] },
        discardedFrontierDependencies: {},
      });

      expect(evalOrder("C1")).toEqual(["B1", "A1", "A2", "C1"]);
      expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
          },
          "deps": [
            {
              "key": "B1",
              "resolved": false,
            },
            {
              "key": "A1",
              "resolved": false,
            },
          ],
          "frontierDependencies": {
            "B1:B3": [
              {
                "key": "A2",
                "resolved": false,
              },
            ],
          },
          "key": "C1",
          "resolved": false,
        }
      `);

      evaluate();

      expect(directDeps("A2")).toEqual({
        deps: ["F1"],
        frontierDependencies: {},
        discardedFrontierDependencies: {},
      });

      expect(directDeps("B1")).toEqual({
        deps: ["D11"],
        frontierDependencies: {},
        discardedFrontierDependencies: {},
      });

      expect(evalOrder("C1")).toEqual(["D11", "B1", "A1", "F1", "A2", "C1"]);

      expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
          },
          "deps": [
            {
              "deps": [
                {
                  "key": "D11",
                  "resolved": false,
                },
              ],
              "key": "B1",
              "resolved": false,
            },
            {
              "key": "A1",
              "resolved": true,
            },
          ],
          "frontierDependencies": {
            "B1:B3": [
              {
                "deps": [
                  {
                    "key": "F1",
                    "resolved": false,
                  },
                ],
                "key": "A2",
                "resolved": false,
              },
            ],
          },
          "key": "C1",
          "resolved": false,
        }
      `);

      evaluate();

      expect(directDeps("D11")).toEqual({
        deps: [],
        frontierDependencies: { "D11:D11": ["D10", "A2", "C1"] },
        discardedFrontierDependencies: {},
      });

      expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
          },
          "deps": [
            {
              "deps": [
                {
                  "_debug": {
                    "activeFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "A2",
                        "C1",
                      ],
                    },
                    "discardedFrontierDependencies": undefined,
                    "rawFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "A2",
                        "C1",
                      ],
                    },
                  },
                  "frontierDependencies": {
                    "D11:D11": [
                      {
                        "key": "D10",
                        "resolved": false,
                      },
                      {
                        "deps": [
                          {
                            "key": "F1",
                            "resolved": true,
                          },
                        ],
                        "key": "A2",
                        "resolved": true,
                      },
                      {
                        "key": "C1",
                        "resolved": false,
                        "self": true,
                      },
                    ],
                  },
                  "key": "D11",
                  "resolved": false,
                },
              ],
              "key": "B1",
              "resolved": false,
            },
            {
              "key": "A1",
              "resolved": true,
            },
            {
              "deps": [
                {
                  "key": "F1",
                  "resolved": true,
                },
              ],
              "key": "A2",
              "resolved": true,
            },
          ],
          "key": "C1",
          "resolved": false,
        }
      `);

      expect(evalOrder("C1")).toEqual([
        "D10",
        "F1",
        "A2",
        "D11",
        "B1",
        "A1",
        "C1",
      ]);

      // now that
      evaluate();

      // A1 * B1 resolves fine now so SUM doesn't short circuit on that and continues to A2 * B2
      expect(evalOrder("C1")).toEqual([
        "F1",
        "A2",
        "A1",
        "D10",
        "D11",
        "B1",
        "B3",
        "A3",
        "C1",
      ]);

      expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
          },
          "deps": [
            {
              "deps": [
                {
                  "_debug": {
                    "activeFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "A2",
                        "C1",
                      ],
                    },
                    "discardedFrontierDependencies": undefined,
                    "rawFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "A2",
                        "C1",
                      ],
                    },
                  },
                  "deps": [
                    {
                      "deps": [
                        {
                          "deps": [
                            {
                              "key": "F1",
                              "resolved": true,
                            },
                          ],
                          "key": "A2",
                          "resolved": true,
                        },
                        {
                          "key": "A1",
                          "resolved": true,
                        },
                      ],
                      "key": "D10",
                      "resolved": false,
                    },
                    {
                      "deps": [
                        {
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "key": "A2",
                      "resolved": true,
                    },
                  ],
                  "frontierDependencies": {
                    "D11:D11": [
                      {
                        "key": "C1",
                        "resolved": false,
                        "self": true,
                      },
                    ],
                  },
                  "key": "D11",
                  "resolved": false,
                },
                {
                  "deps": [
                    {
                      "deps": [
                        {
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "key": "A2",
                      "resolved": true,
                    },
                    {
                      "key": "A1",
                      "resolved": true,
                    },
                  ],
                  "key": "D10",
                  "resolved": false,
                },
                {
                  "deps": [
                    {
                      "key": "F1",
                      "resolved": true,
                    },
                  ],
                  "key": "A2",
                  "resolved": true,
                },
              ],
              "key": "B1",
              "resolved": false,
            },
            {
              "key": "A1",
              "resolved": true,
            },
            {
              "deps": [
                {
                  "key": "F1",
                  "resolved": true,
                },
              ],
              "key": "A2",
              "resolved": true,
            },
            {
              "key": "B3",
              "resolved": false,
            },
            {
              "key": "A3",
              "resolved": false,
            },
          ],
          "key": "C1",
          "resolved": false,
        }
      `);

      evaluate();

      expect(directDeps("D11")).toEqual({
        deps: ["D10", "A2"],
        discardedFrontierDependencies: {},
        frontierDependencies: {
          "D11:D11": ["D10", "A2", "C1"],
        },
      });

      // SUM can now continue to A3 * B3
      expect(evalOrder("C1")).toEqual([
        "F1",
        "A2",
        "A1",
        "D10",
        "D11",
        "B1",
        "B3",
        "A3",
        "C1",
      ]);

      evaluate();

      // C1 is now resolved
      expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "_debug": {
            "activeFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
            "discardedFrontierDependencies": undefined,
            "rawFrontierDependencies": {
              "B1:B3": [
                "A2",
              ],
            },
          },
          "deps": [
            {
              "deps": [
                {
                  "_debug": {
                    "activeFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "A2",
                        "C1",
                      ],
                    },
                    "discardedFrontierDependencies": undefined,
                    "rawFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "A2",
                        "C1",
                      ],
                    },
                  },
                  "deps": [
                    {
                      "deps": [
                        {
                          "deps": [
                            {
                              "key": "F1",
                              "resolved": true,
                            },
                          ],
                          "key": "A2",
                          "resolved": true,
                        },
                        {
                          "key": "A1",
                          "resolved": true,
                        },
                      ],
                      "key": "D10",
                      "resolved": true,
                    },
                    {
                      "deps": [
                        {
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "key": "A2",
                      "resolved": true,
                    },
                  ],
                  "frontierDependencies": {
                    "D11:D11": [
                      {
                        "key": "C1",
                        "resolved": true,
                        "self": true,
                      },
                    ],
                  },
                  "key": "D11",
                  "resolved": true,
                },
                {
                  "deps": [
                    {
                      "deps": [
                        {
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "key": "A2",
                      "resolved": true,
                    },
                    {
                      "key": "A1",
                      "resolved": true,
                    },
                  ],
                  "key": "D10",
                  "resolved": true,
                },
                {
                  "deps": [
                    {
                      "key": "F1",
                      "resolved": true,
                    },
                  ],
                  "key": "A2",
                  "resolved": true,
                },
              ],
              "key": "B1",
              "resolved": true,
            },
            {
              "key": "A1",
              "resolved": true,
            },
            {
              "deps": [
                {
                  "key": "F1",
                  "resolved": true,
                },
              ],
              "key": "A2",
              "resolved": true,
            },
            {
              "key": "B3",
              "resolved": true,
            },
            {
              "key": "A3",
              "resolved": true,
            },
          ],
          "key": "C1",
          "resolved": true,
        }
      `);
    });
  });
});
