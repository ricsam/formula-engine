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
      const node = engine._dependencyManager.getEvaluatedNode(
        cellToDepKey(cell)
      );
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
    const evaluate = (cell: string) => {
      for (const c of evalOrder(cell)) {
        evalCell(c);
      }
      engine._dependencyManager.markResolvedNodes(cellToDepKey(cell));
    };

    test("Should reproduce DEPENDENCY_RESOLUTION_SPEC.md SUM example - tracking eval order at each step", () => {
      const evaluate = () => {
        for (const cell of evalOrder("C1")) {
          evalCell(cell);
        }
        engine._dependencyManager.markResolvedNodes(cellToDepKey("C1"));
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
              "directDepsUpdated": false,
              "key": "B1",
              "resolved": false,
            },
            {
              "directDepsUpdated": false,
              "key": "A1",
              "resolved": false,
            },
          ],
          "directDepsUpdated": true,
          "frontierDependencies": {
            "B1:B3": [
              {
                "directDepsUpdated": false,
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

      expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "D11",
          "B1",
          "A1",
          "F1",
          "A2",
          "C1",
        ]
      `);

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
                  "directDepsUpdated": false,
                  "key": "D11",
                  "resolved": false,
                },
              ],
              "directDepsUpdated": true,
              "key": "B1",
              "resolved": false,
            },
            {
              "directDepsUpdated": false,
              "key": "A1",
              "resolved": false,
            },
          ],
          "directDepsUpdated": false,
          "frontierDependencies": {
            "B1:B3": [
              {
                "deps": [
                  {
                    "directDepsUpdated": false,
                    "key": "F1",
                    "resolved": false,
                  },
                ],
                "directDepsUpdated": true,
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
        frontierDependencies: { "D11:D11": ["D10", "C1", "A2"] },
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
                        "C1",
                        "A2",
                      ],
                    },
                    "discardedFrontierDependencies": undefined,
                    "rawFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "C1",
                        "A2",
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
                        "directDepsUpdated": true,
                        "key": "C1",
                        "resolved": false,
                        "self": true,
                      },
                      {
                        "deps": [
                          {
                            "directDepsUpdated": false,
                            "key": "F1",
                            "resolved": false,
                          },
                        ],
                        "directDepsUpdated": false,
                        "key": "A2",
                        "resolved": false,
                      },
                    ],
                  },
                  "key": "D11",
                  "resolved": false,
                },
              ],
              "directDepsUpdated": false,
              "key": "B1",
              "resolved": false,
            },
            {
              "directDepsUpdated": false,
              "key": "A1",
              "resolved": false,
            },
            {
              "deps": [
                {
                  "directDepsUpdated": false,
                  "key": "F1",
                  "resolved": false,
                },
              ],
              "directDepsUpdated": false,
              "key": "A2",
              "resolved": false,
            },
          ],
          "directDepsUpdated": true,
          "key": "C1",
          "resolved": false,
        }
      `);

      expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "D10",
          "F1",
          "A2",
          "D11",
          "B1",
          "A1",
          "C1",
        ]
      `);

      // now that
      evaluate();

      // A1 * B1 resolves fine now so SUM doesn't short circuit on that and continues to A2 * B2
      expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "D10",
          "D11",
          "B1",
          "B3",
          "A3",
          "C1",
        ]
      `);

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
                        "C1",
                        "A2",
                      ],
                    },
                    "discardedFrontierDependencies": undefined,
                    "rawFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "C1",
                        "A2",
                      ],
                    },
                  },
                  "deps": [
                    {
                      "deps": [
                        {
                          "deps": [
                            {
                              "directDepsUpdated": false,
                              "key": "F1",
                              "resolved": true,
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "A2",
                          "resolved": true,
                        },
                        {
                          "directDepsUpdated": false,
                          "key": "A1",
                          "resolved": true,
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
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                    },
                  ],
                  "directDepsUpdated": true,
                  "frontierDependencies": {
                    "D11:D11": [
                      {
                        "directDepsUpdated": true,
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
                          "directDepsUpdated": false,
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                    },
                    {
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
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
                      "key": "F1",
                      "resolved": true,
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A2",
                  "resolved": true,
                },
              ],
              "directDepsUpdated": true,
              "key": "B1",
              "resolved": false,
            },
            {
              "directDepsUpdated": false,
              "key": "A1",
              "resolved": true,
            },
            {
              "deps": [
                {
                  "directDepsUpdated": false,
                  "key": "F1",
                  "resolved": true,
                },
              ],
              "directDepsUpdated": false,
              "key": "A2",
              "resolved": true,
            },
            {
              "directDepsUpdated": false,
              "key": "B3",
              "resolved": false,
            },
            {
              "directDepsUpdated": false,
              "key": "A3",
              "resolved": false,
            },
          ],
          "directDepsUpdated": true,
          "key": "C1",
          "resolved": false,
        }
      `);

      evaluate();

      expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "C1",
        ]
      `);

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
                        "C1",
                        "A2",
                      ],
                    },
                    "discardedFrontierDependencies": undefined,
                    "rawFrontierDependencies": {
                      "D11:D11": [
                        "D10",
                        "C1",
                        "A2",
                      ],
                    },
                  },
                  "deps": [
                    {
                      "deps": [
                        {
                          "deps": [
                            {
                              "directDepsUpdated": false,
                              "key": "F1",
                              "resolved": true,
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "A2",
                          "resolved": true,
                        },
                        {
                          "directDepsUpdated": false,
                          "key": "A1",
                          "resolved": true,
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "D10",
                      "resolved": true,
                    },
                    {
                      "deps": [
                        {
                          "directDepsUpdated": false,
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                    },
                  ],
                  "directDepsUpdated": false,
                  "frontierDependencies": {
                    "D11:D11": [
                      {
                        "directDepsUpdated": false,
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
                          "directDepsUpdated": false,
                          "key": "F1",
                          "resolved": true,
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                    },
                    {
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "D10",
                  "resolved": true,
                },
                {
                  "deps": [
                    {
                      "directDepsUpdated": false,
                      "key": "F1",
                      "resolved": true,
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A2",
                  "resolved": true,
                },
              ],
              "directDepsUpdated": false,
              "key": "B1",
              "resolved": true,
            },
            {
              "directDepsUpdated": false,
              "key": "A1",
              "resolved": true,
            },
            {
              "deps": [
                {
                  "directDepsUpdated": false,
                  "key": "F1",
                  "resolved": true,
                },
              ],
              "directDepsUpdated": false,
              "key": "A2",
              "resolved": true,
            },
            {
              "directDepsUpdated": false,
              "key": "B3",
              "resolved": true,
            },
            {
              "directDepsUpdated": false,
              "key": "A3",
              "resolved": true,
            },
          ],
          "directDepsUpdated": false,
          "key": "C1",
          "resolved": true,
        }
      `);

      expect(directDeps("D11")).toEqual({
        deps: ["D10", "A2"],
        discardedFrontierDependencies: {},
        frontierDependencies: {
          "D11:D11": ["D10", "C1", "A2"],
        },
      });

      expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "C1",
        ]
      `);
    });

    test("should detect cycles", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=B1"],
          ["B1", "=A1"],
        ])
      );
      evaluate("A1");
      expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "B1",
          "A1",
        ]
      `);
    });

    test("should handle self-referencing cell", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", "=A1"]])
      );
      expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "A1",
        ]
      `);
      evaluate("A1");
      expect(dependencyTree("A1")).toMatchInlineSnapshot(`
        {
          "deps": [
            {
              "directDepsUpdated": true,
              "key": "A1",
              "resolved": false,
              "self": true,
            },
          ],
          "directDepsUpdated": true,
          "key": "A1",
          "resolved": false,
        }
      `);
      expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "A1",
        ]
      `)
      evaluate("A1");
      expect(dependencyTree("A1")).toMatchInlineSnapshot(`
        {
          "deps": [
            {
              "directDepsUpdated": false,
              "key": "A1",
              "resolved": true,
              "self": true,
            },
          ],
          "directDepsUpdated": false,
          "key": "A1",
          "resolved": true,
        }
      `);
    });
  });
});
