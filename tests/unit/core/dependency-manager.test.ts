import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import type { DependencyNode } from "../../../src/core/managers/dependency-node";
import { type SerializedCellValue } from "../../../src/core/types";
import { parseCellReference } from "../../../src/core/utils";
import { AstEvaluationNode } from "../../../src/evaluator/dependency-nodes/ast-evaluation-node";
import { EmptyCellEvaluationNode } from "../../../src/evaluator/dependency-nodes/empty-cell-evaluation-node";
import { SpillMetaNode } from "../../../src/evaluator/dependency-nodes/spill-meta-node";
import {
  getContextDependencyKey,
  NO_TABLE_CONTEXT_NAME,
} from "../../../src/evaluator/evaluation-context";
import { RangeEvaluationNode } from "../../../src/evaluator/range-evaluation-node";
import { parseFormula } from "../../../src/parser/parser";

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

  const cellToDepNode = (cell: string): DependencyNode => {
    if (cell.includes("TestWorkbook")) {
      if (cell.startsWith("cell-value:")) {
        return engine._dependencyManager.getCellValueNode(cell);
      }

      if (cell.startsWith("empty-cell:")) {
        return engine._dependencyManager.getEmptyCellNode(cell);
      }

      if (cell.startsWith("spill-meta:")) {
        return engine._dependencyManager.getSpillMetaNode(cell);
      }

      if (cell.startsWith("range:")) {
        return engine._dependencyManager.getRangeNode(cell);
      }
    }
    if (
      cell.startsWith("range:") ||
      (!cell.startsWith("cell:") &&
        !cell.startsWith("empty-cell:") &&
        !cell.startsWith("spill-meta:") &&
        cell.includes(":"))
    ) {
      return engine._dependencyManager.getRangeNode(
        `range:TestWorkbook:TestSheet:${cell.replace(/^[^:]+:/, "")}`
      );
    }

    const node = engine._dependencyManager.getCellValueOrEmptyCellNode(
      `cell-value:TestWorkbook:TestSheet:${cell.replace(/^[^:]+:/, "")}`
    );

    return node;
  };

  const depToKey = (dep: DependencyNode) => {
    if (dep instanceof RangeEvaluationNode) {
      return dep.key.split(":")[3]! + ":" + dep.key.split(":")[4]!;
    }
    if (dep instanceof EmptyCellEvaluationNode) {
      return "empty:" + dep.key.split(":")[3]!;
    }
    if (dep instanceof SpillMetaNode) {
      return "spill-meta:" + dep.key.split(":")[3]!;
    }
    if (dep instanceof AstEvaluationNode) {
      // AST keys are like "ast:SUM(A1:A3)" - return as-is
      return dep.key;
    }
    return "cell-value:" + dep.key.split(":")[3]!;
  };

  const directDeps = (cell: string) => {
    const node = cellToDepNode(cell);

    const result: any = {
      deps: [...node.getDependencies()].map(depToKey),
      frontierDependencies:
        node instanceof RangeEvaluationNode ||
        node instanceof EmptyCellEvaluationNode
          ? [...node.getFrontierDependencies()].map(depToKey)
          : [],
      discardedFrontierDependencies:
        node instanceof RangeEvaluationNode ||
        node instanceof EmptyCellEvaluationNode
          ? [...node.getDiscardedFrontierDependencies()].map(depToKey)
          : [],
    };

    if (!result.deps.length) {
      delete result.deps;
    }
    if (!result.frontierDependencies.length) {
      delete result.frontierDependencies;
    }
    if (!result.discardedFrontierDependencies.length) {
      delete result.discardedFrontierDependencies;
    }
    return result;
  };

  const evalOrder = (cell: string) => {
    const key = `cell-value:TestWorkbook:TestSheet:${cell.replace(
      /^[^:]+:/,
      ""
    )}`;
    const node = engine._dependencyManager.getCellValueNode(key);
    const o = Array.from(
      engine._dependencyManager.buildEvaluationOrder(node).evaluationOrder
    );
    const order = o.map(depToKey);
    return order;
  };

  const dependencyTree = (cell: string) => {
    const key = `cell-value:TestWorkbook:TestSheet:${cell.replace(
      /^[^:]+:/,
      ""
    )}`;
    const node = engine._dependencyManager.getCellValueNode(key);
    return engine._dependencyManager.getDependencyTree(node);
  };
  const markAsResolved = (cell: string) => {
    const key = `cell-value:TestWorkbook:TestSheet:${cell.replace(
      /^[^:]+:/,
      ""
    )}`;
    const node = engine._dependencyManager.getCellValueNode(key);

    // After marking resolved, need to iterate until hash stabilizes
    // to discover dependencies that were blocked before
    let prevHash: string | undefined;
    let maxIterations = 10;
    let iteration = 0;

    while (iteration < maxIterations) {
      engine._dependencyManager.markResolvedNodes(node);

      const evalOrder = engine._dependencyManager.buildEvaluationOrder(node);

      if (prevHash && prevHash === evalOrder.hash) {
        break;
      }

      prevHash = evalOrder.hash;

      // Evaluate any newly discovered dependencies
      for (const c of evalOrder.evaluationOrder) {
        if (!c.resolved) {
          engine._evaluationManager.evaluateDependencyNode(c);
        }
      }

      iteration++;
    }
  };
  const evaluate = (cell: string) => {
    const key = `cell-value:TestWorkbook:TestSheet:${cell.replace(
      /^[^:]+:/,
      ""
    )}`;
    const node = engine._dependencyManager.getCellValueNode(key);
    const evalOrder = engine._dependencyManager.buildEvaluationOrder(node);
    for (const c of evalOrder.evaluationOrder) {
      engine._evaluationManager.evaluateDependencyNode(c);
    }
  };
  const generalEvaluate = evaluate;
  const generalMarkAsResolved = markAsResolved;

  test("Should reproduce DEPENDENCY_RESOLUTION_SPEC.md SUM example - tracking eval order at each step", () => {
    const evaluate = () => {
      generalEvaluate("C1");
    };
    const markAsResolved = () => {
      generalMarkAsResolved("C1");
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

    //#region step 1
    evaluate();
    markAsResolved();

    expect(directDeps("C1")).toMatchInlineSnapshot(`
      {
        "deps": [
          "ast:SUM(A1:A3*B1:B3)",
        ],
      }
    `);

    expect(evalOrder("C1")).toMatchInlineSnapshot(`
      [
        "spill-meta:D10",
        "empty:D11",
        "ast:D11",
        "ast:D11*0.5",
        "spill-meta:B1",
        "empty:B2",
        "ast:B2",
        "ast:B2+A1",
        "ast:A1:A2*(B2+A1)",
        "cell-value:B1",
        "B1:B3",
        "ast:SUM(A1:A3*B1:B3)",
        "cell-value:C1",
      ]
    `);
    expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "deps": [
            {
              "canResolve": true,
              "deps": [
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:A1:A3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:B1:B3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "ast:A1:A3*B1:B3",
                  "resolved": true,
                  "resultType": "spilled-values",
                  "type": "cell",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A1:A3",
                  "resolved": true,
                  "resultType": "range",
                  "type": "range",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "deps": [
                                    {
                                      "canResolve": true,
                                      "deps": [
                                        {
                                          "canResolve": true,
                                          "deps": [
                                            {
                                              "canResolve": true,
                                              "directDepsUpdated": false,
                                              "key": "ast:A1:A2",
                                              "resolved": true,
                                              "resultType": "spilled-values",
                                              "type": "cell",
                                            },
                                            {
                                              "canResolve": true,
                                              "deps": [
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": false,
                                                      "deps": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "deps": [
                                                                {
                                                                  "canResolve": true,
                                                                  "deps": [
                                                                    {
                                                                      "canResolve": true,
                                                                      "directDepsUpdated": false,
                                                                      "key": "F1",
                                                                      "resolved": true,
                                                                      "resultType": "value",
                                                                      "type": "cell",
                                                                    },
                                                                  ],
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:F1",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                                {
                                                                  "canResolve": true,
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:2",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                              ],
                                                              "directDepsUpdated": false,
                                                              "key": "ast:SEQUENCE(F1,2)",
                                                              "resolved": true,
                                                              "resultType": "spilled-values",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "A2",
                                                          "resolved": true,
                                                          "resultType": "spilled-values",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "directDepsUpdated": false,
                                                      "frontierDependencies": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "circular": true,
                                                              "directDepsUpdated": false,
                                                              "key": "ast:D11*0.5",
                                                              "resolved": false,
                                                              "resultType": "value",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "B1",
                                                          "resolved": false,
                                                          "resultType": "does-not-spill",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "key": "B2",
                                                      "resolved": false,
                                                      "resultType": "value",
                                                      "type": "empty",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:B2",
                                                  "resolved": false,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": true,
                                                      "directDepsUpdated": false,
                                                      "key": "A1",
                                                      "resolved": true,
                                                      "resultType": "value",
                                                      "type": "cell",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:A1",
                                                  "resolved": true,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                              ],
                                              "directDepsUpdated": false,
                                              "key": "ast:B2+A1",
                                              "resolved": false,
                                              "resultType": "value",
                                              "type": "cell",
                                            },
                                          ],
                                          "directDepsUpdated": false,
                                          "key": "ast:A1:A2*(B2+A1)",
                                          "resolved": false,
                                          "resultType": "spilled-values",
                                          "type": "cell",
                                        },
                                      ],
                                      "directDepsUpdated": false,
                                      "key": "D10",
                                      "resolved": false,
                                      "resultType": "spilled-values",
                                      "type": "cell",
                                    },
                                    {
                                      "canResolve": true,
                                      "directDepsUpdated": false,
                                      "key": "A2",
                                      "resolved": true,
                                      "resultType": "value",
                                      "type": "cell",
                                    },
                                  ],
                                  "directDepsUpdated": false,
                                  "key": "D11",
                                  "resolved": false,
                                  "resultType": "value",
                                  "type": "empty",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:D11",
                              "resolved": false,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:0.5",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:D11*0.5",
                          "resolved": false,
                          "resultType": "value",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "B1",
                      "resolved": false,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "B3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "F1",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:F1",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:2",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:SEQUENCE(F1,2)",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "B1:B3",
                  "resolved": false,
                  "resultType": "range",
                  "type": "range",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:SUM(A1:A3*B1:B3)",
              "resolved": false,
              "resultType": "value",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "C1",
          "resolved": false,
          "resultType": "value",
          "type": "cell",
        }
      `);

    //#endregion
    //#region step 2

    evaluate();
    markAsResolved();

    expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "deps": [
            {
              "canResolve": true,
              "deps": [
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:A1:A3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:B1:B3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "ast:A1:A3*B1:B3",
                  "resolved": true,
                  "resultType": "spilled-values",
                  "type": "cell",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A1:A3",
                  "resolved": true,
                  "resultType": "range",
                  "type": "range",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "deps": [
                                    {
                                      "canResolve": true,
                                      "deps": [
                                        {
                                          "canResolve": true,
                                          "deps": [
                                            {
                                              "canResolve": true,
                                              "directDepsUpdated": false,
                                              "key": "ast:A1:A2",
                                              "resolved": true,
                                              "resultType": "spilled-values",
                                              "type": "cell",
                                            },
                                            {
                                              "canResolve": true,
                                              "deps": [
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": false,
                                                      "deps": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "deps": [
                                                                {
                                                                  "canResolve": true,
                                                                  "deps": [
                                                                    {
                                                                      "canResolve": true,
                                                                      "directDepsUpdated": false,
                                                                      "key": "F1",
                                                                      "resolved": true,
                                                                      "resultType": "value",
                                                                      "type": "cell",
                                                                    },
                                                                  ],
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:F1",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                                {
                                                                  "canResolve": true,
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:2",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                              ],
                                                              "directDepsUpdated": false,
                                                              "key": "ast:SEQUENCE(F1,2)",
                                                              "resolved": true,
                                                              "resultType": "spilled-values",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "A2",
                                                          "resolved": true,
                                                          "resultType": "spilled-values",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "directDepsUpdated": false,
                                                      "frontierDependencies": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "circular": true,
                                                              "directDepsUpdated": false,
                                                              "key": "ast:D11*0.5",
                                                              "resolved": false,
                                                              "resultType": "value",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "B1",
                                                          "resolved": false,
                                                          "resultType": "does-not-spill",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "key": "B2",
                                                      "resolved": false,
                                                      "resultType": "value",
                                                      "type": "empty",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:B2",
                                                  "resolved": false,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": true,
                                                      "directDepsUpdated": false,
                                                      "key": "A1",
                                                      "resolved": true,
                                                      "resultType": "value",
                                                      "type": "cell",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:A1",
                                                  "resolved": true,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                              ],
                                              "directDepsUpdated": false,
                                              "key": "ast:B2+A1",
                                              "resolved": false,
                                              "resultType": "value",
                                              "type": "cell",
                                            },
                                          ],
                                          "directDepsUpdated": false,
                                          "key": "ast:A1:A2*(B2+A1)",
                                          "resolved": false,
                                          "resultType": "spilled-values",
                                          "type": "cell",
                                        },
                                      ],
                                      "directDepsUpdated": false,
                                      "key": "D10",
                                      "resolved": false,
                                      "resultType": "spilled-values",
                                      "type": "cell",
                                    },
                                    {
                                      "canResolve": true,
                                      "directDepsUpdated": false,
                                      "key": "A2",
                                      "resolved": true,
                                      "resultType": "value",
                                      "type": "cell",
                                    },
                                  ],
                                  "directDepsUpdated": false,
                                  "key": "D11",
                                  "resolved": false,
                                  "resultType": "value",
                                  "type": "empty",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:D11",
                              "resolved": false,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:0.5",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:D11*0.5",
                          "resolved": false,
                          "resultType": "value",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "B1",
                      "resolved": false,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "B3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "F1",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:F1",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:2",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:SEQUENCE(F1,2)",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "B1:B3",
                  "resolved": false,
                  "resultType": "range",
                  "type": "range",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:SUM(A1:A3*B1:B3)",
              "resolved": false,
              "resultType": "value",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "C1",
          "resolved": false,
          "resultType": "value",
          "type": "cell",
        }
      `);

    expect(directDeps("A2")).toMatchInlineSnapshot(`{}`);

    expect(directDeps("B1")).toMatchInlineSnapshot(`
      {
        "deps": [
          "ast:D11*0.5",
        ],
      }
    `);

    expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "spill-meta:D10",
          "empty:D11",
          "ast:D11",
          "ast:D11*0.5",
          "spill-meta:B1",
          "empty:B2",
          "ast:B2",
          "ast:B2+A1",
          "ast:A1:A2*(B2+A1)",
          "cell-value:B1",
          "B1:B3",
          "ast:SUM(A1:A3*B1:B3)",
          "cell-value:C1",
        ]
      `);

    //#endregion
    //#region step 3

    evaluate();
    markAsResolved();

    expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "deps": [
            {
              "canResolve": true,
              "deps": [
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:A1:A3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:B1:B3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "ast:A1:A3*B1:B3",
                  "resolved": true,
                  "resultType": "spilled-values",
                  "type": "cell",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A1:A3",
                  "resolved": true,
                  "resultType": "range",
                  "type": "range",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "deps": [
                                    {
                                      "canResolve": true,
                                      "deps": [
                                        {
                                          "canResolve": true,
                                          "deps": [
                                            {
                                              "canResolve": true,
                                              "directDepsUpdated": false,
                                              "key": "ast:A1:A2",
                                              "resolved": true,
                                              "resultType": "spilled-values",
                                              "type": "cell",
                                            },
                                            {
                                              "canResolve": true,
                                              "deps": [
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": false,
                                                      "deps": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "deps": [
                                                                {
                                                                  "canResolve": true,
                                                                  "deps": [
                                                                    {
                                                                      "canResolve": true,
                                                                      "directDepsUpdated": false,
                                                                      "key": "F1",
                                                                      "resolved": true,
                                                                      "resultType": "value",
                                                                      "type": "cell",
                                                                    },
                                                                  ],
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:F1",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                                {
                                                                  "canResolve": true,
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:2",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                              ],
                                                              "directDepsUpdated": false,
                                                              "key": "ast:SEQUENCE(F1,2)",
                                                              "resolved": true,
                                                              "resultType": "spilled-values",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "A2",
                                                          "resolved": true,
                                                          "resultType": "spilled-values",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "directDepsUpdated": false,
                                                      "frontierDependencies": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "circular": true,
                                                              "directDepsUpdated": false,
                                                              "key": "ast:D11*0.5",
                                                              "resolved": false,
                                                              "resultType": "value",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "B1",
                                                          "resolved": false,
                                                          "resultType": "does-not-spill",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "key": "B2",
                                                      "resolved": false,
                                                      "resultType": "value",
                                                      "type": "empty",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:B2",
                                                  "resolved": false,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": true,
                                                      "directDepsUpdated": false,
                                                      "key": "A1",
                                                      "resolved": true,
                                                      "resultType": "value",
                                                      "type": "cell",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:A1",
                                                  "resolved": true,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                              ],
                                              "directDepsUpdated": false,
                                              "key": "ast:B2+A1",
                                              "resolved": false,
                                              "resultType": "value",
                                              "type": "cell",
                                            },
                                          ],
                                          "directDepsUpdated": false,
                                          "key": "ast:A1:A2*(B2+A1)",
                                          "resolved": false,
                                          "resultType": "spilled-values",
                                          "type": "cell",
                                        },
                                      ],
                                      "directDepsUpdated": false,
                                      "key": "D10",
                                      "resolved": false,
                                      "resultType": "spilled-values",
                                      "type": "cell",
                                    },
                                    {
                                      "canResolve": true,
                                      "directDepsUpdated": false,
                                      "key": "A2",
                                      "resolved": true,
                                      "resultType": "value",
                                      "type": "cell",
                                    },
                                  ],
                                  "directDepsUpdated": false,
                                  "key": "D11",
                                  "resolved": false,
                                  "resultType": "value",
                                  "type": "empty",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:D11",
                              "resolved": false,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:0.5",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:D11*0.5",
                          "resolved": false,
                          "resultType": "value",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "B1",
                      "resolved": false,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "B3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "F1",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:F1",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:2",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:SEQUENCE(F1,2)",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "B1:B3",
                  "resolved": false,
                  "resultType": "range",
                  "type": "range",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:SUM(A1:A3*B1:B3)",
              "resolved": false,
              "resultType": "value",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "C1",
          "resolved": false,
          "resultType": "value",
          "type": "cell",
        }
      `);

    expect(directDeps("D11")).toMatchInlineSnapshot(`
      {
        "deps": [
          "spill-meta:D10",
          "cell-value:A2",
        ],
      }
    `);

    expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "spill-meta:D10",
          "empty:D11",
          "ast:D11",
          "ast:D11*0.5",
          "spill-meta:B1",
          "empty:B2",
          "ast:B2",
          "ast:B2+A1",
          "ast:A1:A2*(B2+A1)",
          "cell-value:B1",
          "B1:B3",
          "ast:SUM(A1:A3*B1:B3)",
          "cell-value:C1",
        ]
      `);
    //#endregion

    //#region step 4
    // now that
    evaluate();
    markAsResolved();

    // A1 * B1 resolves fine now so SUM doesn't short circuit on that and continues to A2 * B2
    expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "spill-meta:D10",
          "empty:D11",
          "ast:D11",
          "ast:D11*0.5",
          "spill-meta:B1",
          "empty:B2",
          "ast:B2",
          "ast:B2+A1",
          "ast:A1:A2*(B2+A1)",
          "cell-value:B1",
          "B1:B3",
          "ast:SUM(A1:A3*B1:B3)",
          "cell-value:C1",
        ]
      `);

    expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "deps": [
            {
              "canResolve": true,
              "deps": [
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:A1:A3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:B1:B3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "ast:A1:A3*B1:B3",
                  "resolved": true,
                  "resultType": "spilled-values",
                  "type": "cell",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A1:A3",
                  "resolved": true,
                  "resultType": "range",
                  "type": "range",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "deps": [
                                    {
                                      "canResolve": true,
                                      "deps": [
                                        {
                                          "canResolve": true,
                                          "deps": [
                                            {
                                              "canResolve": true,
                                              "directDepsUpdated": false,
                                              "key": "ast:A1:A2",
                                              "resolved": true,
                                              "resultType": "spilled-values",
                                              "type": "cell",
                                            },
                                            {
                                              "canResolve": true,
                                              "deps": [
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": false,
                                                      "deps": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "deps": [
                                                                {
                                                                  "canResolve": true,
                                                                  "deps": [
                                                                    {
                                                                      "canResolve": true,
                                                                      "directDepsUpdated": false,
                                                                      "key": "F1",
                                                                      "resolved": true,
                                                                      "resultType": "value",
                                                                      "type": "cell",
                                                                    },
                                                                  ],
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:F1",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                                {
                                                                  "canResolve": true,
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:2",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                              ],
                                                              "directDepsUpdated": false,
                                                              "key": "ast:SEQUENCE(F1,2)",
                                                              "resolved": true,
                                                              "resultType": "spilled-values",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "A2",
                                                          "resolved": true,
                                                          "resultType": "spilled-values",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "directDepsUpdated": false,
                                                      "frontierDependencies": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "circular": true,
                                                              "directDepsUpdated": false,
                                                              "key": "ast:D11*0.5",
                                                              "resolved": false,
                                                              "resultType": "value",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "B1",
                                                          "resolved": false,
                                                          "resultType": "does-not-spill",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "key": "B2",
                                                      "resolved": false,
                                                      "resultType": "value",
                                                      "type": "empty",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:B2",
                                                  "resolved": false,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": true,
                                                      "directDepsUpdated": false,
                                                      "key": "A1",
                                                      "resolved": true,
                                                      "resultType": "value",
                                                      "type": "cell",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:A1",
                                                  "resolved": true,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                              ],
                                              "directDepsUpdated": false,
                                              "key": "ast:B2+A1",
                                              "resolved": false,
                                              "resultType": "value",
                                              "type": "cell",
                                            },
                                          ],
                                          "directDepsUpdated": false,
                                          "key": "ast:A1:A2*(B2+A1)",
                                          "resolved": false,
                                          "resultType": "spilled-values",
                                          "type": "cell",
                                        },
                                      ],
                                      "directDepsUpdated": false,
                                      "key": "D10",
                                      "resolved": false,
                                      "resultType": "spilled-values",
                                      "type": "cell",
                                    },
                                    {
                                      "canResolve": true,
                                      "directDepsUpdated": false,
                                      "key": "A2",
                                      "resolved": true,
                                      "resultType": "value",
                                      "type": "cell",
                                    },
                                  ],
                                  "directDepsUpdated": false,
                                  "key": "D11",
                                  "resolved": false,
                                  "resultType": "value",
                                  "type": "empty",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:D11",
                              "resolved": false,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:0.5",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:D11*0.5",
                          "resolved": false,
                          "resultType": "value",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "B1",
                      "resolved": false,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "B3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "F1",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:F1",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:2",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:SEQUENCE(F1,2)",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "B1:B3",
                  "resolved": false,
                  "resultType": "range",
                  "type": "range",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:SUM(A1:A3*B1:B3)",
              "resolved": false,
              "resultType": "value",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "C1",
          "resolved": false,
          "resultType": "value",
          "type": "cell",
        }
      `);
    //#endregion
    //#region step 5

    evaluate();
    markAsResolved();
    expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "spill-meta:D10",
          "empty:D11",
          "ast:D11",
          "ast:D11*0.5",
          "spill-meta:B1",
          "empty:B2",
          "ast:B2",
          "ast:B2+A1",
          "ast:A1:A2*(B2+A1)",
          "cell-value:B1",
          "B1:B3",
          "ast:SUM(A1:A3*B1:B3)",
          "cell-value:C1",
        ]
      `);

    expect(dependencyTree("C1")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "deps": [
            {
              "canResolve": true,
              "deps": [
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:A1:A3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "ast:B1:B3",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "ast:A1:A3*B1:B3",
                  "resolved": true,
                  "resultType": "spilled-values",
                  "type": "cell",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A1",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "A3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "A1:A3",
                  "resolved": true,
                  "resultType": "range",
                  "type": "range",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "deps": [
                                    {
                                      "canResolve": true,
                                      "deps": [
                                        {
                                          "canResolve": true,
                                          "deps": [
                                            {
                                              "canResolve": true,
                                              "directDepsUpdated": false,
                                              "key": "ast:A1:A2",
                                              "resolved": true,
                                              "resultType": "spilled-values",
                                              "type": "cell",
                                            },
                                            {
                                              "canResolve": true,
                                              "deps": [
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": false,
                                                      "deps": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "deps": [
                                                                {
                                                                  "canResolve": true,
                                                                  "deps": [
                                                                    {
                                                                      "canResolve": true,
                                                                      "directDepsUpdated": false,
                                                                      "key": "F1",
                                                                      "resolved": true,
                                                                      "resultType": "value",
                                                                      "type": "cell",
                                                                    },
                                                                  ],
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:F1",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                                {
                                                                  "canResolve": true,
                                                                  "directDepsUpdated": false,
                                                                  "key": "ast:2",
                                                                  "resolved": true,
                                                                  "resultType": "value",
                                                                  "type": "cell",
                                                                },
                                                              ],
                                                              "directDepsUpdated": false,
                                                              "key": "ast:SEQUENCE(F1,2)",
                                                              "resolved": true,
                                                              "resultType": "spilled-values",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "A2",
                                                          "resolved": true,
                                                          "resultType": "spilled-values",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "directDepsUpdated": false,
                                                      "frontierDependencies": [
                                                        {
                                                          "canResolve": true,
                                                          "deps": [
                                                            {
                                                              "canResolve": true,
                                                              "circular": true,
                                                              "directDepsUpdated": false,
                                                              "key": "ast:D11*0.5",
                                                              "resolved": false,
                                                              "resultType": "value",
                                                              "type": "cell",
                                                            },
                                                          ],
                                                          "directDepsUpdated": false,
                                                          "key": "B1",
                                                          "resolved": false,
                                                          "resultType": "does-not-spill",
                                                          "type": "cell",
                                                        },
                                                      ],
                                                      "key": "B2",
                                                      "resolved": false,
                                                      "resultType": "value",
                                                      "type": "empty",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:B2",
                                                  "resolved": false,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                                {
                                                  "canResolve": true,
                                                  "deps": [
                                                    {
                                                      "canResolve": true,
                                                      "directDepsUpdated": false,
                                                      "key": "A1",
                                                      "resolved": true,
                                                      "resultType": "value",
                                                      "type": "cell",
                                                    },
                                                  ],
                                                  "directDepsUpdated": false,
                                                  "key": "ast:A1",
                                                  "resolved": true,
                                                  "resultType": "value",
                                                  "type": "cell",
                                                },
                                              ],
                                              "directDepsUpdated": false,
                                              "key": "ast:B2+A1",
                                              "resolved": false,
                                              "resultType": "value",
                                              "type": "cell",
                                            },
                                          ],
                                          "directDepsUpdated": false,
                                          "key": "ast:A1:A2*(B2+A1)",
                                          "resolved": false,
                                          "resultType": "spilled-values",
                                          "type": "cell",
                                        },
                                      ],
                                      "directDepsUpdated": false,
                                      "key": "D10",
                                      "resolved": false,
                                      "resultType": "spilled-values",
                                      "type": "cell",
                                    },
                                    {
                                      "canResolve": true,
                                      "directDepsUpdated": false,
                                      "key": "A2",
                                      "resolved": true,
                                      "resultType": "value",
                                      "type": "cell",
                                    },
                                  ],
                                  "directDepsUpdated": false,
                                  "key": "D11",
                                  "resolved": false,
                                  "resultType": "value",
                                  "type": "empty",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:D11",
                              "resolved": false,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:0.5",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:D11*0.5",
                          "resolved": false,
                          "resultType": "value",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "B1",
                      "resolved": false,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "directDepsUpdated": false,
                      "key": "B3",
                      "resolved": true,
                      "resultType": "value",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "F1",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:F1",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                            {
                              "canResolve": true,
                              "directDepsUpdated": false,
                              "key": "ast:2",
                              "resolved": true,
                              "resultType": "value",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "ast:SEQUENCE(F1,2)",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "A2",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "B1:B3",
                  "resolved": false,
                  "resultType": "range",
                  "type": "range",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:SUM(A1:A3*B1:B3)",
              "resolved": false,
              "resultType": "value",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "C1",
          "resolved": false,
          "resultType": "value",
          "type": "cell",
        }
      `);

    expect(directDeps("D11")).toMatchInlineSnapshot(`
        {
          "deps": [
            "spill-meta:D10",
            "cell-value:A2",
          ],
        }
      `);

    expect(evalOrder("C1")).toMatchInlineSnapshot(`
        [
          "spill-meta:D10",
          "empty:D11",
          "ast:D11",
          "ast:D11*0.5",
          "spill-meta:B1",
          "empty:B2",
          "ast:B2",
          "ast:B2+A1",
          "ast:A1:A2*(B2+A1)",
          "cell-value:B1",
          "B1:B3",
          "ast:SUM(A1:A3*B1:B3)",
          "cell-value:C1",
        ]
      `);

    //#endregion
    //#region step 6

    evaluate();
    markAsResolved();

    expect(cell("C1", true)).toBe(24.5);

    //#endregion
  });

  test("prefers the most specific cached AST match over wildcard table entries", () => {
    const ast = parseFormula("ValueTable[@Value]");
    const broadContext = {
      workbookName,
      rowIndex: 1,
    };
    const specificContext = {
      workbookName,
      sheetName,
      rowIndex: 1,
      colIndex: 5,
    };

    const broadNode = new AstEvaluationNode(ast, broadContext);
    const specificNode = new AstEvaluationNode(ast, specificContext);

    engine._dependencyManager.asts.set("ast:ValueTable[@Value]", {
      entries: new Map([
        [
          getContextDependencyKey(broadContext),
          {
            evalNode: broadNode,
            contextDependency: broadContext,
          },
        ],
        [
          getContextDependencyKey(specificContext),
          {
            evalNode: specificNode,
            contextDependency: specificContext,
          },
        ],
      ]),
    });

    expect(
      engine._dependencyManager.getAstNode(ast, specificContext)
    ).toBe(specificNode);
  });

  test("keeps implicit current-row AST cache entries distinct across no-table and table contexts", () => {
    const ast = parseFormula("[@Value]");
    const noTableContext = {
      workbookName,
      sheetName,
      tableName: NO_TABLE_CONTEXT_NAME,
      rowIndex: 1,
      colIndex: 5,
    };
    const tableContext = {
      workbookName,
      sheetName,
      tableName: "ValueTable",
      rowIndex: 1,
      colIndex: 5,
    };

    const staleNode = new AstEvaluationNode(ast, noTableContext);
    engine._dependencyManager.asts.set("ast:[@Value]", {
      entries: new Map([
        [
          getContextDependencyKey(noTableContext),
          {
            evalNode: staleNode,
            contextDependency: noTableContext,
          },
        ],
      ]),
    });

    const node = engine._dependencyManager.getAstNode(ast, tableContext);

    expect(node).toBeInstanceOf(AstEvaluationNode);
    expect(node).not.toBe(staleNode);
    expect(engine._dependencyManager.asts.get(node.key)?.entries.size).toBe(2);
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
    markAsResolved("A1");
    expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "cell-value:B1",
          "ast:B1",
          "cell-value:A1",
          "ast:A1",
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
          "cell-value:A1",
        ]
      `);
    //#region step 1
    evaluate("A1");
    markAsResolved("A1");

    expect(dependencyTree("A1")).toMatchInlineSnapshot(`
        {
          "canResolve": false,
          "deps": [
            {
              "canResolve": false,
              "deps": [
                {
                  "canResolve": false,
                  "circular": true,
                  "directDepsUpdated": false,
                  "key": "A1",
                  "resolved": false,
                  "resultType": "awaiting-evaluation",
                  "type": "cell",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:A1",
              "resolved": false,
              "resultType": "awaiting-evaluation",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "A1",
          "resolved": false,
          "resultType": "awaiting-evaluation",
          "type": "cell",
        }
      `);
    expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "cell-value:A1",
          "ast:A1",
        ]
      `);
    //#endregion
    //#region step 2
    evaluate("A1");
    markAsResolved("A1");

    expect(dependencyTree("A1")).toMatchInlineSnapshot(`
        {
          "canResolve": false,
          "deps": [
            {
              "canResolve": false,
              "deps": [
                {
                  "canResolve": false,
                  "circular": true,
                  "directDepsUpdated": false,
                  "key": "A1",
                  "resolved": false,
                  "resultType": "awaiting-evaluation",
                  "type": "cell",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:A1",
              "resolved": false,
              "resultType": "awaiting-evaluation",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "A1",
          "resolved": false,
          "resultType": "awaiting-evaluation",
          "type": "cell",
        }
      `);
    //#endregion
  });
  test("Should handle resolved correctly", () => {
    const evaluate = () => {
      generalEvaluate("A1");
    };
    const markAsResolved = () => {
      generalMarkAsResolved("A1");
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
    markAsResolved();

    expect(directDeps("A1")).toMatchInlineSnapshot(`
      {
        "deps": [
          "ast:SUM(C3:D4)",
        ],
      }
    `);

    expect(evalOrder("A1")).toMatchInlineSnapshot(`
        [
          "cell-value:A1",
        ]
      `);
    expect(dependencyTree("A1")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "deps": [
            {
              "canResolve": true,
              "deps": [
                {
                  "canResolve": true,
                  "directDepsUpdated": false,
                  "key": "ast:C3:D4",
                  "resolved": true,
                  "resultType": "spilled-values",
                  "type": "cell",
                },
                {
                  "canResolve": true,
                  "deps": [
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "ast:10",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:SEQUENCE(10,10)",
                              "resolved": true,
                              "resultType": "spilled-values",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "H10",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "J13:J14",
                      "resolved": true,
                      "resultType": "range",
                      "type": "range",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "directDepsUpdated": false,
                          "key": "ast:I12:K14",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "B2",
                      "resolved": true,
                      "resultType": "spilled-values",
                      "type": "cell",
                    },
                    {
                      "canResolve": true,
                      "deps": [
                        {
                          "canResolve": true,
                          "deps": [
                            {
                              "canResolve": true,
                              "deps": [
                                {
                                  "canResolve": true,
                                  "directDepsUpdated": false,
                                  "key": "ast:10",
                                  "resolved": true,
                                  "resultType": "value",
                                  "type": "cell",
                                },
                              ],
                              "directDepsUpdated": false,
                              "key": "ast:SEQUENCE(10,10)",
                              "resolved": true,
                              "resultType": "spilled-values",
                              "type": "cell",
                            },
                          ],
                          "directDepsUpdated": false,
                          "key": "H10",
                          "resolved": true,
                          "resultType": "spilled-values",
                          "type": "cell",
                        },
                      ],
                      "directDepsUpdated": false,
                      "key": "K13:K14",
                      "resolved": true,
                      "resultType": "range",
                      "type": "range",
                    },
                  ],
                  "directDepsUpdated": false,
                  "key": "C3:D4",
                  "resolved": true,
                  "resultType": "range",
                  "type": "range",
                },
              ],
              "directDepsUpdated": false,
              "key": "ast:SUM(C3:D4)",
              "resolved": true,
              "resultType": "value",
              "type": "cell",
            },
          ],
          "directDepsUpdated": false,
          "key": "A1",
          "resolved": true,
          "resultType": "value",
          "type": "cell",
        }
      `);
  });

  test("should discard non-spilling frontier candidates before a range resolves", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["AA1", "=SUM(Q1:Q)"],
        ["H1", "=1"],
        ["H2", "=2"],
        ["H3", "=3"],
      ])
    );

    generalEvaluate("AA1");
    generalMarkAsResolved("AA1");

    expect(cell("AA1")).toBe(0);

    const rangeNode = engine._dependencyManager.getRangeNode(
      `range:${workbookName}:${sheetName}:Q1:Q`
    );

    expect(rangeNode.resolved).toBe(true);
    expect(directDeps("range:Q1:Q")).toMatchInlineSnapshot(`
      {
        "discardedFrontierDependencies": [
          "spill-meta:H1",
          "spill-meta:H2",
          "spill-meta:H3",
        ],
      }
    `);
    expect(evalOrder("AA1")).toMatchInlineSnapshot(`
      [
        "cell-value:AA1",
      ]
    `);
  });

  test("Should handle resolved correctly when evaluating a spilled cell first", () => {
    const evaluate = () => {
      generalEvaluate("D11");
    };
    const markAsResolved = () => {
      generalMarkAsResolved("D11");
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
    markAsResolved();

    expect(directDeps("D11")).toMatchInlineSnapshot(`
        {
          "frontierDependencies": [
            "spill-meta:D10",
          ],
        }
      `);
    expect(evalOrder("D11")).toMatchInlineSnapshot(`
        [
          "cell-value:D11",
        ]
      `);
    expect(dependencyTree("D11")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "directDepsUpdated": false,
          "key": "D11",
          "resolved": true,
          "resultType": "value",
          "type": "cell",
        }
      `);
    //#endregion

    //#region step 2
    evaluate();
    markAsResolved();
    expect(evalOrder("D11")).toMatchInlineSnapshot(`
        [
          "cell-value:D11",
        ]
      `);
    expect(dependencyTree("D11")).toMatchInlineSnapshot(`
        {
          "canResolve": true,
          "directDepsUpdated": false,
          "key": "D11",
          "resolved": true,
          "resultType": "value",
          "type": "cell",
        }
      `);
    //#endregion
  });

  test("debug an issue", () => {
    const evaluate = () => {
      generalEvaluate("D1");
    };
    const markAsResolved = () => {
      generalMarkAsResolved("D1");
    };

    // Setup: Reproduce the exact scenario from the spec
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=C1:C2"],
        ["B1", "=SUM(A1:A2)"],
        ["B2", "5"],
        ["D1", "=SUM(B1:B2)"],
        ["C1", "=SEQUENCE(2)"],
      ])
    );

    //#region step 1
    evaluate();
    markAsResolved();
    expect(evalOrder("D1")).toMatchInlineSnapshot(`
        [
          "cell-value:D1",
        ]
      `);
    expect(dependencyTree("D1")).toMatchInlineSnapshot(`
      {
        "canResolve": true,
        "deps": [
          {
            "canResolve": true,
            "deps": [
              {
                "canResolve": true,
                "directDepsUpdated": false,
                "key": "ast:B1:B2",
                "resolved": true,
                "resultType": "spilled-values",
                "type": "cell",
              },
              {
                "canResolve": true,
                "deps": [
                  {
                    "canResolve": true,
                    "deps": [
                      {
                        "canResolve": true,
                        "deps": [
                          {
                            "canResolve": true,
                            "directDepsUpdated": false,
                            "key": "ast:A1:A2",
                            "resolved": true,
                            "resultType": "spilled-values",
                            "type": "cell",
                          },
                          {
                            "canResolve": true,
                            "deps": [
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": true,
                                    "resultType": "value",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "value",
                                "type": "cell",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "deps": [
                                      {
                                        "canResolve": true,
                                        "deps": [
                                          {
                                            "canResolve": true,
                                            "directDepsUpdated": false,
                                            "key": "ast:2",
                                            "resolved": true,
                                            "resultType": "value",
                                            "type": "cell",
                                          },
                                        ],
                                        "directDepsUpdated": false,
                                        "key": "ast:SEQUENCE(2)",
                                        "resolved": true,
                                        "resultType": "spilled-values",
                                        "type": "cell",
                                      },
                                    ],
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": false,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "C2",
                                "resolved": true,
                                "resultType": "value",
                                "type": "empty",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "ast:C1:C2",
                                    "resolved": true,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "spilled-values",
                                "type": "cell",
                              },
                            ],
                            "directDepsUpdated": false,
                            "key": "A1:A2",
                            "resolved": true,
                            "resultType": "range",
                            "type": "range",
                          },
                        ],
                        "directDepsUpdated": false,
                        "key": "ast:SUM(A1:A2)",
                        "resolved": true,
                        "resultType": "value",
                        "type": "cell",
                      },
                    ],
                    "directDepsUpdated": false,
                    "key": "B1",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                  {
                    "canResolve": true,
                    "directDepsUpdated": false,
                    "key": "B2",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                ],
                "directDepsUpdated": false,
                "key": "B1:B2",
                "resolved": true,
                "resultType": "range",
                "type": "range",
              },
            ],
            "directDepsUpdated": false,
            "key": "ast:SUM(B1:B2)",
            "resolved": true,
            "resultType": "value",
            "type": "cell",
          },
        ],
        "directDepsUpdated": false,
        "key": "D1",
        "resolved": true,
        "resultType": "value",
        "type": "cell",
      }
    `);
    //#endregion

    //#region step 2
    evaluate();
    markAsResolved();
    expect(evalOrder("D1")).toMatchInlineSnapshot(`
      [
        "cell-value:D1",
      ]
    `);
    expect(dependencyTree("D1")).toMatchInlineSnapshot(`
      {
        "canResolve": true,
        "deps": [
          {
            "canResolve": true,
            "deps": [
              {
                "canResolve": true,
                "directDepsUpdated": false,
                "key": "ast:B1:B2",
                "resolved": true,
                "resultType": "spilled-values",
                "type": "cell",
              },
              {
                "canResolve": true,
                "deps": [
                  {
                    "canResolve": true,
                    "deps": [
                      {
                        "canResolve": true,
                        "deps": [
                          {
                            "canResolve": true,
                            "directDepsUpdated": false,
                            "key": "ast:A1:A2",
                            "resolved": true,
                            "resultType": "spilled-values",
                            "type": "cell",
                          },
                          {
                            "canResolve": true,
                            "deps": [
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": true,
                                    "resultType": "value",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "value",
                                "type": "cell",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "deps": [
                                      {
                                        "canResolve": true,
                                        "deps": [
                                          {
                                            "canResolve": true,
                                            "directDepsUpdated": false,
                                            "key": "ast:2",
                                            "resolved": true,
                                            "resultType": "value",
                                            "type": "cell",
                                          },
                                        ],
                                        "directDepsUpdated": false,
                                        "key": "ast:SEQUENCE(2)",
                                        "resolved": true,
                                        "resultType": "spilled-values",
                                        "type": "cell",
                                      },
                                    ],
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": false,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "C2",
                                "resolved": true,
                                "resultType": "value",
                                "type": "empty",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "ast:C1:C2",
                                    "resolved": true,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "spilled-values",
                                "type": "cell",
                              },
                            ],
                            "directDepsUpdated": false,
                            "key": "A1:A2",
                            "resolved": true,
                            "resultType": "range",
                            "type": "range",
                          },
                        ],
                        "directDepsUpdated": false,
                        "key": "ast:SUM(A1:A2)",
                        "resolved": true,
                        "resultType": "value",
                        "type": "cell",
                      },
                    ],
                    "directDepsUpdated": false,
                    "key": "B1",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                  {
                    "canResolve": true,
                    "directDepsUpdated": false,
                    "key": "B2",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                ],
                "directDepsUpdated": false,
                "key": "B1:B2",
                "resolved": true,
                "resultType": "range",
                "type": "range",
              },
            ],
            "directDepsUpdated": false,
            "key": "ast:SUM(B1:B2)",
            "resolved": true,
            "resultType": "value",
            "type": "cell",
          },
        ],
        "directDepsUpdated": false,
        "key": "D1",
        "resolved": true,
        "resultType": "value",
        "type": "cell",
      }
    `);
    //#endregion

    //#region step 3
    evaluate();
    markAsResolved();
    expect(evalOrder("D1")).toMatchInlineSnapshot(`
      [
        "cell-value:D1",
      ]
    `);
    expect(dependencyTree("D1")).toMatchInlineSnapshot(`
      {
        "canResolve": true,
        "deps": [
          {
            "canResolve": true,
            "deps": [
              {
                "canResolve": true,
                "directDepsUpdated": false,
                "key": "ast:B1:B2",
                "resolved": true,
                "resultType": "spilled-values",
                "type": "cell",
              },
              {
                "canResolve": true,
                "deps": [
                  {
                    "canResolve": true,
                    "deps": [
                      {
                        "canResolve": true,
                        "deps": [
                          {
                            "canResolve": true,
                            "directDepsUpdated": false,
                            "key": "ast:A1:A2",
                            "resolved": true,
                            "resultType": "spilled-values",
                            "type": "cell",
                          },
                          {
                            "canResolve": true,
                            "deps": [
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": true,
                                    "resultType": "value",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "value",
                                "type": "cell",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "deps": [
                                      {
                                        "canResolve": true,
                                        "deps": [
                                          {
                                            "canResolve": true,
                                            "directDepsUpdated": false,
                                            "key": "ast:2",
                                            "resolved": true,
                                            "resultType": "value",
                                            "type": "cell",
                                          },
                                        ],
                                        "directDepsUpdated": false,
                                        "key": "ast:SEQUENCE(2)",
                                        "resolved": true,
                                        "resultType": "spilled-values",
                                        "type": "cell",
                                      },
                                    ],
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": false,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "C2",
                                "resolved": true,
                                "resultType": "value",
                                "type": "empty",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "ast:C1:C2",
                                    "resolved": true,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "spilled-values",
                                "type": "cell",
                              },
                            ],
                            "directDepsUpdated": false,
                            "key": "A1:A2",
                            "resolved": true,
                            "resultType": "range",
                            "type": "range",
                          },
                        ],
                        "directDepsUpdated": false,
                        "key": "ast:SUM(A1:A2)",
                        "resolved": true,
                        "resultType": "value",
                        "type": "cell",
                      },
                    ],
                    "directDepsUpdated": false,
                    "key": "B1",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                  {
                    "canResolve": true,
                    "directDepsUpdated": false,
                    "key": "B2",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                ],
                "directDepsUpdated": false,
                "key": "B1:B2",
                "resolved": true,
                "resultType": "range",
                "type": "range",
              },
            ],
            "directDepsUpdated": false,
            "key": "ast:SUM(B1:B2)",
            "resolved": true,
            "resultType": "value",
            "type": "cell",
          },
        ],
        "directDepsUpdated": false,
        "key": "D1",
        "resolved": true,
        "resultType": "value",
        "type": "cell",
      }
    `);
    //#endregion

    //#region step 4
    evaluate();
    // HERE
    expect(dependencyTree("D1")).toMatchInlineSnapshot(`
      {
        "canResolve": true,
        "deps": [
          {
            "canResolve": true,
            "deps": [
              {
                "canResolve": true,
                "directDepsUpdated": false,
                "key": "ast:B1:B2",
                "resolved": true,
                "resultType": "spilled-values",
                "type": "cell",
              },
              {
                "canResolve": true,
                "deps": [
                  {
                    "canResolve": true,
                    "deps": [
                      {
                        "canResolve": true,
                        "deps": [
                          {
                            "canResolve": true,
                            "directDepsUpdated": false,
                            "key": "ast:A1:A2",
                            "resolved": true,
                            "resultType": "spilled-values",
                            "type": "cell",
                          },
                          {
                            "canResolve": true,
                            "deps": [
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": true,
                                    "resultType": "value",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "value",
                                "type": "cell",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "deps": [
                                      {
                                        "canResolve": true,
                                        "deps": [
                                          {
                                            "canResolve": true,
                                            "directDepsUpdated": false,
                                            "key": "ast:2",
                                            "resolved": true,
                                            "resultType": "value",
                                            "type": "cell",
                                          },
                                        ],
                                        "directDepsUpdated": false,
                                        "key": "ast:SEQUENCE(2)",
                                        "resolved": true,
                                        "resultType": "spilled-values",
                                        "type": "cell",
                                      },
                                    ],
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": false,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "C2",
                                "resolved": true,
                                "resultType": "value",
                                "type": "empty",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "ast:C1:C2",
                                    "resolved": true,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "spilled-values",
                                "type": "cell",
                              },
                            ],
                            "directDepsUpdated": false,
                            "key": "A1:A2",
                            "resolved": true,
                            "resultType": "range",
                            "type": "range",
                          },
                        ],
                        "directDepsUpdated": false,
                        "key": "ast:SUM(A1:A2)",
                        "resolved": true,
                        "resultType": "value",
                        "type": "cell",
                      },
                    ],
                    "directDepsUpdated": false,
                    "key": "B1",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                  {
                    "canResolve": true,
                    "directDepsUpdated": false,
                    "key": "B2",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                ],
                "directDepsUpdated": false,
                "key": "B1:B2",
                "resolved": true,
                "resultType": "range",
                "type": "range",
              },
            ],
            "directDepsUpdated": false,
            "key": "ast:SUM(B1:B2)",
            "resolved": true,
            "resultType": "value",
            "type": "cell",
          },
        ],
        "directDepsUpdated": false,
        "key": "D1",
        "resolved": true,
        "resultType": "value",
        "type": "cell",
      }
    `);
    evaluate(); // to discover new range deps
    evaluate(); // to discover new range deps
    expect(evalOrder("D1")).toMatchInlineSnapshot(`
      [
        "cell-value:D1",
      ]
    `);
    markAsResolved();
    expect(evalOrder("D1")).toMatchInlineSnapshot(`
      [
        "cell-value:D1",
      ]
    `);
    expect(dependencyTree("D1")).toMatchInlineSnapshot(`
      {
        "canResolve": true,
        "deps": [
          {
            "canResolve": true,
            "deps": [
              {
                "canResolve": true,
                "directDepsUpdated": false,
                "key": "ast:B1:B2",
                "resolved": true,
                "resultType": "spilled-values",
                "type": "cell",
              },
              {
                "canResolve": true,
                "deps": [
                  {
                    "canResolve": true,
                    "deps": [
                      {
                        "canResolve": true,
                        "deps": [
                          {
                            "canResolve": true,
                            "directDepsUpdated": false,
                            "key": "ast:A1:A2",
                            "resolved": true,
                            "resultType": "spilled-values",
                            "type": "cell",
                          },
                          {
                            "canResolve": true,
                            "deps": [
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": true,
                                    "resultType": "value",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "value",
                                "type": "cell",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "deps": [
                                      {
                                        "canResolve": true,
                                        "deps": [
                                          {
                                            "canResolve": true,
                                            "directDepsUpdated": false,
                                            "key": "ast:2",
                                            "resolved": true,
                                            "resultType": "value",
                                            "type": "cell",
                                          },
                                        ],
                                        "directDepsUpdated": false,
                                        "key": "ast:SEQUENCE(2)",
                                        "resolved": true,
                                        "resultType": "spilled-values",
                                        "type": "cell",
                                      },
                                    ],
                                    "directDepsUpdated": false,
                                    "key": "C1",
                                    "resolved": false,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "C2",
                                "resolved": true,
                                "resultType": "value",
                                "type": "empty",
                              },
                              {
                                "canResolve": true,
                                "deps": [
                                  {
                                    "canResolve": true,
                                    "directDepsUpdated": false,
                                    "key": "ast:C1:C2",
                                    "resolved": true,
                                    "resultType": "spilled-values",
                                    "type": "cell",
                                  },
                                ],
                                "directDepsUpdated": false,
                                "key": "A1",
                                "resolved": true,
                                "resultType": "spilled-values",
                                "type": "cell",
                              },
                            ],
                            "directDepsUpdated": false,
                            "key": "A1:A2",
                            "resolved": true,
                            "resultType": "range",
                            "type": "range",
                          },
                        ],
                        "directDepsUpdated": false,
                        "key": "ast:SUM(A1:A2)",
                        "resolved": true,
                        "resultType": "value",
                        "type": "cell",
                      },
                    ],
                    "directDepsUpdated": false,
                    "key": "B1",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                  {
                    "canResolve": true,
                    "directDepsUpdated": false,
                    "key": "B2",
                    "resolved": true,
                    "resultType": "value",
                    "type": "cell",
                  },
                ],
                "directDepsUpdated": false,
                "key": "B1:B2",
                "resolved": true,
                "resultType": "range",
                "type": "range",
              },
            ],
            "directDepsUpdated": false,
            "key": "ast:SUM(B1:B2)",
            "resolved": true,
            "resultType": "value",
            "type": "cell",
          },
        ],
        "directDepsUpdated": false,
        "key": "D1",
        "resolved": true,
        "resultType": "value",
        "type": "cell",
      }
    `);
    //#endregion
  });
});
