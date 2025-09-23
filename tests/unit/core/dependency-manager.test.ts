import { test, expect, describe, beforeEach } from "bun:test";
import { DependencyManager } from "../../../src/core/managers/dependency-manager";
import { StoreManager } from "../../../src/core/managers/store-manager";
import { WorkbookManager } from "../../../src/core/managers/workbook-manager";
import { NamedExpressionManager } from "../../../src/core/managers/named-expression-manager";
import type { EvaluatedDependencyNode } from "../../../src/core/types";

describe("DependencyManager", () => {
  let dependencyManager: DependencyManager;
  let storeManager: StoreManager;
  let workbookManager: WorkbookManager;

  beforeEach(() => {
    storeManager = new StoreManager();
    workbookManager = new WorkbookManager();
    dependencyManager = new DependencyManager(storeManager, workbookManager);
  });

  describe("Basic dependency resolution", () => {
    test("should handle simple linear dependencies", () => {
      // Set up A1 -> B1 -> C1
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(["B1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["C1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("A1");

      expect(result.hasCycle).toBe(false);
      expect(result.evaluationOrder).toEqual(["C1", "B1", "A1"]);
    });

    test("should handle multiple dependencies", () => {
      // Set up A1 -> [B1, B2], B1 -> C1, B2 -> C2
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(["B1", "B2"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["C1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("B2", {
        deps: new Set(["C2"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 4 },
        },
      });
      storeManager.setEvaluatedNode("C2", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 5 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("A1");

      expect(result.hasCycle).toBe(false);
      expect(result.evaluationOrder).toContain("C1");
      expect(result.evaluationOrder).toContain("C2");
      expect(result.evaluationOrder).toContain("B1");
      expect(result.evaluationOrder).toContain("B2");

      // C1 and C2 should come before B1 and B2
      const c1Index = result.evaluationOrder.indexOf("C1");
      const b1Index = result.evaluationOrder.indexOf("B1");
      expect(c1Index).toBeLessThan(b1Index);

      const c2Index = result.evaluationOrder.indexOf("C2");
      const b2Index = result.evaluationOrder.indexOf("B2");
      expect(c2Index).toBeLessThan(b2Index);
    });
  });

  describe("Cycle detection", () => {
    test("should detect simple cycle", () => {
      // Set up A1 -> B1 -> C1 -> A1
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(["B1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["C1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(["A1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("A1");

      expect(result.hasCycle).toBe(true);
      expect(result.cycleNodes).toBeDefined();
      expect(result.cycleNodes?.has("A1")).toBe(true);
      expect(result.cycleNodes?.has("B1")).toBe(true);
      expect(result.cycleNodes?.has("C1")).toBe(true);
    });

    test("should detect self-reference", () => {
      // Set up A1 -> A1
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(["A1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("A1");

      expect(result.hasCycle).toBe(true);
      expect(result.cycleNodes?.has("A1")).toBe(true);
    });
  });

  describe("Frontier dependency handling", () => {
    test("should handle frontier dependencies without cycles", () => {
      // Set up: C1 has frontier dependency on B1, B1 depends on A1
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(),
        frontierDependencies: new Map([["C1:C1", new Set(["B1"])]]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["A1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("C1");

      expect(result.hasCycle).toBe(false);
      expect(result.evaluationOrder).toContain("A1");
      expect(result.evaluationOrder).toContain("B1");

      // A1 should be evaluated before B1
      const a1Index = result.evaluationOrder.indexOf("A1");
      const b1Index = result.evaluationOrder.indexOf("B1");
      expect(a1Index).toBeLessThan(b1Index);
    });

    test("should handle frontier dependency that depends on target cell", () => {
      // This is the key test case:
      // C1 has frontier dependency on B1
      // B1 depends on D11 (which is created by D10's spill)
      // D10 depends on C1
      // This looks like a cycle but isn't because B1 is a frontier dependency

      storeManager.setEvaluatedNode("C1", {
        deps: new Set(["A1"]),
        frontierDependencies: new Map([["C1:C1", new Set(["B1"])]]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["D11"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("D10", {
        deps: new Set(["C1"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });
      storeManager.setEvaluatedNode("D11", {
        deps: new Set(["D10"]), // D11 is created from D10's spill
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 4 },
        },
      });
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 5 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("C1");

      // This should NOT be a cycle
      expect(result.hasCycle).toBe(false);
      expect(result.evaluationOrder).toContain("A1");
      expect(result.evaluationOrder).toContain("B1");
      expect(result.evaluationOrder).toContain("D10");
      expect(result.evaluationOrder).toContain("D11");
    });

    test("should handle discarded frontier dependencies", () => {
      // Set up: C1 has frontier dependency on B1, but B1 is discarded
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(["A1"]),
        frontierDependencies: new Map([["C1:C1", new Set(["B1"])]]),
        discardedFrontierDependencies: new Map([["C1:C1", new Set(["B1"])]]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("A1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("C1");

      expect(result.hasCycle).toBe(false);
      // B1 should not be in the evaluation order since it's discarded
      expect(result.evaluationOrder).not.toContain("B1");
      expect(result.evaluationOrder).toContain("A1");
    });
  });

  describe("Complex scenarios", () => {
    test("should handle the multiplication test scenario", () => {
      // This mimics the failing multiplication test:
      // C1 = A1:A3 * B1:B3 (spills)
      // B1 = D11 * 0.5
      // D10 = A1:A2 * (B2 + A1) (spills to D11)

      // In the initial state, before evaluation:
      // - C1 has not been evaluated yet
      // - B1 depends on D11 but D11 doesn't exist yet (will be created by D10's spill)
      // - D10 hasn't been evaluated yet

      // First, let's simulate the state after initial discovery:
      // C1 is being evaluated and has discovered its dependencies
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(["A1", "A2", "A3", "B1", "B2", "B3"]),
        frontierDependencies: new Map([["C1:C1", new Set(["B1"])]]), // B1 is a frontier candidate
        evaluationResult: undefined, // Not evaluated yet
      });

      // B1 depends on D11 (which doesn't exist yet)
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["D11"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });

      // D10 depends on A1, A2, B2
      storeManager.setEvaluatedNode("D10", {
        deps: new Set(["A1", "A2", "B2"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      // D11 depends on D10 (it's part of D10's spill result)
      storeManager.setEvaluatedNode("D11", {
        deps: new Set(["D10"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 4 },
        },
      });

      // Basic cells
      ["A1", "A2", "A3", "B2", "B3"].forEach((cell) => {
        storeManager.setEvaluatedNode(cell, {
          deps: new Set(),
          evaluationResult: {
            type: "value",
            result: { type: "number", value: 5 },
          },
        });
      });

      // When we try to build evaluation order for C1, it should include B1's dependencies
      const result = dependencyManager.buildEvaluationOrder("C1");

      expect(result.hasCycle).toBe(false);

      // The evaluation order should include all dependencies
      expect(result.evaluationOrder).toContain("A1");
      expect(result.evaluationOrder).toContain("A2");
      expect(result.evaluationOrder).toContain("B2");
      expect(result.evaluationOrder).toContain("D10");
      expect(result.evaluationOrder).toContain("B1");
      expect(result.evaluationOrder).toContain("D11");

      // Verify the correct order:
      // D10 should come before D11 (D11 depends on D10)
      // D11 should come before B1 (B1 depends on D11)
      const d10Index = result.evaluationOrder.indexOf("D10");
      const d11Index = result.evaluationOrder.indexOf("D11");
      const b1Index = result.evaluationOrder.indexOf("B1");

      expect(d10Index).toBeLessThan(d11Index);
      expect(d11Index).toBeLessThan(b1Index);
    });

    test("should handle dependencies on non-existent nodes", () => {
      // This simulates the case where B1 depends on D11,
      // but D11 doesn't exist yet (will be created by D10's spill)

      // C1 has B1 as a frontier dependency
      storeManager.setEvaluatedNode("C1", {
        deps: new Set(["A1"]),
        frontierDependencies: new Map([["C1:C1", new Set(["B1"])]]),
        evaluationResult: undefined,
      });

      // B1 depends on D11 which doesn't exist
      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["D11"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });

      // D11 doesn't exist in the store yet
      // (it will be created when D10 is evaluated)

      storeManager.setEvaluatedNode("A1", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("C1");

      expect(result.hasCycle).toBe(false);

      // The evaluation order should include D11 even though it doesn't exist yet
      expect(result.evaluationOrder).toContain("A1");
      expect(result.evaluationOrder).toContain("B1");
      expect(result.evaluationOrder).toContain("D11");

      // D11 should come before B1 (B1 depends on D11)
      const d11Index = result.evaluationOrder.indexOf("D11");
      const b1Index = result.evaluationOrder.indexOf("B1");
      expect(d11Index).toBeLessThan(b1Index);
    });
  });

  describe("Evaluation order verification", () => {
    test("should evaluate in correct order: frontier deps' transitive deps -> frontier deps -> regular transitive deps -> target", () => {
      // Set up a complex scenario:
      // A has frontier dependency on F1
      // F1 depends on F2, F3
      // A also depends on B which depends on C

      // Target cell A
      storeManager.setEvaluatedNode("A", {
        deps: new Set(["B"]),
        frontierDependencies: new Map([["A:A", new Set(["F1"])]]),
        evaluationResult: undefined,
      });

      // Frontier dependency F1 with its own dependencies
      storeManager.setEvaluatedNode("F1", {
        deps: new Set(["F2", "F3"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });

      // Transitive dependencies of F1
      storeManager.setEvaluatedNode("F2", {
        deps: new Set(["F4"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });

      storeManager.setEvaluatedNode("F3", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      storeManager.setEvaluatedNode("F4", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 4 },
        },
      });

      // Regular dependency chain
      storeManager.setEvaluatedNode("B", {
        deps: new Set(["C"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 5 },
        },
      });

      storeManager.setEvaluatedNode("C", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 6 },
        },
      });

      const result = dependencyManager.buildEvaluationOrder("A");

      expect(result.hasCycle).toBe(false);

      // Verify the order
      const order = result.evaluationOrder;

      // All dependencies should be included
      expect(order).toContain("F1");
      expect(order).toContain("F2");
      expect(order).toContain("F3");
      expect(order).toContain("F4");
      expect(order).toContain("B");
      expect(order).toContain("C");

      // Get indices
      const f1Index = order.indexOf("F1");
      const f2Index = order.indexOf("F2");
      const f3Index = order.indexOf("F3");
      const f4Index = order.indexOf("F4");
      const bIndex = order.indexOf("B");
      const cIndex = order.indexOf("C");

      // 1. Dependencies should be resolved in correct order
      expect(f4Index).toBeLessThan(f2Index); // F4 before F2 (F2 depends on F4)
      expect(f2Index).toBeLessThan(f1Index); // F2 before F1 (F1 depends on F2)
      expect(f3Index).toBeLessThan(f1Index); // F3 before F1 (F1 depends on F3)

      // 2. Regular dependencies should be resolved correctly
      expect(cIndex).toBeLessThan(bIndex); // C before B (B depends on C)

      // 3. All dependencies should come before the target
      expect(f1Index).toBeLessThan(order.indexOf("A")); // F1 before A
      expect(bIndex).toBeLessThan(order.indexOf("A")); // B before A
    });

    test("should handle multiple frontier dependencies with overlapping transitive deps", () => {
      // A has frontier dependencies on F1 and F2
      // F1 depends on X, Y
      // F2 depends on Y, Z
      // A also depends on B

      storeManager.setEvaluatedNode("A", {
        deps: new Set(["B"]),
        frontierDependencies: new Map([["A:A", new Set(["F1", "F2"])]]),
        evaluationResult: undefined,
      });

      storeManager.setEvaluatedNode("F1", {
        deps: new Set(["X", "Y"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });

      storeManager.setEvaluatedNode("F2", {
        deps: new Set(["Y", "Z"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });

      ["X", "Y", "Z", "B"].forEach((cell) => {
        storeManager.setEvaluatedNode(cell, {
          deps: new Set(),
          evaluationResult: {
            type: "value",
            result: { type: "number", value: 3 },
          },
        });
      });

      const result = dependencyManager.buildEvaluationOrder("A");

      expect(result.hasCycle).toBe(false);

      const order = result.evaluationOrder;

      // Y should only appear once even though it's a dependency of both F1 and F2
      const yCount = order.filter((node) => node === "Y").length;
      expect(yCount).toBe(1);

      // Dependencies should be resolved in correct order
      const f1Index = order.indexOf("F1");
      const f2Index = order.indexOf("F2");
      const xIndex = order.indexOf("X");
      const yIndex = order.indexOf("Y");
      const zIndex = order.indexOf("Z");
      const bIndex = order.indexOf("B");

      // Transitive dependencies should come before their dependents
      expect(xIndex).toBeLessThan(f1Index); // X before F1 (F1 depends on X)
      expect(yIndex).toBeLessThan(f1Index); // Y before F1 (F1 depends on Y)
      expect(yIndex).toBeLessThan(f2Index); // Y before F2 (F2 depends on Y)
      expect(zIndex).toBeLessThan(f2Index); // Z before F2 (F2 depends on Z)

      // All dependencies should come before the target
      expect(f1Index).toBeLessThan(order.indexOf("A")); // F1 before A
      expect(f2Index).toBeLessThan(order.indexOf("A")); // F2 before A
      expect(bIndex).toBeLessThan(order.indexOf("A")); // B before A
    });

    test("should not create cycle when frontier dep depends on target through spill", () => {
      // This is the key test for the multiplication scenario:
      // C1 has frontier dependency on B1
      // B1 depends on D11
      // D11 is created by D10's spill
      // D10 depends on A1, A2, B2
      // C1 also depends on A1, A2, A3, B1, B2, B3

      // This looks like a potential cycle but isn't because:
      // 1. B1 is a frontier dependency (may or may not affect C1)
      // 2. The evaluation order ensures D10 creates D11 before B1 needs it

      storeManager.setEvaluatedNode("C1", {
        deps: new Set(["A1", "A2", "A3", "B1", "B2", "B3"]),
        frontierDependencies: new Map([["C1:C1", new Set(["B1"])]]),
        evaluationResult: undefined,
      });

      storeManager.setEvaluatedNode("B1", {
        deps: new Set(["D11"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });

      storeManager.setEvaluatedNode("D10", {
        deps: new Set(["A1", "A2", "B2"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });

      storeManager.setEvaluatedNode("D11", {
        deps: new Set(["D10"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      ["A1", "A2", "A3", "B2", "B3"].forEach((cell) => {
        storeManager.setEvaluatedNode(cell, {
          deps: new Set(),
          evaluationResult: {
            type: "value",
            result: { type: "number", value: 4 },
          },
        });
      });

      const result = dependencyManager.buildEvaluationOrder("C1");

      expect(result.hasCycle).toBe(false);

      const order = result.evaluationOrder;

      // Verify complete order:
      // 1. Base cells (A1, A2, B2) - transitive deps of frontier deps
      // 2. D10 - dependency of D11
      // 3. D11 - dependency of B1
      // 4. B1 - frontier dependency
      // 5. A3, B3 - remaining regular deps

      const a1Index = order.indexOf("A1");
      const a2Index = order.indexOf("A2");
      const b2Index = order.indexOf("B2");
      const d10Index = order.indexOf("D10");
      const d11Index = order.indexOf("D11");
      const b1Index = order.indexOf("B1");

      // Transitive deps of frontier deps come first
      expect(d10Index).toBeLessThan(d11Index);
      expect(d11Index).toBeLessThan(b1Index);

      // Base cells needed by D10 come before D10
      expect(a1Index).toBeLessThan(d10Index);
      expect(a2Index).toBeLessThan(d10Index);
      expect(b2Index).toBeLessThan(d10Index);
    });

    test("should handle discarded frontier dependencies correctly in evaluation order", () => {
      // A has frontier dependency on F1 (discarded) and F2 (not discarded)
      // Only F2 and its dependencies should be in the evaluation order

      storeManager.setEvaluatedNode("A", {
        deps: new Set(["B"]),
        frontierDependencies: new Map([["A:A", new Set(["F1", "F2"])]]),
        discardedFrontierDependencies: new Map([["A:A", new Set(["F1"])]]),
        evaluationResult: undefined,
      });

      storeManager.setEvaluatedNode("F1", {
        deps: new Set(["X"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });

      storeManager.setEvaluatedNode("F2", {
        deps: new Set(["Y"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });

      ["X", "Y", "B"].forEach((cell) => {
        storeManager.setEvaluatedNode(cell, {
          deps: new Set(),
          evaluationResult: {
            type: "value",
            result: { type: "number", value: 3 },
          },
        });
      });

      const result = dependencyManager.buildEvaluationOrder("A");

      expect(result.hasCycle).toBe(false);

      const order = result.evaluationOrder;

      // F1 and X should not be in the evaluation order (F1 is discarded)
      expect(order).not.toContain("F1");
      expect(order).not.toContain("X");

      // F2 and Y should be in the order
      expect(order).toContain("F2");
      expect(order).toContain("Y");
      expect(order).toContain("B");

      // Dependencies should be resolved correctly
      const yIndex = order.indexOf("Y");
      const f2Index = order.indexOf("F2");
      const bIndex = order.indexOf("B");

      // Y should come before F2 (F2 depends on Y)
      expect(yIndex).toBeLessThan(f2Index);

      // All dependencies should come before the target
      expect(f2Index).toBeLessThan(order.indexOf("A"));
      expect(bIndex).toBeLessThan(order.indexOf("A"));
    });
  });

  describe("Transitive dependency methods", () => {
    test("should get transitive dependencies correctly", () => {
      // Set up A -> B -> C -> D
      storeManager.setEvaluatedNode("A", {
        deps: new Set(["B"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B", {
        deps: new Set(["C"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("C", {
        deps: new Set(["D"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });
      storeManager.setEvaluatedNode("D", {
        deps: new Set(),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 4 },
        },
      });

      const transitiveDeps = dependencyManager.getTransitiveDeps("A", (key) =>
        dependencyManager.getNodeDeps(key)
      );

      expect(transitiveDeps.has("B")).toBe(true);
      expect(transitiveDeps.has("C")).toBe(true);
      expect(transitiveDeps.has("D")).toBe(true);
      expect(transitiveDeps.size).toBe(3);
    });

    test("should handle circular dependencies in getTransitiveDeps", () => {
      // Set up A -> B -> C -> A
      storeManager.setEvaluatedNode("A", {
        deps: new Set(["B"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 1 },
        },
      });
      storeManager.setEvaluatedNode("B", {
        deps: new Set(["C"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 2 },
        },
      });
      storeManager.setEvaluatedNode("C", {
        deps: new Set(["A"]),
        evaluationResult: {
          type: "value",
          result: { type: "number", value: 3 },
        },
      });

      const transitiveDeps = dependencyManager.getTransitiveDeps("A", (key) =>
        dependencyManager.getNodeDeps(key)
      );

      // Should include all nodes but not infinitely loop
      expect(transitiveDeps.has("B")).toBe(true);
      expect(transitiveDeps.has("C")).toBe(true);
      expect(transitiveDeps.has("A")).toBe(true); // A is also included as a transitive dep of itself through the cycle
      expect(transitiveDeps.size).toBe(3);
    });
  });
});
