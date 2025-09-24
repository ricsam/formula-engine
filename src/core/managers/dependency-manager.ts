import type {
  DependencyAttributes,
  EvaluationContext,
} from "src/evaluator/evaluation-context";
import {
  type CellAddress,
  type EvaluationOrder,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpilledValue,
} from "../types";
import { cellAddressToKey, isCellInRange } from "../utils";

import { DependencyNode } from "src/evaluator/dependency-node";

interface TopologicalSortResult {
  type: "success" | "cycle";
  sorted?: string[];
  inCycle?: Set<string>;
}

export interface DependencyTreeNode {
  key: string;
  resolved?: boolean;
  deps?: DependencyTreeNode[];
  frontierDependencies?:
    | DependencyTreeNode[]
    | Record<string, DependencyTreeNode[]>;
  self?: boolean;
  _debug?: {
    rawFrontierDependencies?: Record<string, string[]>;
    discardedFrontierDependencies?: Record<string, string[]>;
    activeFrontierDependencies?: Record<string, string[]>;
  };
}

/**
 * The DependencyManager is responsible for storing the evaluated values and their dependencies.
 */
export class DependencyManager {
  /**
   * The dependency graph
   */
  private evaluatedNodes: Map<
    /**
     * key is the cell key, from cellAddressToKey
     */
    string,
    DependencyNode
  > = new Map();

  /**
   * registry of spilled values
   */
  public spilledValues: Map<
    /**
     * key is the dependency node key, from dependencyNodeToKey for the origin cell
     */
    string,
    SpilledValue
  > = new Map();

  constructor() {}

  getSpillValue(cellAddress: CellAddress): SpilledValue | undefined {
    for (const spilledValue of this.spilledValues.values()) {
      if (
        spilledValue.origin.sheetName !== cellAddress.sheetName ||
        spilledValue.origin.workbookName !== cellAddress.workbookName
      ) {
        continue;
      }
      if (
        spilledValue.origin.colIndex === cellAddress.colIndex &&
        spilledValue.origin.rowIndex === cellAddress.rowIndex
      ) {
        return undefined;
      }
      if (isCellInRange(cellAddress, spilledValue.spillOnto)) {
        return spilledValue;
      }
    }
    return undefined;
  }

  getSpilledAddress(
    cellAddress: CellAddress,
    /**
     * if the spilled value is already available, we can use it to get the source address
     */
    passedSpilledValue?: SpilledValue
  ): { address: CellAddress; spillOffset: { x: number; y: number } } {
    const spilledValue = passedSpilledValue ?? this.getSpillValue(cellAddress);
    if (!spilledValue) {
      throw new Error("Cell is not spilled");
    }
    const offsetLeft = cellAddress.colIndex - spilledValue.origin.colIndex;
    const offsetTop = cellAddress.rowIndex - spilledValue.origin.rowIndex;
    const address: CellAddress = {
      ...cellAddress,
      colIndex: spilledValue.origin.colIndex + offsetLeft,
      rowIndex: spilledValue.origin.rowIndex + offsetTop,
    };
    if (offsetLeft === 0 && offsetTop === 0) {
      throw new Error(
        "Spilled value is the same as the cell address! The origin has a pre-calculated value that can be used"
      );
    }
    return { address, spillOffset: { x: offsetLeft, y: offsetTop } };
  }

  /**
   * During evaluation, we can't use the formula evaluator to evaluate a cell, because it will create a cycle.
   * This method can be used to "evaluate" a cell during evaluation, without creating a cycle.
   *
   * Internally this method will try to look up the evaluated result for the cell,
   * if it doesn't exist it will push the cell address to the dependency graph,
   * causing the engine to re-evaluate the dependency graph,
   * such that on the second evaluation the cell's evaluation result will be available.
   */
  evalTimeSafeEvaluateCell(
    cellAddress: CellAddress,
    context: EvaluationContext
  ): FunctionEvaluationResult | undefined {
    const spilled = this.getSpillValue(cellAddress);
    if (spilled) {
      const spillSource = this.getSpilledAddress(cellAddress, spilled);
      const spillOrigin = this.evalTimeSafeEvaluateCell(
        spilled.origin,
        context
      );
      if (spillOrigin && spillOrigin.type === "spilled-values") {
        return spillOrigin.evaluate(spillSource.spillOffset, context);
      }
    }
    const key = cellAddressToKey(cellAddress);
    context.addDependency(key);
    const result = this.evaluatedNodes.get(key)?.evaluationResult;
    return result;
  }

  clearEvaluationCache(): void {
    this.evaluatedNodes.clear();
    this.spilledValues.clear();
  }

  setEvaluatedNode(
    nodeKey: string,
    attributes: DependencyAttributes,
    result: FunctionEvaluationResult,
    originSpillResult?: SingleEvaluationResult
  ): void {
    const currentNode =
      this.evaluatedNodes.get(nodeKey) ??
      new DependencyNode(nodeKey, result, originSpillResult);
    if (!currentNode) {
      throw new Error("Node not found");
    }
    currentNode.setDependencyAttributes(attributes);
    currentNode.setEvaluationResult(result, originSpillResult);

    this.evaluatedNodes.set(nodeKey, currentNode);
  }

  setEvaluatedResult(
    nodeKey: string,
    result: FunctionEvaluationResult,
    originSpillResult?: SingleEvaluationResult
  ): void {
    const currentNode =
      this.evaluatedNodes.get(nodeKey) ??
      new DependencyNode(nodeKey, result, originSpillResult);
    if (!currentNode) {
      throw new Error("Node not found");
    }
    currentNode.setEvaluationResult(result, originSpillResult);

    this.evaluatedNodes.set(nodeKey, currentNode);
  }

  getEvaluatedNode(nodeKey: string): DependencyNode | undefined {
    return this.evaluatedNodes.get(nodeKey);
  }

  getEvaluatedNodes(): Map<string, DependencyNode> {
    return this.evaluatedNodes;
  }

  //#region dependency graph methods
  /**
   * Get the direct dependencies of a node
   */
  getNodeDeps(nodeKey: string): Set<string> {
    return this.getEvaluatedNode(nodeKey)?.deps ?? new Set();
  }

  /**
   * Get the frontier dependencies of a node
   */
  getNodeFrontierDependencies(nodeKey: string): Set<string> {
    const node = this.getEvaluatedNode(nodeKey);
    if (!node) return new Set();

    return node.getAllFrontierDependencies();
  }

  /**
   * Get transitive dependencies and transitive frontier dependencies
   */
  getTransitiveDeps(
    nodeKey: string,
    visited: Set<string> = new Set()
  ): Set<string> {
    // Prevent infinite recursion
    if (visited.has(nodeKey)) {
      return new Set();
    }

    const node = this.getEvaluatedNode(nodeKey);
    // If we have cached transitive deps for a resolved node, use them
    if (node && node.resolved && node?.transitiveDeps) {
      return node.transitiveDeps;
    }

    // Mark this node as visited for cycle detection
    visited.add(nodeKey);

    const allNodes = new Set<string>();
    allNodes.add(nodeKey);

    // Get direct dependencies (regular + frontier)
    const regularDeps = this.getNodeDeps(nodeKey);
    const frontierDeps = this.getNodeFrontierDependencies(nodeKey);
    const directDeps = regularDeps.union(frontierDeps);

    // Recursively get transitive dependencies for each direct dependency
    for (const dep of directDeps) {
      if (!visited.has(dep)) {
        const depTransitiveDeps = this.getTransitiveDeps(dep, visited);
        for (const transitiveDep of depTransitiveDeps) {
          allNodes.add(transitiveDep);
        }
      }
    }

    // Remove this node from visited set for other branches
    visited.delete(nodeKey);

    // Cache the result if the node is resolved
    if (node && node.resolved) {
      node.setTransitiveDeps(allNodes);
    }

    return allNodes;
  }

  /**
   * Perform topological sort on dependencies with cycle detection
   */
  topologicalSort(deps: Set<string>): TopologicalSortResult {
    const sorted: string[] = [];
    const visited = new Set<string>();
    const visiting = new Set<string>();
    const inCycle = new Set<string>();

    const visit = (node: string): boolean => {
      if (visiting.has(node)) {
        // Found a cycle
        inCycle.add(node);
        return true;
      }
      if (visited.has(node)) {
        return false;
      }

      visiting.add(node);
      const nodeDeps = this.getNodeDeps(node);

      for (const dep of nodeDeps) {
        if (deps.has(dep) && visit(dep)) {
          inCycle.add(node);
        }
      }

      visiting.delete(node);
      visited.add(node);
      sorted.push(node);
      return inCycle.has(node);
    };

    for (const node of deps) {
      visit(node);
    }

    if (inCycle.size > 0) {
      return { type: "cycle", inCycle };
    }

    return { type: "success", sorted };
  }

  /**
   * Build evaluation order for a cell, handling frontier dependencies specially
   *
   * Evaluation order:
   * [A] = frontier dependency
   * {A} = transitive dependency of a frontier dependency
   * (A) = transitive dependency of a regular dependency
   * A = target cell
   *
   * 1. Frontier dependencies without dependencies
   * 1. Transitive dependencies of frontier dependencies
   * 2. Frontier dependencies
   * 3. Transitive dependencies of regular dependencies
   * 4. Target cell
   *
   * e.g: [B1],[B10],[D2],{A1},[B2],{C4},[E4],(B1),(G4),(H1),Y10
   */
  buildEvaluationOrder(nodeKey: string): EvaluationOrder {
    const evaluationOrder: string[] = [];

    const node = this.getEvaluatedNode(nodeKey);
    if (node && node.resolved && node?.evaluationOrder) {
      return node.evaluationOrder;
    }

    // Collect all nodes in the dependency graph (regular + frontier)
    const allNodes = this.getTransitiveDeps(nodeKey);

    // Separate frontier and regular dependencies
    const allFrontierDeps = new Set<string>();
    const allRegularDeps = new Set<string>();

    for (const node of allNodes) {
      // Check if this node is a frontier dependency of any node in the graph
      let isFrontierDep = false;
      for (const graphNode of allNodes) {
        if (graphNode !== node) {
          const frontierDeps = this.getNodeFrontierDependencies(graphNode);
          if (frontierDeps.has(node)) {
            isFrontierDep = true;
            break;
          }
        }
      }

      // Also check if it's a frontier dependency of the target node
      const targetFrontierDeps = this.getNodeFrontierDependencies(nodeKey);
      if (targetFrontierDeps.has(node)) {
        isFrontierDep = true;
      }

      if (isFrontierDep) {
        allFrontierDeps.add(node);
      } else {
        allRegularDeps.add(node);
      }
    }

    // Check for cycles in regular dependencies only
    const cycleCheckResult = this.topologicalSort(allRegularDeps);
    if (cycleCheckResult.type === "cycle") {
      const result: EvaluationOrder = {
        evaluationOrder: [],
        hasCycle: true,
        cycleNodes: cycleCheckResult.inCycle,
        hash: this.computeHash(allNodes),
      };
      if (node && node.resolved) {
        node.setEvaluationOrder(result);
      }
      return result;
    }

    // Create a specialized topological sort that handles frontier dependencies
    const visitedNodes = new Set<string>();
    const visitingNodes = new Set<string>();

    const visit = (node: string): void => {
      if (visitingNodes.has(node) || visitedNodes.has(node)) {
        return;
      }

      visitingNodes.add(node);

      // First, visit all regular dependencies
      const regularDeps = this.getNodeDeps(node);
      for (const dep of regularDeps) {
        if (allNodes.has(dep)) {
          visit(dep);
        }
      }

      // Then, visit all frontier dependencies (but don't create cycles)
      const frontierDeps = this.getNodeFrontierDependencies(node);
      for (const dep of frontierDeps) {
        if (allNodes.has(dep) && !visitingNodes.has(dep)) {
          visit(dep);
        }
      }

      visitingNodes.delete(node);
      if (!visitedNodes.has(node)) {
        visitedNodes.add(node);
        evaluationOrder.push(node);
      }
    };

    // Visit all nodes
    for (const node of allNodes) {
      visit(node);
    }

    const result: EvaluationOrder = {
      evaluationOrder,
      hasCycle: false,
      hash: this.computeHash(allNodes),
    };
    if (node && node.resolved) {
      node.setEvaluationOrder(result);
    }
    return result;
  }

  /**
   * Compute a hash representing the current state of evaluated nodes
   * This hash changes when dependencies, frontier dependencies, or discarded frontier dependencies change
   */
  private computeHash(allNodes: Set<string>): string {
    const nodeStates: string[] = [];

    for (const nodeKey of Array.from(allNodes).sort()) {
      const node = this.getEvaluatedNode(nodeKey);
      if (node) {
        const deps = Array.from(node.deps || [])
          .sort()
          .join(",");

        // Handle frontier dependencies (Map<string, Set<string>>)
        const frontierDeps: string[] = [];
        if (node.frontierDependencies) {
          for (const [range, rangeDeps] of Array.from(
            node.frontierDependencies.entries()
          ).sort()) {
            const sortedDeps = Array.from(rangeDeps).sort().join(",");
            frontierDeps.push(`${range}:[${sortedDeps}]`);
          }
        }

        // Handle discarded frontier dependencies (Map<string, Set<string>>)
        const discardedFrontierDeps: string[] = [];
        if (node.discardedFrontierDependencies) {
          for (const [range, rangeDeps] of Array.from(
            node.discardedFrontierDependencies.entries()
          ).sort()) {
            const sortedDeps = Array.from(rangeDeps).sort().join(",");
            discardedFrontierDeps.push(`${range}:[${sortedDeps}]`);
          }
        }

        const nodeState = `${nodeKey}:{deps:[${deps}],frontier:{${frontierDeps.join(";")}},discarded:{${discardedFrontierDeps.join(";")}}}`;
        nodeStates.push(nodeState);
      }
    }

    return nodeStates.join("|");
  }

  /**
   * Get a hierarchical dependency tree for a node
   */
  getDependencyTree(nodeKey: string): DependencyTreeNode {
    const visited = new Set<string>();

    const buildTree = (key: string, isSelf = false): DependencyTreeNode => {
      const cellRef = key.split(":")[3] || key;

      // Handle self-reference to avoid infinite recursion
      if (isSelf) {
        const node = this.getEvaluatedNode(key);
        return {
          key: cellRef,
          resolved: node?.resolved ?? false,
          self: true,
        };
      }

      // Avoid infinite recursion for circular dependencies
      if (visited.has(key)) {
        const node = this.getEvaluatedNode(key);
        return {
          key: cellRef,
          resolved: node?.resolved ?? false,
        };
      }

      visited.add(key);

      // Get regular dependencies
      const regularDeps = this.getNodeDeps(key);
      const deps: DependencyTreeNode[] = [];
      for (const dep of regularDeps) {
        deps.push(buildTree(dep, dep === nodeKey));
      }

      // Get frontier dependencies and debug info
      const node = this.getEvaluatedNode(key);
      const rawFrontierDepsMap: Map<
        string,
        Set<string>
      > = node?.frontierDependencies ?? new Map();
      const discardedFrontierDepsMap: Map<
        string,
        Set<string>
      > = node?.discardedFrontierDependencies ?? new Map();
      const activeFrontierDeps = this.getNodeFrontierDependencies(key);

      // Build frontier dependencies structure
      let frontierDependencies:
        | DependencyTreeNode[]
        | Record<string, DependencyTreeNode[]>
        | undefined;

      if (rawFrontierDepsMap.size > 0) {
        // Always use object format for consistency with the new Map-based structure
        const ranges = Array.from(rawFrontierDepsMap.keys());
        frontierDependencies = {};

        for (const range of ranges) {
          const rangeDeps: Set<string> =
            rawFrontierDepsMap.get(range) ?? new Set();
          const discardedRangeDeps: Set<string> =
            discardedFrontierDepsMap.get(range) ?? new Set();
          // Exclude frontier dependencies that are already regular dependencies
          const activeRangeDeps = rangeDeps
            .difference(discardedRangeDeps)
            .difference(regularDeps);

          if (activeRangeDeps.size > 0) {
            frontierDependencies[range] = [];
            for (const dep of activeRangeDeps) {
              frontierDependencies[range].push(buildTree(dep, dep === nodeKey));
            }
          }
        }

        // If no active dependencies, don't include the frontierDependencies field
        if (Object.keys(frontierDependencies).length === 0) {
          frontierDependencies = undefined;
        }
      }

      visited.delete(key);

      const result: DependencyTreeNode = {
        key: cellRef,
        resolved: node?.resolved ?? false,
      };

      // Only include deps and frontierDependencies if they have content
      if (deps.length > 0) {
        result.deps = deps;
      }
      if (frontierDependencies) {
        result.frontierDependencies = frontierDependencies;
      }

      // Add debug information if there are any frontier dependencies (raw, discarded, or active)
      const hasAnyFrontierDeps =
        rawFrontierDepsMap.size > 0 ||
        discardedFrontierDepsMap.size > 0 ||
        activeFrontierDeps.size > 0;

      if (hasAnyFrontierDeps) {
        const rawFrontierDebug: Record<string, string[]> = {};
        const discardedFrontierDebug: Record<string, string[]> = {};
        const activeFrontierDebug: Record<string, string[]> = {};

        // Build raw frontier dependencies by range
        for (const [range, rangeDeps] of rawFrontierDepsMap.entries()) {
          if (rangeDeps.size > 0) {
            rawFrontierDebug[range] = Array.from(rangeDeps).map(
              (k) => k.split(":")[3] || k
            );
          }
        }

        // Build discarded frontier dependencies by range
        for (const [range, rangeDeps] of discardedFrontierDepsMap.entries()) {
          if (rangeDeps.size > 0) {
            discardedFrontierDebug[range] = Array.from(rangeDeps).map(
              (k) => k.split(":")[3] || k
            );
          }
        }

        // Build active frontier dependencies by range
        for (const [range, rangeDeps] of rawFrontierDepsMap.entries()) {
          const discardedRangeDeps: Set<string> =
            discardedFrontierDepsMap.get(range) ?? new Set<string>();
          const activeRangeDeps: Set<string> =
            rangeDeps.difference(discardedRangeDeps);
          if (activeRangeDeps.size > 0) {
            activeFrontierDebug[range] = Array.from(activeRangeDeps).map(
              (k) => k.split(":")[3] || k
            );
          }
        }

        result._debug = {
          rawFrontierDependencies:
            Object.keys(rawFrontierDebug).length > 0
              ? rawFrontierDebug
              : undefined,
          discardedFrontierDependencies:
            Object.keys(discardedFrontierDebug).length > 0
              ? discardedFrontierDebug
              : undefined,
          activeFrontierDependencies:
            Object.keys(activeFrontierDebug).length > 0
              ? activeFrontierDebug
              : undefined,
        };
      }

      return result;
    };

    return buildTree(nodeKey);
  }
  //#endregion
}
