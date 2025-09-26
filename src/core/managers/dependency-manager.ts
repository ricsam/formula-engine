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
import { flags } from "src/debug/flags";

interface TopologicalSortResult {
  type: "success" | "cycle";
  sorted?: string[];
  inCycle?: Set<string>;
}

export interface DependencyTreeNode {
  key: string;
  directDepsUpdated?: boolean;
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
  ): FunctionEvaluationResult {
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
    if (!result) {
      return {
        type: "awaiting-evaluation",
        cellAddress,
      };
    }
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
   * This is only used by buildEvaluationOrder, so we'll optimize it there
   */
  getTransitiveDepsForEvalOrder(
    nodeKey: string,
    visited: Set<string> = new Set()
  ): Set<string> {
    // Prevent infinite recursion
    if (visited.has(nodeKey)) {
      return new Set();
    }

    const node = this.getEvaluatedNode(nodeKey);
    // If the node is resolved, then we don't need to evaluate it
    if (node && node.resolved) {
      return new Set();
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
        const depTransitiveDeps = this.getTransitiveDepsForEvalOrder(
          dep,
          visited
        );
        for (const transitiveDep of depTransitiveDeps) {
          allNodes.add(transitiveDep);
        }
      }
    }

    // Remove this node from visited set for other branches
    visited.delete(nodeKey);

    return allNodes;
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
    const node = this.getEvaluatedNode(nodeKey);
    if (node && node.resolved && node?.evaluationOrder) {
      return node.evaluationOrder;
    }

    if (flags.isProfiling) {
      console.time("buildEvaluationOrder");
    }

    // Single-pass algorithm that:
    // 1. Collects all transitive dependencies
    // 2. Identifies which are frontier vs regular
    // 3. Detects cycles
    // 4. Builds evaluation order

    interface NodeInfo {
      isFrontierDep: boolean;
      frontierParents: Set<string>; // Nodes that have this as a frontier dependency
    }

    const allNodes = new Map<string, NodeInfo>();
    const visitedForDiscovery = new Set<string>();
    const visitingForCycle = new Set<string>();
    const cycleNodes = new Set<string>();
    const evaluationOrder: string[] = [];
    const visitedForOrder = new Set<string>();

    // Phase 1: Discover all nodes and their types, skipping resolved nodes
    const discoverNodes = (
      current: string,
      parentKey?: string,
      isFrontierEdge: boolean = false
    ) => {
      // Check if this node is resolved - if so, skip it entirely
      const currentNode = this.getEvaluatedNode(current);
      if (currentNode && currentNode.resolved) {
        return; // Skip resolved nodes completely
      }

      if (!allNodes.has(current)) {
        allNodes.set(current, {
          isFrontierDep: false,
          frontierParents: new Set(),
        });
      }

      const nodeInfo = allNodes.get(current)!;

      // If this edge is a frontier dependency, mark it
      if (isFrontierEdge && parentKey) {
        nodeInfo.isFrontierDep = true;
        nodeInfo.frontierParents.add(parentKey);
      }

      if (visitedForDiscovery.has(current)) {
        return;
      }

      visitedForDiscovery.add(current);

      // Get regular dependencies
      const regularDeps = this.getNodeDeps(current);
      for (const dep of regularDeps) {
        discoverNodes(dep, current, false);
      }

      // Get frontier dependencies
      const frontierDeps = this.getNodeFrontierDependencies(current);
      for (const dep of frontierDeps) {
        discoverNodes(dep, current, true);
      }
    };

    // Start discovery from the target node
    discoverNodes(nodeKey);

    // If the target node itself was resolved, return empty evaluation order
    if (allNodes.size === 0 && node && node.resolved) {
      const result: EvaluationOrder = {
        evaluationOrder: [nodeKey],
        hasCycle: false,
        hash: this.computeHash(new Set(nodeKey)),
      };

      if (node && node.resolved) {
        node.setEvaluationOrder(result);
      }

      if (flags.isProfiling) {
        console.timeEnd("buildEvaluationOrder");
        console.log(`Nodes: 0 (target node was resolved)`);
      }

      return result;
    }

    // Phase 2: Check for cycles (only considering regular dependencies)
    const checkCycles = (current: string): boolean => {
      if (visitingForCycle.has(current)) {
        cycleNodes.add(current);
        return true;
      }

      if (visitedForOrder.has(current)) {
        return false;
      }

      visitingForCycle.add(current);

      // Only check regular dependencies for cycles
      const regularDeps = this.getNodeDeps(current);
      for (const dep of regularDeps) {
        if (allNodes.has(dep) && checkCycles(dep)) {
          cycleNodes.add(current);
        }
      }

      visitingForCycle.delete(current);
      return cycleNodes.has(current);
    };

    // Check for cycles starting from all nodes
    for (const [nodeKey] of allNodes) {
      if (!visitedForOrder.has(nodeKey)) {
        checkCycles(nodeKey);
      }
    }

    // If there are cycles, return early
    if (cycleNodes.size > 0) {
      const result: EvaluationOrder = {
        evaluationOrder: [nodeKey],
        hasCycle: true,
        cycleNodes,
        hash: this.computeHash(new Set(allNodes.keys())),
      };

      if (node && node.resolved) {
        node.setEvaluationOrder(result);
      }

      if (flags.isProfiling) {
        console.timeEnd("buildEvaluationOrder");
        console.log(`Nodes: ${allNodes.size}`);
      }

      return result;
    }

    // Phase 3: Build evaluation order
    visitedForOrder.clear();
    const buildOrder = (current: string) => {
      if (visitedForOrder.has(current) || visitingForCycle.has(current)) {
        return;
      }

      visitingForCycle.add(current);

      // First visit regular dependencies
      const regularDeps = this.getNodeDeps(current);
      for (const dep of regularDeps) {
        if (allNodes.has(dep)) {
          buildOrder(dep);
        }
      }

      // Then visit frontier dependencies (without creating cycles)
      const frontierDeps = this.getNodeFrontierDependencies(current);
      for (const dep of frontierDeps) {
        if (allNodes.has(dep) && !visitingForCycle.has(dep)) {
          buildOrder(dep);
        }
      }

      visitingForCycle.delete(current);

      if (!visitedForOrder.has(current)) {
        visitedForOrder.add(current);
        evaluationOrder.push(current);
      }
    };

    // Build order for all nodes
    for (const [nodeKey] of allNodes) {
      buildOrder(nodeKey);
    }

    const result: EvaluationOrder = {
      evaluationOrder,
      hasCycle: false,
      hash: this.computeHash(new Set(allNodes.keys())),
    };

    if (node && node.resolved) {
      node.setEvaluationOrder(result);
    }

    if (flags.isProfiling) {
      console.timeEnd("buildEvaluationOrder");
      console.log(`Nodes: ${allNodes.size}`);
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
          directDepsUpdated: node?.directDepsUpdated ?? false,
          resolved: node?.resolved ?? false,
          self: true,
        };
      }

      // Avoid infinite recursion for circular dependencies
      if (visited.has(key)) {
        const node = this.getEvaluatedNode(key);
        return {
          key: cellRef,
          directDepsUpdated: node?.directDepsUpdated ?? false,
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
        directDepsUpdated: node?.directDepsUpdated ?? false,
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

  markResolvedNodes(nodeKey: string): void {
    const node = this.getEvaluatedNode(nodeKey);
    if (!node) {
      return;
    }

    // Track visited nodes to avoid infinite loops in circular dependencies
    const visited = new Set<string>();
    visited.add(nodeKey); // Don't revisit the current cell

    const checkDidUpdate = (nodeKeys: Set<string>): boolean => {
      for (const nodeKey of nodeKeys) {
        if (visited.has(nodeKey)) {
          continue;
        }
        visited.add(nodeKey);

        const node = this.getEvaluatedNode(nodeKey);
        if (!node) {
          return true; // Node doesn't exist yet, not resolved
        }

        // Check the node's direct dependencies
        const directDeps = this.getNodeDeps(nodeKey);
        if (checkDidUpdate(directDeps)) {
          return true;
        }

        const frontierDeps = this.getNodeFrontierDependencies(nodeKey);
        if (checkDidUpdate(frontierDeps)) {
          return true;
        }

        if (node.directDepsUpdated) {
          return true; // Node itself is not update
        }

        node.resolve();
      }
      return false;
    };

    const didUpdate = checkDidUpdate(
      this.getNodeDeps(nodeKey).union(this.getNodeFrontierDependencies(nodeKey))
    );
    if (!didUpdate && !node.directDepsUpdated) {
      node.resolve();
    }
  }
}
