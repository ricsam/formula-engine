import type { EvaluationContext } from "src/evaluator/evaluation-context";
import {
  type CellAddress,
  type EvaluationOrder,
  type FunctionEvaluationResult,
  type RangeAddress,
  type SingleEvaluationResult,
  type SpilledValue,
} from "../types";
import {
  cellAddressToKey,
  getCellReference,
  getRangeKey,
  isCellInRange,
  keyToCellAddress,
  rangeAddressToKey,
} from "../utils";

import { flags } from "src/debug/flags";
import { CellEvalNode } from "src/evaluator/cell-eval-node";
import { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import { CacheManager } from "./cache-manager";
import { WorkbookManager } from "./workbook-manager";
import { EmptyCellEvaluationNode } from "src/evaluator/empty-cell-evaluation-node";
import type { DependencyNode } from "./dependency-node";
import { AwaitingEvaluationError } from "src/evaluator/evaluation-error";

interface NodeInfo {
  isFrontierDep: boolean;
  frontierParents: Set<string>; // Nodes that have this as a frontier dependency
  node: CellEvalNode | RangeEvaluationNode;
}

export interface DependencyTreeNode {
  type: "cell" | "range" | "empty";
  circular?: boolean;
  key: string;
  directDepsUpdated?: boolean;
  resolved?: boolean;
  canResolve: boolean;
  resultType:
    | "awaiting-evaluation"
    | "spilled-values"
    | "value"
    | "range"
    | "error";
  deps?: DependencyTreeNode[];
  frontierDependencies?: DependencyTreeNode[];
  self?: boolean;
  _debug?: {
    rawFrontierDependencies?: string[];
    discardedFrontierDependencies?: string[];
    activeFrontierDependencies?: string[];
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
    CellEvalNode
  > = new Map();

  private emptyCells: Map<string, EmptyCellEvaluationNode> = new Map();

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

  /**
   * Key is workbook:sheetName:rangeKey, e.g. Workbook1:Sheet1:A1:D10, from rangeAddressToKey
   */
  private ranges: Map<string, RangeEvaluationNode> = new Map();

  constructor(
    private cacheManager: CacheManager,
    private workbookManager: WorkbookManager
  ) {}

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
    const node = this.getCellNode(key);
    context.dependencyNode.addDependency(node);
    const result = node.evaluationResult;
    if (!result) {
      throw new AwaitingEvaluationError(
        context.originCell.cellAddress,
        cellAddress
      );
    }
    return result;
  }

  clearEvaluationCache(): void {
    this.cacheManager.clear();
    this.evaluatedNodes.clear();
    this.emptyCells.clear();
    this.ranges.clear();
    this.spilledValues.clear();
  }

  getEmptyCellNode(nodeKey: string): EmptyCellEvaluationNode {
    if (!nodeKey.startsWith("empty:")) {
      throw new Error("Invalid empty cell node key: " + nodeKey);
    }
    if (!this.emptyCells.has(nodeKey)) {
      const node = new EmptyCellEvaluationNode(
        nodeKey,
        this,
        this.workbookManager
      );
      this.emptyCells.set(nodeKey, node);
      return node;
    }
    return this.emptyCells.get(nodeKey)!;
  }

  getCellNode(nodeKey: string): CellEvalNode | EmptyCellEvaluationNode {
    if (!nodeKey.startsWith("cell:") && !nodeKey.startsWith("empty:")) {
      throw new Error("Invalid cell node key: " + nodeKey);
    }
    if (!this.evaluatedNodes.has(nodeKey)) {
      const cellAddress = keyToCellAddress(nodeKey);
      if (this.workbookManager.isCellEmpty(cellAddress)) {
        return this.getEmptyCellNode(nodeKey.replace(/^cell:/, "empty:"));
      }

      const node = new CellEvalNode(nodeKey);
      this.evaluatedNodes.set(nodeKey, node);
      return node;
    }
    return this.evaluatedNodes.get(nodeKey)!;
  }

  getRangeNode(rangeKey: string): RangeEvaluationNode {
    if (!rangeKey.startsWith("range:")) {
      throw new Error("Invalid range node key: " + rangeKey);
    }
    if (!this.ranges.has(rangeKey)) {
      const node = new RangeEvaluationNode(
        rangeKey,
        this.cacheManager,
        this,
        this.workbookManager
      );
      this.ranges.set(rangeKey, node);
      return node;
    }
    return this.ranges.get(rangeKey)!;
  }

  getEvaluatedNodes(): Map<string, CellEvalNode> {
    return this.evaluatedNodes;
  }

  //#region dependency graph methods

  /**
   * Get transitive dependencies and transitive frontier dependencies
   * This is only used by buildEvaluationOrder, so we'll optimize it there
   */
  getTransitiveDepsForEvalOrder(
    node: DependencyNode,
    visited: Set<DependencyNode> = new Set()
  ): Set<DependencyNode> {
    // Prevent infinite recursion
    if (visited.has(node)) {
      return new Set();
    }

    // If the node is resolved, then we don't need to evaluate it
    if (node && node.resolved) {
      return new Set();
    }

    // Mark this node as visited for cycle detection
    visited.add(node);

    const allNodes = new Set<DependencyNode>();
    allNodes.add(node);

    // Get direct dependencies (regular + frontier)
    const directDeps = node.getDependencies();

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
    visited.delete(node);

    return allNodes;
  }

  /**
   * Build evaluation order for a cell, handling frontier dependencies specially
   *
   * Evaluation order is a topolocial sorted list of dependencies, where nodeKey is the last element:
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
   * e.g: nodeKey is Y10, then the evaluation order is [B1],[B10],[D2],{A1},[B2],{C4},[E4],(B1),(G4),(H1),Y10
   * 
   * Only dependencies can cause cycles, frontier dependencies cannot cause cycles.
   */
  buildEvaluationOrder(nodeKey: string): EvaluationOrder {
    const node = this.getCellNode(nodeKey);
    if (
      node &&
      node.resolved &&
      this.cacheManager.getEvaluationOrder(node.key)
    ) {
      return this.cacheManager.getEvaluationOrder(node.key)!;
    }

    // Algorithm:
    // 1. Discover all transitive dependencies (skipping resolved nodes)
    // 2. Detect cycles using Tarjan's algorithm
    // 3. Build topologically sorted evaluation order
    // 4. Return order with cycle information (cyclic nodes are marked but still included)

    const allNodes = new Map<string, DependencyNode>();
    const visitedForDiscovery = new Set<DependencyNode>();
    const cycleNodes = new Set<DependencyNode>();

    // Phase 1: Discover all nodes, skipping resolved nodes
    const discoverNodes = (currentNode: DependencyNode) => {
      // Skip resolved nodes completely
      if (currentNode && currentNode.resolved) {
        return;
      }

      if (!allNodes.has(currentNode.key)) {
        allNodes.set(currentNode.key, currentNode);
      }

      if (visitedForDiscovery.has(currentNode)) {
        return;
      }

      visitedForDiscovery.add(currentNode);

      // Get all dependencies (regular and frontier)
      const allDeps = currentNode.getAllDependencies();
      for (const dep of allDeps) {
        discoverNodes(dep);
      }
    };

    // Start discovery from the target node
    discoverNodes(node);

    // If the target node itself was resolved, return it alone
    if (allNodes.size === 0 && node && node.resolved) {
      const result: EvaluationOrder = {
        evaluationOrder: new Set([node]),
        hasCycle: false,
        hash: this.computeHash(new Set([node])),
      };

      if (node && node.resolved) {
        this.cacheManager.setEvaluationOrder(nodeKey, result);
      }

      return result;
    }

    // Phase 2: Detect cycles using Tarjan's algorithm (only for regular dependencies)
    const getStronglyConnectedComponents = (): Set<DependencyNode>[] => {
      const index = new Map<DependencyNode, number>();
      const lowlink = new Map<DependencyNode, number>();
      const onStack = new Set<DependencyNode>();
      const stack: DependencyNode[] = [];
      const sccs: Set<DependencyNode>[] = [];
      let currentIndex = 0;

      const strongConnect = (v: DependencyNode) => {
        index.set(v, currentIndex);
        lowlink.set(v, currentIndex);
        currentIndex++;
        stack.push(v);
        onStack.add(v);

        // Only consider regular dependencies (not frontier) for cycle detection
        const successors = v.getDependencies();
        for (const w of successors) {
          if (!allNodes.has(w.key)) {
            continue;
          }

          if (!index.has(w)) {
            strongConnect(w);
            lowlink.set(v, Math.min(lowlink.get(v)!, lowlink.get(w)!));
          } else if (onStack.has(w)) {
            lowlink.set(v, Math.min(lowlink.get(v)!, index.get(w)!));
          }
        }

        // Pop SCC from stack
        if (lowlink.get(v) === index.get(v)) {
          const scc = new Set<DependencyNode>();
          let w: DependencyNode;
          do {
            w = stack.pop()!;
            onStack.delete(w);
            scc.add(w);
          } while (w !== v);

          // Mark as cycle if: multiple nodes in SCC, or single node with self-loop
          if (scc.size > 1 || (scc.size === 1 && v.getDependencies().has(v))) {
            sccs.push(scc);
          }
        }
      };

      // Run on all discovered nodes
      for (const [_, node] of allNodes) {
        if (!index.has(node)) {
          strongConnect(node);
        }
      }

      return sccs;
    };

    // Identify all cyclic nodes
    for (const scc of getStronglyConnectedComponents()) {
      for (const node of scc) {
        cycleNodes.add(node);
      }
    }

    // Phase 3: Build topologically sorted evaluation order using DFS post-order
    // This creates proper topological order where dependencies come before dependents
    const evaluationOrderArray: DependencyNode[] = [];
    const visitedForOrder = new Set<DependencyNode>();
    const visitingForOrder = new Set<DependencyNode>(); // Track nodes currently being visited

    const buildOrder = (current: DependencyNode) => {
      // Already processed
      if (visitedForOrder.has(current)) {
        return;
      }

      // Detect cycles during traversal (shouldn't happen if Tarjan worked correctly)
      if (visitingForOrder.has(current)) {
        return; // Skip back edges
      }

      visitingForOrder.add(current);

      // Visit all dependencies first (regular dependencies)
      const regularDeps = current.getDependencies();
      for (const dep of regularDeps) {
        if (allNodes.has(dep.key) && !cycleNodes.has(dep)) {
          // Skip cyclic dependencies to avoid infinite recursion
          buildOrder(dep);
        } else if (allNodes.has(dep.key) && cycleNodes.has(dep)) {
          // Still try to visit cyclic nodes, but guard against infinite loops
          if (!visitingForOrder.has(dep)) {
            buildOrder(dep);
          }
        }
      }

      // Visit frontier dependencies
      const frontierDeps = current.getFrontierDependencies();
      for (const dep of frontierDeps) {
        if (allNodes.has(dep.key) && !visitingForOrder.has(dep)) {
          buildOrder(dep);
        }
      }

      visitingForOrder.delete(current);

      // Add to order in post-order (after visiting all dependencies)
      if (!visitedForOrder.has(current)) {
        visitedForOrder.add(current);
        evaluationOrderArray.push(current);
      }
    };

    // Build order starting from the target node
    // This ensures we get a proper topological sort
    buildOrder(node);

    // Convert array to Set while preserving insertion order (JS Sets maintain insertion order)
    const evaluationOrder = new Set(evaluationOrderArray);

    const hasCycle = cycleNodes.size > 0;
    const result: EvaluationOrder = {
      evaluationOrder,
      hasCycle,
      ...(hasCycle && { cycleNodes }),
      hash: this.computeHash(new Set(allNodes.values())),
    };

    if (node && node.resolved) {
      this.cacheManager.setEvaluationOrder(nodeKey, result);
    }

    return result;
  }

  /**
   * Compute a hash representing the current state of evaluated nodes
   * This hash changes when dependencies, frontier dependencies, or discarded frontier dependencies change
   */
  private computeHash(allNodes: Set<DependencyNode>): string {
    const nodeStates: string[] = [];

    for (const node of Array.from(allNodes).sort()) {
      if (node) {
        const deps = Array.from(node.getDependencies() || [])
          .map((dep) => dep.key)
          .sort()
          .join(",");

        // Handle frontier dependencies (Map<string, Set<string>>)
        const frontierDeps: string = Array.from(node.getFrontierDependencies())
          .map((dep) => dep.key)
          .sort()
          .join(";");

        const nodeState = `${node.key}:{deps:[${deps}],frontier:[${frontierDeps}]}`;
        nodeStates.push(nodeState);
      }
    }

    return nodeStates.join("|");
  }

  /**
   * Get a hierarchical dependency tree for a node
   */
  getDependencyTree(nodeKey: string): DependencyTreeNode {
    const visited = new Set<DependencyNode>();

    const nodeToType = (node: DependencyNode): "cell" | "range" | "empty" => {
      if (node instanceof RangeEvaluationNode) {
        return "range";
      }
      if (node instanceof EmptyCellEvaluationNode) {
        return "empty";
      }
      return "cell";
    };

    const buildTree = (
      node: DependencyNode,
      isSelf = false
    ): DependencyTreeNode => {
      const cellRef =
        node instanceof RangeEvaluationNode
          ? getRangeKey(node.address.range)
          : getCellReference(node.cellAddress);

      // Handle self-reference to avoid infinite recursion
      if (isSelf) {
        return {
          type: nodeToType(node),
          resultType:
            node instanceof RangeEvaluationNode
              ? "range"
              : node.evaluationResult ? node.evaluationResult.type : "awaiting-evaluation",
          canResolve: node.canResolve(),
          key: cellRef,
          directDepsUpdated: node.directDepsUpdated,
          resolved: node.resolved,
          self: true,
          circular: true,
        };
      }

      // Avoid infinite recursion for circular dependencies
      if (visited.has(node)) {
        return {
          type: nodeToType(node),
          resultType:
            node instanceof RangeEvaluationNode
              ? "range"
              : node.evaluationResult ? node.evaluationResult.type : "awaiting-evaluation",
          canResolve: node.canResolve(),
          key: cellRef,
          directDepsUpdated: node.directDepsUpdated,
          resolved: node.resolved,
          circular: true,
        };
      }

      visited.add(node);

      const directDeps = Array.from(node.getDependencies());
      let frontierDeps = Array.from(node.getFrontierDependencies());

      // Get regular dependencies
      const deps: DependencyTreeNode[] = directDeps.map((dep) =>
        buildTree(dep, dep.key === node.key)
      );

      const frontierDependencies: DependencyTreeNode[] = frontierDeps.map(
        (dep) => buildTree(dep, false)
      );

      visited.delete(node);

      const result: DependencyTreeNode = {
        type: nodeToType(node),
        resultType:
          node instanceof RangeEvaluationNode
            ? "range"
            : node.evaluationResult
              ? node.evaluationResult.type
              : "awaiting-evaluation",
        canResolve: node.canResolve(),
        key: cellRef,
        directDepsUpdated: node.directDepsUpdated,
        resolved: node.resolved,
      };

      // Only include deps and frontierDependencies if they have content
      if (deps.length > 0) {
        result.deps = deps;
      }
      if (frontierDependencies.length > 0) {
        result.frontierDependencies = frontierDependencies;
      }

      return result;
    };

    return buildTree(this.getCellNode(nodeKey));
  }
  //#endregion

  markResolvedNodes(nodeKey: string): void {
    const node = this.getCellNode(nodeKey);
    if (!node) {
      return;
    }

    // Track visited nodes to avoid infinite loops in circular dependencies
    const visited = new Set<DependencyNode>();
    visited.add(node); // Don't revisit the current cell

    const areTransitiveDepsResolved = (nodes: Set<DependencyNode>): boolean => {
      let canResolve = true;
      for (const node of nodes) {
        if (visited.has(node)) {
          continue;
        }
        visited.add(node);

        // Check the node's direct dependencies
        const directDeps = node.getAllDependencies();

        const a = areTransitiveDepsResolved(directDeps);
        const b = node.canResolve();

        if (!a || !b) {
          canResolve = false;
        }
        if (a && b) {
          node.resolve();
        }
      }
      return canResolve;
    };

    if (
      areTransitiveDepsResolved(node.getAllDependencies()) &&
      node.canResolve()
    ) {
      node.resolve();
    }
  }
}
