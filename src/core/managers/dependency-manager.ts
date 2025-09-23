import type { EvaluatedDependencyNode, DependencyNode } from "../types";
import type { StoreManager } from "./store-manager";
import type { WorkbookManager } from "./workbook-manager";
import { keyToDependencyNode } from "../utils/dependency-node-key";
import { FormulaError } from "../types";

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

export class DependencyManager {
  constructor(
    private storeManager: StoreManager,
    private workbookManager: WorkbookManager
  ) {}

  /**
   * Get the direct dependencies of a node
   */
  getNodeDeps(nodeKey: string): Set<string> {
    return this.storeManager.getEvaluatedNode(nodeKey)?.deps ?? new Set();
  }

  /**
   * Get the frontier dependencies of a node
   */
  getNodeFrontierDependencies(nodeKey: string): Set<string> {
    const node = this.storeManager.getEvaluatedNode(nodeKey);
    if (!node) return new Set();

    const frontierDepsMap = node.frontierDependencies ?? new Map();
    const discardedDepsMap = node.discardedFrontierDependencies ?? new Map();

    const allFrontierDeps = new Set<string>();
    const allDiscardedDeps = new Set<string>();

    // Collect all frontier dependencies across all ranges
    for (const rangeDeps of frontierDepsMap.values()) {
      for (const dep of rangeDeps) {
        allFrontierDeps.add(dep);
      }
    }

    // Collect all discarded dependencies across all ranges
    for (const rangeDeps of discardedDepsMap.values()) {
      for (const dep of rangeDeps) {
        allDiscardedDeps.add(dep);
      }
    }

    // Return frontier dependencies minus discarded ones
    return allFrontierDeps.difference(allDiscardedDeps);
  }

  /**
   * Get transitive dependencies using the provided dependency getter
   */
  getTransitiveDeps(
    nodeKey: string,
    getDeps: (key: string) => Set<string>,
    visited: Set<string> = new Set()
  ): Set<string> {
    if (visited.has(nodeKey)) {
      return new Set();
    }
    visited.add(nodeKey);

    const directDeps = getDeps(nodeKey);
    const transitiveDeps = new Set(directDeps);

    for (const dep of directDeps) {
      const depTransitive = this.getTransitiveDeps(dep, getDeps, visited);
      for (const transitiveDep of depTransitive) {
        transitiveDeps.add(transitiveDep);
      }
    }

    return transitiveDeps;
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
  buildEvaluationOrder(nodeKey: string): {
    evaluationOrder: string[];
    hasCycle: boolean;
    cycleNodes?: Set<string>;
    hash: string;
  } {
    const evaluationOrder: string[] = [];

    // Collect all nodes in the dependency graph (regular + frontier)
    const allNodes = new Set<string>();
    const toVisit = [nodeKey];
    const visited = new Set<string>();

    while (toVisit.length > 0) {
      const current = toVisit.pop()!;
      if (visited.has(current)) continue;
      visited.add(current);

      allNodes.add(current);

      // Add regular dependencies
      const regularDeps = this.getNodeDeps(current);
      for (const dep of regularDeps) {
        if (!visited.has(dep)) {
          toVisit.push(dep);
        }
      }

      // Add frontier dependencies
      const frontierDeps = this.getNodeFrontierDependencies(current);
      for (const dep of frontierDeps) {
        if (!visited.has(dep)) {
          toVisit.push(dep);
        }
      }
    }

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
      return {
        evaluationOrder: [],
        hasCycle: true,
        cycleNodes: cycleCheckResult.inCycle,
        hash: this.computeHash(allNodes),
      };
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

    return {
      evaluationOrder,
      hasCycle: false,
      hash: this.computeHash(allNodes),
    };
  }

  /**
   * Compute a hash representing the current state of evaluated nodes
   * This hash changes when dependencies, frontier dependencies, or discarded frontier dependencies change
   */
  private computeHash(allNodes: Set<string>): string {
    const nodeStates: string[] = [];

    for (const nodeKey of Array.from(allNodes).sort()) {
      const node = this.storeManager.getEvaluatedNode(nodeKey);
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
   * Check if a node has been evaluated
   */
  isNodeEvaluated(nodeKey: string): boolean {
    return this.storeManager.getEvaluatedNode(nodeKey) !== undefined;
  }

  /**
   * Iterate over all transitive dependencies of the given dependencies
   * Returns an iterator that yields dependency information including whether deps were updated
   */
  *iterateOverTransitiveDependencies(dependencies: Set<string>): Generator<{
    nodeKey: string;
    updatedDirectDeps: boolean;
    resolved: boolean;
  }> {
    const visited = new Set<string>();
    const toVisit = Array.from(dependencies);

    while (toVisit.length > 0) {
      const current = toVisit.pop()!;
      if (visited.has(current)) continue;
      visited.add(current);

      const node = this.storeManager.getEvaluatedNode(current);
      if (!node) continue;

      // Yield information about this dependency
      yield {
        nodeKey: current,
        updatedDirectDeps: node.updatedDirectDeps ?? false,
        resolved: node.resolved ?? false,
      };

      // Add direct dependencies to visit queue
      const directDeps = this.getNodeDeps(current);
      for (const dep of directDeps) {
        if (!visited.has(dep)) {
          toVisit.push(dep);
        }
      }

      // Add active frontier dependencies to visit queue
      const frontierDeps = this.getNodeFrontierDependencies(current);
      for (const dep of frontierDeps) {
        if (!visited.has(dep)) {
          toVisit.push(dep);
        }
      }
    }
  }

  /**
   * Mark nodes as having a cycle error
   */
  markNodesAsCycle(nodes: Set<string>, message: string): void {
    for (const node of nodes) {
      const currentNode = this.storeManager.getEvaluatedNode(node);
      this.storeManager.setEvaluatedNode(node, {
        deps: currentNode?.deps ?? new Set(),
        frontierDependencies: currentNode?.frontierDependencies ?? new Map(),
        discardedFrontierDependencies:
          currentNode?.discardedFrontierDependencies ?? new Map(),
        evaluationResult: {
          type: "error",
          err: FormulaError.CYCLE,
          message,
        },
      });
    }
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
        const node = this.storeManager.getEvaluatedNode(key);
        return {
          key: cellRef,
          resolved: node?.resolved ?? false,
          self: true,
        };
      }

      // Avoid infinite recursion for circular dependencies
      if (visited.has(key)) {
        const node = this.storeManager.getEvaluatedNode(key);
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
      const node = this.storeManager.getEvaluatedNode(key);
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
          const activeRangeDeps = rangeDeps.difference(discardedRangeDeps);

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
}
