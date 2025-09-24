import type { DependencyManager, EvaluationManager } from "src/core/managers";
import type { StoreManager } from "src/core/managers/store-manager";
import type {
  CellAddress,
  FunctionEvaluationResult,
  SingleEvaluationResult,
  SpreadsheetRange,
} from "src/core/types";
import { cellAddressToKey, getCellReference, getRangeKey, indexToColumn } from "src/core/utils";
import { DependencyNode } from "./dependency-node";

export class EvaluationContext {
  private _currentCell: CellAddress;
  private dependencies: Set<string>;
  /**
   * Map is keyed by the range key, e.g. A4:D8
   */
  private frontierDependencies: Map<string, Set<string>>;
  /**
   * Map is keyed by the range key, e.g. A4:D8
   */
  private discardedFrontierDependencies: Map<string, Set<string>>;

  private dependenciesDidUpdate: boolean;

  constructor(
    private dependencyManager: DependencyManager,
    private storeManager: StoreManager,
    currentCell: CellAddress,
    currentDepNode?: DependencyNode
  ) {
    this._currentCell = currentCell;
    this.dependencies = currentDepNode?.deps ?? new Set();
    this.frontierDependencies =
      currentDepNode?.frontierDependencies ?? new Map();
    this.discardedFrontierDependencies =
      currentDepNode?.discardedFrontierDependencies ?? new Map();
    this.dependenciesDidUpdate = false;
  }

  get currentCell() {
    return this._currentCell;
  }

  addDependency(dependency: string) {
    if (!this.dependencies.has(dependency)) {
      this.dependenciesDidUpdate = true;
    }
    this.dependencies.add(dependency);
  }

  addFrontierDependency(dependency: string, range: SpreadsheetRange) {
    const rangeKey = getRangeKey(range);
    if (
      !this.frontierDependencies.has(rangeKey) &&
      !this.frontierDependencies.get(rangeKey)?.has(dependency)
    ) {
      this.dependenciesDidUpdate = true;
    }
    if (!this.frontierDependencies.has(rangeKey)) {
      this.frontierDependencies.set(rangeKey, new Set());
    }
    this.frontierDependencies.get(rangeKey)?.add(dependency);
  }

  private discardFrontierDependency(
    dependency: string,
    range: SpreadsheetRange
  ) {
    const rangeKey = getRangeKey(range);
    if (
      !this.discardedFrontierDependencies.has(rangeKey) &&
      !this.discardedFrontierDependencies.get(rangeKey)?.has(dependency)
    ) {
      this.dependenciesDidUpdate = true;
    }
    if (!this.discardedFrontierDependencies.has(rangeKey)) {
      this.discardedFrontierDependencies.set(rangeKey, new Set());
    }
    this.discardedFrontierDependencies.get(rangeKey)?.add(dependency);
  }

  maybeDiscardFrontierDependency(dependency: string, range: SpreadsheetRange) {
    if (this.isFrontierDependencyDiscarded(dependency, range)) {
      return;
    }
    // Only discard if the frontier dependency itself is resolved
    const depNode = this.storeManager.getEvaluatedNode(dependency);
    if (depNode?.resolved) {
      this.discardFrontierDependency(dependency, range);
    }
  }

  maybeUpgradeFrontierDependency(dependency: string, range: SpreadsheetRange) {
    if (this.isFrontierDependencyDiscarded(dependency, range)) {
      return;
    }
    // Only upgrade if the frontier dependency itself is resolved
    const depNode = this.storeManager.getEvaluatedNode(dependency);
    if (depNode?.resolved) {
      this.addDependency(dependency);
    }
  }

  getCurrentCell() {
    return this.currentCell;
  }

  getDependencies() {
    return this.dependencies;
  }

  getFrontierDependencies() {
    return this.frontierDependencies;
  }

  getDiscardedFrontierDependencies() {
    return this.discardedFrontierDependencies;
  }

  getDependenciesDidUpdate() {
    return this.dependenciesDidUpdate;
  }

  /**
   * TODO move to dependency manager
   */
  getTransitiveDependenciesResolved(deps: Set<string>) {
    // Get the current cell's dependency key to exclude self-references
    const currentCellKey = cellAddressToKey(this._currentCell);

    // Track visited nodes to avoid infinite loops in circular dependencies
    const visited = new Set<string>();
    visited.add(currentCellKey); // Don't revisit the current cell

    const checkResolved = (nodeKeys: Set<string>): boolean => {
      for (const nodeKey of nodeKeys) {
        if (visited.has(nodeKey)) {
          continue; // Skip already visited nodes (circular references)
        }
        visited.add(nodeKey);

        const node = this.storeManager.getEvaluatedNode(nodeKey);
        if (!node) {
          return false; // Node doesn't exist yet, not resolved
        }

        if (node.didUpdate) {
          return false; // Node itself is not update
        }

        // Check the node's direct dependencies
        const directDeps = this.dependencyManager.getNodeDeps(nodeKey);
        if (!checkResolved(directDeps)) {
          return false;
        }

        const frontierDeps =
          this.dependencyManager.getNodeFrontierDependencies(nodeKey);
        if (!checkResolved(frontierDeps)) {
          return false;
        }
      }
      return true;
    };

    return checkResolved(deps);
  }

  getDependencyAttributes(): DependencyAttributes {
    // A node is resolved when:
    // 1. This context didn't update any dependencies during evaluation
    // 2. AND all its transitive dependencies (including frontier) are also resolved
    const thisNodeDidNotUpdateDeps = !this.getDependenciesDidUpdate();
    const depsToCheck = new Set(this.dependencies);
    const discardedFrontierDepsToCheck = new Set<string>();
    for (const discardedFrontierDep of this.discardedFrontierDependencies.values()) {
      for (const dep of discardedFrontierDep) {
        discardedFrontierDepsToCheck.add(dep);
      }
    }
    for (const frontierDep of this.frontierDependencies.values()) {
      for (const dep of frontierDep) {
        if (discardedFrontierDepsToCheck.has(dep)) {
          continue;
        }
        depsToCheck.add(dep);
      }
    }
    const allTransitiveDepsResolved =
      this.getTransitiveDependenciesResolved(depsToCheck);

    return {
      deps: this.dependencies,
      frontierDependencies: this.frontierDependencies,
      discardedFrontierDependencies: this.discardedFrontierDependencies,
      didUpdate: !thisNodeDidNotUpdateDeps,
      resolved: thisNodeDidNotUpdateDeps && allTransitiveDepsResolved,
    };
  }

  isFrontierDependencyDiscarded(dependency: string, range: SpreadsheetRange) {
    const key = getRangeKey(range);
    return (
      this.discardedFrontierDependencies.has(key) &&
      this.discardedFrontierDependencies.get(key)?.has(dependency)
    );
  }
}

export type DependencyAttributes = {
  deps: Set<string>;
  frontierDependencies: Map<string, Set<string>>;
  discardedFrontierDependencies: Map<string, Set<string>>;
  didUpdate: boolean;
  resolved: boolean;
};
