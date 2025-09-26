import type { DependencyManager } from "src/core/managers";
import type { CellAddress, SpreadsheetRange } from "src/core/types";
import { cellAddressToKey, getRangeKey } from "src/core/utils";
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
    const depNode = this.dependencyManager.getEvaluatedNode(dependency);
    if (depNode?.resolved) {
      this.discardFrontierDependency(dependency, range);
    }
  }

  maybeUpgradeFrontierDependency(dependency: string, range: SpreadsheetRange) {
    if (this.isFrontierDependencyDiscarded(dependency, range)) {
      return;
    }
    // Only upgrade if the frontier dependency itself is resolved
    const depNode = this.dependencyManager.getEvaluatedNode(dependency);
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

  getDependencyAttributes(): DependencyAttributes {
    const depsToCheck = new Set(this.dependencies);

    for (const [rangeKey, frontierDep] of this.frontierDependencies) {
      for (const dep of frontierDep) {
        if (this.discardedFrontierDependencies.get(rangeKey)?.has(dep)) {
          continue;
        }
        depsToCheck.add(dep);
      }
    }

    return {
      deps: this.dependencies,
      frontierDependencies: this.frontierDependencies,
      discardedFrontierDependencies: this.discardedFrontierDependencies,
      directDepsUpdated: this.getDependenciesDidUpdate(),
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
  directDepsUpdated: boolean;
};
