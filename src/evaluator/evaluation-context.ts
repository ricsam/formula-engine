import type { DependencyManager, EvaluationManager } from "src/core/managers";
import type {
  CellAddress,
  EvaluatedDependencyNode,
  FunctionEvaluationResult,
  SingleEvaluationResult,
  SpreadsheetRange,
} from "src/core/types";
import { getCellReference, getRangeKey, indexToColumn } from "src/core/utils";

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
    currentDepNode?: EvaluatedDependencyNode
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
    if (this.getTransitiveDependenciesResolved(new Set([dependency]))) {
      this.discardFrontierDependency(dependency, range);
    }
  }

  maybeUpgradeFrontierDependency(dependency: string, range: SpreadsheetRange) {
    if (this.isFrontierDependencyDiscarded(dependency, range)) {
      return;
    }
    if (this.getTransitiveDependenciesResolved(new Set([dependency]))) {
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

  getTransitiveDependenciesResolved(deps: Set<string>) {
    for (const dependency of this.dependencyManager.iterateOverTransitiveDependencies(
      deps
    )) {
      if (dependency.updatedDirectDeps) {
        return false;
      }
    }
    return true;
  }

  getEvaluatedDependencyNode(
    evaluationResult: FunctionEvaluationResult,
    originSpillResult?: SingleEvaluationResult
  ): EvaluatedDependencyNode {
    return {
      deps: this.dependencies,
      frontierDependencies: this.frontierDependencies,
      discardedFrontierDependencies: this.discardedFrontierDependencies,
      evaluationResult,
      originSpillResult,
      resolved: this.getTransitiveDependenciesResolved(this.dependencies),
      updatedDirectDeps: this.getDependenciesDidUpdate(),
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
