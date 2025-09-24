import type {
  CellAddress,
  FunctionEvaluationResult,
  SingleEvaluationResult,
  EvaluationOrder,
} from "src/core/types";
import type { DependencyAttributes } from "./evaluation-context";
import { keyToCellAddress } from "src/core/utils";

export class DependencyNode {
  private _cellAddress: CellAddress;
  private _key: string;
  private _deps: Set<string>;
  private _frontierDependencies: Map<string, Set<string>>;
  private _discardedFrontierDependencies: Map<string, Set<string>>;
  private _evaluationResult: FunctionEvaluationResult;
  private _originSpillResult?: SingleEvaluationResult;
  private _resolved?: boolean;
  private _didUpdate?: boolean;
  private _transitiveDeps?: Set<string>;

  constructor(
    key: string,
    evaluationResult: FunctionEvaluationResult,
    originSpillResult?: SingleEvaluationResult
  ) {
    this._cellAddress = keyToCellAddress(key);
    this._key = key;
    this._deps = new Set();
    this._frontierDependencies = new Map();
    this._discardedFrontierDependencies = new Map();
    this._evaluationResult = evaluationResult;
    this._originSpillResult = originSpillResult;
    this._resolved = false;
    this._didUpdate = false;
  }

  public get key() {
    return this._key;
  }

  public get cellAddress() {
    return this._cellAddress;
  }

  /**
   * deps is the set of dependency node keys
   */
  public get deps() {
    return this._deps;
  }

  /**
   * frontierDependencies is the set of dependency node keys that could spill values onto the target range (if evaluationResult is spilled-values)
   *
   * Map is keyed by the range key, e.g. A4:D8
   */
  public get frontierDependencies() {
    return this._frontierDependencies;
  }

  /**
   * discardedFrontierDependencies is the set of dependency node keys that were discarded as frontier dependencies because
   * they they do not produce spilled values that spill onto the target range
   *
   * Map is keyed by the range key, e.g. A4:D8
   */
  public get discardedFrontierDependencies() {
    return this._discardedFrontierDependencies;
  }

  /**
   * evaluationResult is the evaluation result
   */
  public get evaluationResult() {
    return this._evaluationResult;
  }

  /**
   * originSpillResult is the evaluation result of the spilled origin
   */
  public get originSpillResult() {
    return this._originSpillResult;
  }

  /**
   * resolved is true if all transitive dependencies are resolved.
   * A dependency is resolved if it had no updates to its dependencies during the evaluation.
   *
   * The dependencies considered include regular dependencies, frontier dependencies and discarded frontier dependencies.
   *
   */
  public get resolved() {
    return this._resolved;
  }

  public get didUpdate() {
    return this._didUpdate;
  }

  public setDependencyAttributes(attributes: DependencyAttributes) {
    // Check if dependencies have changed
    const depsChanged = !this.setsEqual(this._deps, attributes.deps);
    const frontierDepsChanged = !this.mapsEqual(
      this._frontierDependencies,
      attributes.frontierDependencies
    );
    const discardedDepsChanged = !this.mapsEqual(
      this._discardedFrontierDependencies,
      attributes.discardedFrontierDependencies
    );

    this._deps = attributes.deps;
    this._frontierDependencies = attributes.frontierDependencies;
    this._discardedFrontierDependencies =
      attributes.discardedFrontierDependencies;
    this._didUpdate = attributes.didUpdate;
    this._resolved = attributes.resolved;

    if (!this.resolved) {
      // invalidate transitive deps
      this._transitiveDeps = undefined;
      this._evaluationOrder = undefined;
    }

    // Invalidate cache only if dependencies actually changed
    if (depsChanged || frontierDepsChanged || discardedDepsChanged) {
      this._allFrontierDependenciesCache = undefined;
    }
  }

  public setEvaluationResult(
    result: FunctionEvaluationResult,
    originSpillResult?: SingleEvaluationResult
  ) {
    if (result.type === "spilled-values" && !originSpillResult) {
      throw new Error(
        "Evaluation result is spilled-values but origin spill result is not set"
      );
    }
    this._evaluationResult = result;
    this._originSpillResult = originSpillResult;
  }

  private _allFrontierDependenciesCache: Set<string> | undefined;

  public getAllFrontierDependencies(): Set<string> {
    if (this._allFrontierDependenciesCache) {
      return this._allFrontierDependenciesCache;
    }
    const frontierDepsMap = this.frontierDependencies ?? new Map();
    const discardedDepsMap = this.discardedFrontierDependencies ?? new Map();
    const deps = this.deps ?? new Set();

    const allFrontierDeps = new Set<string>();
    const allDiscardedDeps = new Set<string>();

    // Collect all frontier dependencies across all ranges
    for (const rangeDeps of frontierDepsMap.values()) {
      for (const dep of rangeDeps) {
        if (deps.has(dep)) {
          continue;
        }
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
    const result = allFrontierDeps.difference(allDiscardedDeps);
    this._allFrontierDependenciesCache = result;
    return result;
  }

  private setsEqual(set1: Set<string>, set2: Set<string>): boolean {
    if (set1.size !== set2.size) {
      return false;
    }
    for (const item of set1) {
      if (!set2.has(item)) {
        return false;
      }
    }
    return true;
  }

  private mapsEqual(
    map1: Map<string, Set<string>>,
    map2: Map<string, Set<string>>
  ): boolean {
    if (map1.size !== map2.size) {
      return false;
    }
    for (const [key, set1] of map1) {
      const set2 = map2.get(key);
      if (!set2 || !this.setsEqual(set1, set2)) {
        return false;
      }
    }
    return true;
  }

  public setTransitiveDeps(deps: Set<string>) {
    if (!this.resolved) {
      throw new Error("Cannot set transitive deps for an unresolved node");
    }
    this._transitiveDeps = deps;
  }

  public get transitiveDeps() {
    return this._transitiveDeps;
  }

  private _evaluationOrder?: EvaluationOrder;

  public get evaluationOrder() {
    return this._evaluationOrder;
  }

  public setEvaluationOrder(order: EvaluationOrder) {
    if (!this.resolved) {
      throw new Error("Cannot set evaluation order for an unresolved node");
    }
    this._evaluationOrder = order;
  }
}
