import type { CellEvalNode } from "src/evaluator/cell-eval-node";
import type { EvaluateAllCellsResult, RangeAddress } from "../types";
import { cellAddressToKey } from "../utils";
import type { WorkbookManager } from "./workbook-manager";
import type { DependencyManager } from "./dependency-manager";
import type { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import type { EmptyCellEvaluationNode } from "src/evaluator/empty-cell-evaluation-node";

export class FrontierDependencyManager {
  constructor(
    protected workbookManager: WorkbookManager,
    protected evaluationManager: DependencyManager
  ) {}

  private _resolved: boolean = false;

  private _directDepsUpdated: boolean = false;

  /**
   * frontierDependencies is the set of dependency node keys that could spill values onto the target range (if evaluationResult is spilled-values)
   * Key is from cellAddressToKey
   *
   * soft edge dependencies, which can not cause cycles in the dependency graph
   */
  private _frontierDependencies: Set<CellEvalNode> = new Set();

  /**
   * discardedFrontierDependencies is the set of dependency node keys that were discarded as frontier dependencies because
   * they they do not produce spilled values that spill onto the target range
   * Key is from cellAddressToKey
   */
  private _discardedFrontierDependencies: Set<CellEvalNode> = new Set();

  /**
   * hard edge dependencies, which can cause cycles in the dependency graph
   */
  private _dependencies: Set<CellEvalNode | EmptyCellEvaluationNode> =
    new Set();

  /**
   * cache, should maybe be stored in the cache manager
   */
  get iterateAllCells(): undefined | Iterable<EvaluateAllCellsResult> {
    return undefined;
  }

  public get frontierDependencies() {
    return this._frontierDependencies
      .difference(this._discardedFrontierDependencies)
      .difference(this._dependencies);
  }

  public get discardedFrontierDependencies() {
    return this._discardedFrontierDependencies;
  }

  public addFrontierDependency(dependency: CellEvalNode) {
    if (this._frontierDependencies.has(dependency)) {
      return;
    }
    this._directDepsUpdated = true;
    this._frontierDependencies.add(dependency);
  }

  public maybeDiscardFrontierDependency(dependency: CellEvalNode) {
    if (!this._resolved) {
      return;
    }
    if (this._discardedFrontierDependencies.has(dependency)) {
      return;
    }
    this._directDepsUpdated = true;
    this._discardedFrontierDependencies.add(dependency);
  }

  public maybeUpgradeFrontierDependency(dependency: CellEvalNode) {
    if (!this._resolved) {
      return;
    }
    if (this._dependencies.has(dependency)) {
      return;
    }
    this._directDepsUpdated = true;
    this._dependencies.add(dependency);
  }

  /**
   * a range is considered resolved when the dependencies has been established
   */
  public resolve() {
    if (this.canResolve()) {
      this._resolved = true;
    }
  }

  public canResolve() {
    return !this._directDepsUpdated;
  }

  public get resolved() {
    return this._resolved;
  }

  public get directDepsUpdated() {
    return this._directDepsUpdated;
  }

  public resetDirectDepsUpdated() {
    if (this._resolved) {
      return;
    }

    this._directDepsUpdated = false;
  }

  public getDependencies() {
    return this._dependencies;
  }

  public getFrontierDependencies() {
    return this.frontierDependencies;
  }

  public getAllDependencies() {
    return this._dependencies.union(this._frontierDependencies);
  }

  public getDiscardedFrontierDependencies() {
    return this.discardedFrontierDependencies;
  }

  public addDependency(dependency: CellEvalNode | EmptyCellEvaluationNode) {
    this._dependencies.add(dependency);
  }
}
