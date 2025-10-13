import type { CellEvalNode } from "src/evaluator/cell-eval-node";
import type { EvaluateAllCellsResult, RangeAddress } from "../types";
import { cellAddressToKey, checkRangeIntersection, isCellInRange } from "../utils";
import type { WorkbookManager } from "./workbook-manager";
import type { DependencyManager } from "./dependency-manager";
import type { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import { EmptyCellEvaluationNode } from "src/evaluator/empty-cell-evaluation-node";

export class FrontierDependencyManager {
  constructor(
    private frontierRange: RangeAddress,
    protected workbookManager: WorkbookManager,
    protected evaluationManager: DependencyManager
  ) {}

  private _resolved: boolean = false;
  private _hasRegisteredFrontierDependencies: boolean = false;

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
    if (!this._hasRegisteredFrontierDependencies) {
      this.workbookManager
        .getFrontierCandidates(this.frontierRange)
        .forEach((candidate) => {
          const node = this.evaluationManager.getCellNode(
            cellAddressToKey(candidate)
          );
          if (node instanceof EmptyCellEvaluationNode) {
            throw new Error("A frontier dependencies can not be an empty cell");
          }
          this.addFrontierDependency(node);
        });
      this._hasRegisteredFrontierDependencies = true;
    }

    return this._frontierDependencies
      .difference(this._discardedFrontierDependencies)
      .difference(this._dependencies);
  }

  public get discardedFrontierDependencies() {
    return this._discardedFrontierDependencies;
  }

  private addFrontierDependency(dependency: CellEvalNode) {
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
    return !this._directDepsUpdated && this._hasRegisteredFrontierDependencies;
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


  public upgradeFrontierDependencies() {
    // todo can be optimized to not generate the frontier candidates every time
    // we can cache the frontier candidates for the current cell
    const frontierCandidates: CellEvalNode[] = Array.from(
      this.frontierDependencies
    );

    for (const candidateNode of frontierCandidates) {
      const result = candidateNode.evaluationResult;

      // upgrade or downgrade frontier dependency
      if (result) {
        if (result.type === "spilled-values") {
          const spillArea = result.spillArea(candidateNode.cellAddress);
          const intersects = checkRangeIntersection(this.frontierRange.range, spillArea);
          if (intersects) {
            this.maybeUpgradeFrontierDependency(candidateNode); // upgraded!
          } else {
            this.maybeDiscardFrontierDependency(candidateNode); // downgraded!
          }
        } else {
          this.maybeDiscardFrontierDependency(candidateNode); // downgraded!
        }
      }
    }
  }
}
