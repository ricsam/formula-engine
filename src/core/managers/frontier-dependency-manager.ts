import { CellEvalNode } from "src/evaluator/cell-eval-node";
import type {
  CellAddress,
  EvaluateAllCellsResult,
  RangeAddress,
} from "../types";
import {
  cellAddressToKey,
  checkRangeIntersection,
  isCellInRange,
} from "../utils";
import type { WorkbookManager } from "./workbook-manager";
import type { DependencyManager } from "./dependency-manager";
import type { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import { EmptyCellEvaluationNode } from "src/evaluator/empty-cell-evaluation-node";
import type { RangeEvalOrderEntry } from "./range-eval-order-builder";

type EvalOrderEntry =
  | {
      type: "value";
      address: CellAddress;
      node: CellEvalNode;
    }
  | {
      type: "empty_cell";
      address: CellAddress;
      candidates: CellEvalNode[];
    }
  | {
      type: "empty_range";
      address: RangeAddress;
      candidates: CellEvalNode[];
    };

function assertIsCellEvalNode(
  node: CellEvalNode | EmptyCellEvaluationNode
): asserts node is CellEvalNode {
  if (node instanceof EmptyCellEvaluationNode) {
    throw new Error("A frontier dependencies can not be an empty cell");
  }
}

export class FrontierDependencyManager {
  private evalOrder: EvalOrderEntry[];
  constructor(
    private frontierRange: RangeAddress,
    protected workbookManager: WorkbookManager,
    protected evaluationManager: DependencyManager
  ) {
    const addressToNode = (address: CellAddress) => {
      const node = this.evaluationManager.getCellNode(
        cellAddressToKey(address)
      );
      assertIsCellEvalNode(node);
      return node;
    };
    // todo maybe pass in lookupOrder
    this.evalOrder = this.workbookManager
      .buildRangeEvalOrder("col-major", this.frontierRange)
      .map((entry): EvalOrderEntry => {
        if (entry.type === "value") {
          return {
            type: "value",
            address: entry.address,
            node: addressToNode(entry.address),
          };
        } else if (entry.type === "empty_cell") {
          return {
            type: "empty_cell",
            address: entry.address,
            candidates: entry.candidates.map(addressToNode),
          };
        } else if (entry.type === "empty_range") {
          return {
            type: "empty_range",
            address: entry.address,
            candidates: entry.candidates.map(addressToNode),
          };
        }
        throw new Error("Invalid entry type: " + (entry as any).type);
      });
    for (const entry of this.evalOrder) {
      if (entry.type === "empty_cell" || entry.type === "empty_range") {
        for (const candidate of entry.candidates) {
          this.addFrontierDependency(candidate);
        }
      } else if (entry.type === "value") {
        this.addDependency(entry.node);
      }
    }
  }

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

  public getRangeEvalOrder() {
    return this.evalOrder;
  }

  public get frontierDependencies() {
    return this._frontierDependencies
      .difference(this._discardedFrontierDependencies)
      .difference(this._dependencies);
  }

  public get discardedFrontierDependencies() {
    return this._discardedFrontierDependencies;
  }

  protected addFrontierDependency(dependency: CellEvalNode) {
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
    return !this._directDepsUpdated && this.frontierDependencies.size === 0;
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
          const intersects = checkRangeIntersection(
            this.frontierRange.range,
            spillArea
          );
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
