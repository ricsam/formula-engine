import { CellValueNode } from "../../evaluator/dependency-nodes/cell-value-node";
import type { CellAddress, CellInRangeResult, RangeAddress } from "../types";
import { cellAddressToKey, checkRangeIntersection } from "../utils";
import type { DependencyManager } from "./dependency-manager";
import type { DependencyNode } from "./dependency-node";
import type { WorkbookManager } from "./workbook-manager";
import type { SpillMetaNode } from "../../evaluator/dependency-nodes/spill-meta-node";

type EvalOrderEntry =
  | {
      type: "value";
      address: CellAddress;
      node: CellValueNode;
    }
  | {
      type: "empty_cell";
      address: CellAddress;
      candidates: SpillMetaNode[];
    }
  | {
      type: "empty_range";
      address: RangeAddress;
      candidates: SpillMetaNode[];
    };

export class FrontierDependencyManager {
  private evalOrder: EvalOrderEntry[];
  private _resolved: boolean = false;
  private _directDepsUpdated: boolean;

  constructor(
    private frontierRange: RangeAddress,
    protected workbookManager: WorkbookManager,
    protected evaluationManager: DependencyManager
  ) {
    const addressToSpillMetaNode = (address: CellAddress) => {
      const node = this.evaluationManager.getSpillMetaNode(
        cellAddressToKey(address).replace(/^[^:]+:/, "spill-meta:")
      );
      return node;
    };
    // todo maybe pass in lookupOrder
    let directDepsUpdated = false;
    this.evalOrder = this.workbookManager
      .buildRangeEvalOrder("col-major", this.frontierRange)
      .map((entry): EvalOrderEntry => {
        if (entry.type === "value") {
          const addressToNode = (address: CellAddress) => {
            const node = this.evaluationManager.getCellValueNode(
              cellAddressToKey(address)
            );
            return node;
          };
          return {
            type: "value",
            address: entry.address,
            node: addressToNode(entry.address),
          };
        } else if (entry.type === "empty_cell") {
          return {
            type: "empty_cell",
            address: entry.address,
            candidates: entry.candidates.map(addressToSpillMetaNode),
          };
        } else if (entry.type === "empty_range") {
          return {
            type: "empty_range",
            address: entry.address,
            candidates: entry.candidates.map(addressToSpillMetaNode),
          };
        }
        throw new Error("Invalid entry type: " + (entry as any).type);
      });
    for (const entry of this.evalOrder) {
      if (entry.type === "empty_cell" || entry.type === "empty_range") {
        for (const candidate of entry.candidates) {
          this._frontierDependencies.add(candidate);
          directDepsUpdated = true;
        }
      } else if (entry.type === "value") {
        this.addDependency(entry.node);
        directDepsUpdated = true;
      }
    }
    this._directDepsUpdated = directDepsUpdated;
  }

  /**
   * frontierDependencies is the set of dependency node keys that could spill values onto the target range (if evaluationResult is spilled-values)
   * Key is from cellAddressToKey
   *
   * soft edge dependencies, which can not cause cycles in the dependency graph
   */
  private _frontierDependencies: Set<SpillMetaNode> = new Set();

  /**
   * discardedFrontierDependencies is the set of dependency node keys that were discarded as frontier dependencies because
   * they they do not produce spilled values that spill onto the target range
   * Key is from cellAddressToKey
   */
  private _discardedFrontierDependencies: Set<SpillMetaNode> = new Set();

  /**
   * hard edge dependencies, which can cause cycles in the dependency graph
   */
  private _dependencies: Set<DependencyNode> = new Set();

  /**
   * cache, should maybe be stored in the cache manager
   */
  get iterateAllCells(): undefined | Iterable<CellInRangeResult> {
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

  public maybeDiscardFrontierDependency(dependency: SpillMetaNode) {
    if (!this._resolved) {
      return;
    }
    if (this._discardedFrontierDependencies.has(dependency)) {
      return;
    }
    this._directDepsUpdated = true;
    this._discardedFrontierDependencies.add(dependency);
  }

  public maybeUpgradeFrontierDependency(dependency: SpillMetaNode) {
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

  public addDependency(dependency: DependencyNode) {
    this._dependencies.add(dependency);
  }

  public upgradeFrontierDependencies() {
    // todo can be optimized to not generate the frontier candidates every time
    // we can cache the frontier candidates for the current cell
    const frontierCandidates: SpillMetaNode[] = Array.from(
      this.frontierDependencies
    );

    for (const candidateNode of frontierCandidates) {
      const result = candidateNode.evaluationResult;

      // upgrade or downgrade frontier dependency
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
