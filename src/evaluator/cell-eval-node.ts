import type {
  CellAddress,
  FunctionEvaluationResult,
  SingleEvaluationResult,
} from "src/core/types";
import { keyToCellAddress } from "src/core/utils";
import { RangeEvaluationNode } from "./range-evaluation-node";
import { EmptyCellEvaluationNode } from "./empty-cell-evaluation-node";

export class CellEvalNode {
  public readonly cellAddress: CellAddress;
  public readonly key: string;
  private _dependencies: Set<
    CellEvalNode | RangeEvaluationNode | EmptyCellEvaluationNode
  > = new Set();
  private _evaluationResult?: FunctionEvaluationResult;
  private _originSpillResult?: SingleEvaluationResult;
  private _directDepsUpdated: boolean = false;
  private _resolved: boolean = false;

  constructor(
    key: string,
    evaluationResult?: FunctionEvaluationResult,
    originSpillResult?: SingleEvaluationResult
  ) {
    this.cellAddress = keyToCellAddress(key);
    this.key = key;
    this._evaluationResult = evaluationResult;
    this._originSpillResult = originSpillResult;
  }

  public addDependency(
    dep: CellEvalNode | RangeEvaluationNode | EmptyCellEvaluationNode
  ) {
    if (this._dependencies.has(dep)) {
      return;
    }
    this._directDepsUpdated = true;
    this._dependencies.add(dep);
  }

  public get directDepsUpdated() {
    return this._directDepsUpdated;
  }

  public resolve() {
    if (this.canResolve()) {
      this._resolved = true;
    }
  }

  public canResolve() {
    return this.evaluationResult.type !== "awaiting-evaluation" && !this._directDepsUpdated;
  }

  public get resolved() {
    return this._resolved;
  }

  public get evaluationResult(): FunctionEvaluationResult {
    return (
      this._evaluationResult ?? {
        type: "awaiting-evaluation",
        waitingFor: this.cellAddress,
        errAddress: this.cellAddress,
      }
    );
  }

  public get originSpillResult() {
    return this._originSpillResult;
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

  public resetDirectDepsUpdated() {
    if (this._resolved) {
      return;
    }
    this._directDepsUpdated = false;
  }

  /**
   * Get the direct dependencies of the node, either RangeEvaluationNode or DependencyNode
   */
  public getDependencies() {
    return this._dependencies;
  }

  /**
   * Just to mirror the method in RangeEvaluationNode
   */
  public getAllDependencies() {
    return this.getDependencies();
  }

  /**
   * Just to mirror the method in RangeEvaluationNode
   */
  public getFrontierDependencies(): Set<CellEvalNode> {
    return new Set();
  }

  toJSON(visitor: Set<string> = new Set()): any {
    const hasVisited = visitor?.has(this.key);
    if (hasVisited) {
      return {
        key: this.key,
        resolved: this.resolved,
        cycle: true,
        dependencies: [],
      }
    }
    visitor?.add(this.key);
    return {
      key: this.key,
      resolved: this.resolved,
      evaluationResult: this.evaluationResult,
      originSpillResult: this.originSpillResult,
      dependencies: Array.from(this.getDependencies()).map((node) =>
        node.toJSON(visitor)
      ),
    };
  }
}
