import type {
  CellAddress,
  FunctionEvaluationResult,
  SingleEvaluationResult,
} from "src/core/types";
import { getCellReference, keyToCellAddress } from "src/core/utils";
import { RangeEvaluationNode } from "./range-evaluation-node";
import { EmptyCellEvaluationNode } from "./empty-cell-evaluation-node";
import type { AstEvaluationNode } from "./ast-evaluation-node";

export class BaseEvalNode {
  public readonly key: string;
  private _dependencies: Set<
    | CellEvalNode
    | RangeEvaluationNode
    | EmptyCellEvaluationNode
    | AstEvaluationNode
  > = new Set();
  private _evaluationResult: FunctionEvaluationResult;
  private _originSpillResult?: SingleEvaluationResult;
  private _directDepsUpdated: boolean = false;
  private _resolved: boolean = false;

  constructor(key: string, initialEvaluationResult: FunctionEvaluationResult) {
    this.key = key;
    this._evaluationResult = initialEvaluationResult;
  }

  public addDependency(
    dep:
      | CellEvalNode
      | RangeEvaluationNode
      | EmptyCellEvaluationNode
      | AstEvaluationNode
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
    return (
      this.evaluationResult.type !== "awaiting-evaluation" &&
      !this._directDepsUpdated
    );
  }

  public get resolved() {
    return this._resolved;
  }

  public get evaluationResult(): FunctionEvaluationResult {
    return this._evaluationResult;
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
      };
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

export class CellEvalNode extends BaseEvalNode {
  public readonly cellAddress: CellAddress;
  constructor(key: string) {
    const cellAddress = keyToCellAddress(key);
    super(key, {
      type: "awaiting-evaluation",
      waitingFor: cellAddress,
      errAddress: cellAddress,
    });
    this.cellAddress = cellAddress;
  }

  public override toString(): string {
    return getCellReference(this.cellAddress);
  }
}
