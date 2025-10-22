import { type DependencyNode } from "src/core/managers/dependency-node";
import type {
  DoesNotSpillResult,
  FunctionEvaluationResult,
  SpilledValuesEvaluationResult,
} from "src/core/types";
import type { SpillMetaNode } from "./spill-meta-node";

export class BaseEvalNode<T> {
  public key: string;
  private _dependencies: Set<DependencyNode> = new Set();
  private _lastDependencies: Set<DependencyNode> = new Set();
  public _evaluationResult: T;
  private _resolved: boolean = false;

  constructor(key: string) {
    this.key = key;
    this._evaluationResult = {
      type: "awaiting-evaluation",
      waitingFor: this,
      errAddress: this,
    } as any;
  }

  public addDependency(dep: DependencyNode) {
    if (this._dependencies.has(dep)) {
      return;
    }
    this._dependencies.add(dep);
  }

  public get directDepsUpdated() {
    return (
      !this._lastDependencies.isSubsetOf(this._dependencies) ||
      this._dependencies.size !== this._lastDependencies.size
    );
  }

  public resolve() {
    if (this.canResolve()) {
      this._resolved = true;
    }
  }

  public canResolve() {
    return (
      (this.evaluationResult as any).type !== "awaiting-evaluation" &&
      !this.directDepsUpdated
    );
  }

  public get resolved() {
    return this._resolved;
  }

  public get evaluationResult(): T {
    return this._evaluationResult;
  }

  public setEvaluationResult(result: T) {
    this._evaluationResult = result;
  }

  public resetDirectDepsUpdated() {
    if (this._resolved) {
      return;
    }
    this._lastDependencies = new Set(this._dependencies);
    this._dependencies = new Set();
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
  public getFrontierDependencies(): Set<SpillMetaNode> {
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
      dependencies: Array.from(this.getDependencies()).map((node) =>
        node.toJSON(visitor)
      ),
    };
  }
}
