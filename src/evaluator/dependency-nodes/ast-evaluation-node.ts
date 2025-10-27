import type { FunctionEvaluationResult } from "../../core/types";
import type { ASTNode } from "../../parser/ast";
import { astToString } from "../../parser/formatter";
import type { ContextDependency } from "../evaluation-context";
import { BaseEvalNode } from "./base-eval-node";

export class AstEvaluationNode extends BaseEvalNode<FunctionEvaluationResult> {
  private _contextDependency: ContextDependency;
  constructor(
    public ast: ASTNode,
    contextDependency: ContextDependency
  ) {
    const key = `ast:${astToString(ast)}`;
    super(key);
    this._contextDependency = contextDependency;
  }

  public override toString(): string {
    return this.key;
  }

  getContextDependency() {
    return this._contextDependency;
  }

  setContextDependency(contextDependency: ContextDependency) {
    this._contextDependency = contextDependency;
  }
}
