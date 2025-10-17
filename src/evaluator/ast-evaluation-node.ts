import type { ASTNode } from "src/parser/ast";
import { astToString } from "src/parser/formatter";
import { BaseEvalNode } from "./cell-eval-node";
import type { CellAddress } from "src/core/types";
import type { ContextDependency } from "./evaluation-context";

export class AstEvaluationNode extends BaseEvalNode {
  private _contextDependency: ContextDependency;
  constructor(
    public ast: ASTNode,
    contextDependency: ContextDependency
  ) {
    const key = `ast:${astToString(ast)}`;
    const dummyAddress: CellAddress = {
      workbookName: "dummy",
      sheetName: "dummy",
      colIndex: 0,
      rowIndex: 0,
    };
    super(key, {
      type: "awaiting-evaluation",
      waitingFor: dummyAddress,
      errAddress: dummyAddress,
    });
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
