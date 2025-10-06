import type { CellAddress } from "src/core/types";
import type { CellEvalNode } from "./cell-eval-node";
import type { RangeEvaluationNode } from "./range-evaluation-node";
import type { EmptyCellEvaluationNode } from "./empty-cell-evaluation-node";

export class EvaluationContext {
  /**
   * Can be a range or a cell
   */
  private _dependencyNode: CellEvalNode | RangeEvaluationNode | EmptyCellEvaluationNode;
  /**
   * The cell evaluating a cell,e.g.
   * if we are evaluting A1=SUM(B2:B4) + B1, then the origin cell is A1 and the dependency node is A1 as well
   * the open range evaluator will create a new context with the origin cell being A1 and the dependency node being B2:B4
   * 
   * A new dependency will be added to A1 onto B1, and then B1 will be evaluated just like A1 is evaluated where the origin cell is B1
   */
  private _originCell: CellEvalNode | EmptyCellEvaluationNode;

  constructor(dependencyNode: CellEvalNode | RangeEvaluationNode | EmptyCellEvaluationNode, originCell: CellEvalNode | EmptyCellEvaluationNode) {
    this._dependencyNode = dependencyNode;
    this._originCell = originCell;
  }

  get dependencyNode() {
    return this._dependencyNode;
  }

  get originCell() {
    return this._originCell;
  }
}
