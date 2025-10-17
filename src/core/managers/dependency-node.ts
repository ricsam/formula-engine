import type { CellEvalNode } from "src/evaluator/cell-eval-node";
import type { EmptyCellEvaluationNode } from "src/evaluator/empty-cell-evaluation-node";
import type { RangeEvaluationNode } from "src/evaluator/range-evaluation-node";
import type { AstEvaluationNode } from "src/evaluator/ast-evaluation-node";

export type DependencyNode =
  | CellEvalNode
  | RangeEvaluationNode
  | EmptyCellEvaluationNode
  | AstEvaluationNode;
