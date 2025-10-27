import { AstEvaluationNode } from "../../evaluator/dependency-nodes/ast-evaluation-node";
import { CellValueNode } from "../../evaluator/dependency-nodes/cell-value-node";
import { EmptyCellEvaluationNode } from "../../evaluator/dependency-nodes/empty-cell-evaluation-node";
import { SpillMetaNode } from "../../evaluator/dependency-nodes/spill-meta-node";
import { RangeEvaluationNode } from "../../evaluator/range-evaluation-node";
import type { DependencyNode } from "./dependency-node";


export function isDependencyNode(node: any): node is DependencyNode {
  return (
    node instanceof CellValueNode ||
    node instanceof RangeEvaluationNode ||
    node instanceof EmptyCellEvaluationNode ||
    node instanceof AstEvaluationNode ||
    node instanceof SpillMetaNode
  );
}
