import { AstEvaluationNode } from "../../evaluator/dependency-nodes/ast-evaluation-node";
import { CellValueNode } from "../../evaluator/dependency-nodes/cell-value-node";
import { EmptyCellEvaluationNode } from "../../evaluator/dependency-nodes/empty-cell-evaluation-node";
import { ResourceDependencyNode } from "../../evaluator/dependency-nodes/resource-dependency-node";
import { SpillMetaNode } from "../../evaluator/dependency-nodes/spill-meta-node";
import { RangeEvaluationNode } from "../../evaluator/range-evaluation-node";

export type DependencyNode =
  | CellValueNode
  | RangeEvaluationNode
  | EmptyCellEvaluationNode
  | AstEvaluationNode
  | SpillMetaNode
  | ResourceDependencyNode;

export type CellNodeType = "cell-value" | "empty" | "spill-meta";

export type CellNodeKey = `${CellNodeType}:${string}`;
export type CellNodeKeyDictionary = {
  "cell-value": CellValueNode;
  empty: EmptyCellEvaluationNode;
  "spill-meta": SpillMetaNode;
};
