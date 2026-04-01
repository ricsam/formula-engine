import type {
  AwaitingEvaluationResult,
  CellAddress,
  ErrorEvaluationResult,
  SingleEvaluationResult,
  SpilledValuesEvaluationResult,
} from "../../core/types";
import { getCellReference, keyToCellAddress } from "../../core/utils";
import { BaseEvalNode } from "./base-eval-node";
import type { SpillMetaNode } from "./spill-meta-node";

export class CellValueNode extends BaseEvalNode<SingleEvaluationResult> {
  public readonly cellAddress: CellAddress;

  constructor(key: string) {
    const cellAddress = keyToCellAddress(key);
    super(key);

    this.cellAddress = cellAddress;
  }

  public override toString(): string {
    return getCellReference(this.cellAddress);
  }

  spillMeta: SpillMetaNode | undefined;

  setSpillMetaNode(node: SpillMetaNode) {
    this.spillMeta = node;
  }

  clearSpillMetaNode() {
    this.spillMeta = undefined;
  }
}
