import type {
  CellAddress,
  SerializedCellValue,
  SingleEvaluationResult,
} from "../../core/types";
import { getCellReference } from "../../core/utils";
import { BaseEvalNode } from "./base-eval-node";
import type { SpillMetaNode } from "./spill-meta-node";

export class VirtualCellValueNode extends BaseEvalNode<SingleEvaluationResult> {
  public readonly cellAddress: CellAddress;
  public readonly cellValue: SerializedCellValue;

  constructor(
    key: string,
    cellAddress: CellAddress,
    cellValue: SerializedCellValue
  ) {
    super(key);

    this.cellAddress = cellAddress;
    this.cellValue = cellValue;
  }

  public override toString(): string {
    return "virtual:" + getCellReference(this.cellAddress);
  }

  spillMeta: SpillMetaNode | undefined;

  setSpillMetaNode(node: SpillMetaNode) {
    this.spillMeta = node;
  }

  clearSpillMetaNode() {
    this.spillMeta = undefined;
  }
}
