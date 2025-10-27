import type {
  CellAddress,
  DoesNotSpillResult,
  ErrorEvaluationResult,
  SpilledValuesEvaluationResult,
} from "../../core/types";
import { getCellReference, keyToCellAddress } from "../../core/utils";
import { BaseEvalNode } from "./base-eval-node";

export class SpillMetaNode extends BaseEvalNode<
  SpilledValuesEvaluationResult | ErrorEvaluationResult | DoesNotSpillResult
> {
  public readonly cellAddress: CellAddress;

  constructor(key: string) {
    const cellAddress = keyToCellAddress(key);
    super(key);
    this.cellAddress = cellAddress;
  }

  public override toString(): string {
    return getCellReference(this.cellAddress);
  }
}
