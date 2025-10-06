import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type {
  CellAddress,
  FunctionEvaluationResult,
  RangeAddress,
  SingleEvaluationResult,
  SpreadsheetRange,
} from "src/core/types";
import {
  keyToCellAddress,
  keyToRangeAddress,
  rangeAddressToKey,
} from "src/core/utils";
import type { DependencyManager } from "src/core/managers/dependency-manager";

export class EmptyCellEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public cellAddress: CellAddress;

  private _evaluationResult?: FunctionEvaluationResult;

  constructor(
    public emptyCellKey: string,
    evaluationManager: DependencyManager,
    workbookManager: WorkbookManager
  ) {
    const cellAddress = keyToCellAddress(emptyCellKey);
    super(workbookManager, evaluationManager);

    this.cellAddress = cellAddress;
    this.key = emptyCellKey.replace(/^cell:/, "empty:");
  }

  public setEvaluationResult(result: FunctionEvaluationResult) {
    this._evaluationResult = result;
  }

  public get evaluationResult() {
    return (
      this._evaluationResult ?? {
        type: "awaiting-evaluation",
        cellAddress: this.cellAddress,
      }
    );
  }

  public override resolve() {
    if (this.canResolve()) {
      super.resolve();
    }
  }

  public override canResolve() {
    return (
      super.canResolve() && this.evaluationResult.type !== "awaiting-evaluation"
    );
  }

  /**
   * An origin spill result is not possible for an empty cell
   * but we just have it here for being consistent with the DependencyNode
   */
  public get originSpillResult(): SingleEvaluationResult | undefined {
    return undefined;
  }
}
