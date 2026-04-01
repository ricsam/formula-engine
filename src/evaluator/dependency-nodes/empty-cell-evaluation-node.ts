import type { DependencyManager } from "../../core/managers/dependency-manager";
import { FrontierDependencyManager } from "../../core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "../../core/managers/workbook-manager";
import type {
  CellAddress,
  RangeAddress,
  SingleEvaluationResult,
} from "../../core/types";
import { getCellReference, keyToCellAddress } from "../../core/utils";

export class EmptyCellEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public cellAddress: CellAddress;
  private _evaluationResult: SingleEvaluationResult;

  constructor(
    public emptyCellKey: string,
    evaluationManager: DependencyManager,
    workbookManager: WorkbookManager,
    options?: { skipInitialBuild?: boolean }
  ) {
    const cellAddress = keyToCellAddress(emptyCellKey);
    const emptyCellRange: RangeAddress = {
      range: {
        start: {
          col: cellAddress.colIndex,
          row: cellAddress.rowIndex,
        },
        end: {
          col: { type: "number", value: cellAddress.colIndex },
          row: { type: "number", value: cellAddress.rowIndex },
        },
      },
      sheetName: cellAddress.sheetName,
      workbookName: cellAddress.workbookName,
    };

    super(emptyCellRange, workbookManager, evaluationManager, options);

    this.cellAddress = cellAddress;
    this.key = emptyCellKey.replace(/^cell:/, "empty:");
    this._evaluationResult = {
      type: "awaiting-evaluation",
      waitingFor: this,
      errAddress: this,
    };
  }

  public setEvaluationResult(result: SingleEvaluationResult) {
    this._evaluationResult = result;
  }

  public get evaluationResult(): SingleEvaluationResult {
    return this._evaluationResult;
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

  public override restoreResolvedSnapshot(options: {
    dependencies: Set<import("../../core/managers/dependency-node").DependencyNode>;
    evaluationResult: SingleEvaluationResult;
  }) {
    super.restoreResolvedSnapshot({
      dependencies: options.dependencies,
    });
    this._evaluationResult = options.evaluationResult;
  }

  public override invalidate() {
    super.invalidate();
    this._evaluationResult = {
      type: "awaiting-evaluation",
      waitingFor: this,
      errAddress: this,
    };
  }

  toJSON(visitor: Set<string> = new Set()): any {
    const hasVisited = visitor?.has(this.key);
    if (hasVisited) {
      return {
        key: this.key,
        resolved: this.resolved,
        cycle: true,
        dependencies: [],
        frontierDependencies: [],
      };
    }
    visitor?.add(this.key);
    return {
      key: this.key,
      resolved: this.resolved,
      evaluationResult: this.evaluationResult,
      dependencies: Array.from(this.getDependencies()).map((node) =>
        node.toJSON(visitor)
      ),
      frontierDependencies: Array.from(this.getFrontierDependencies()).map(
        (node) => node.toJSON(visitor)
      ),
    };
  }

  public override toString(): string {
    return getCellReference(this.cellAddress);
  }
}
