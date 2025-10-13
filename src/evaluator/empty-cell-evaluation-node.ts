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
  cellAddressToKey,
  isCellInRange,
  keyToCellAddress,
  keyToRangeAddress,
  rangeAddressToKey,
} from "src/core/utils";
import type { DependencyManager } from "src/core/managers/dependency-manager";
import type { CellEvalNode } from "./cell-eval-node";

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

    super(emptyCellRange, workbookManager, evaluationManager);

    this.cellAddress = cellAddress;
    this.key = emptyCellKey.replace(/^cell:/, "empty:");
  }

  public setEvaluationResult(result: FunctionEvaluationResult) {
    this._evaluationResult = result;
  }

  public get evaluationResult(): FunctionEvaluationResult {
    return (
      this._evaluationResult ?? {
        type: "awaiting-evaluation",
        waitingFor: this.cellAddress,
        errAddress: this.cellAddress,
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
}
