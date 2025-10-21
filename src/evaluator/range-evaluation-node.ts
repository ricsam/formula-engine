import type { DependencyManager } from "src/core/managers/dependency-manager";
import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { EvaluateAllCellsResult, RangeAddress } from "src/core/types";
import { getRangeKey, keyToRangeAddress } from "src/core/utils";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

  private _result: EvaluateAllCellsResult;

  constructor(
    public rangeKey: string,
    dependencyManager: DependencyManager,
    workbookManager: WorkbookManager
  ) {
    const rangeAddress = keyToRangeAddress(rangeKey);
    super(rangeAddress, workbookManager, dependencyManager);

    this.address = rangeAddress;
    this.key = rangeKey;
    this._result = {
      type: "awaiting-evaluation",
      waitingFor: this,
      errAddress: this,
    };
  }

  setResult(result: EvaluateAllCellsResult): void {
    this._result = result;
  }

  public get result(): EvaluateAllCellsResult {
    return this._result;
  }

  public override canResolve(): boolean {
    return super.canResolve() && this._result.type !== "awaiting-evaluation";
  }

  toJSON(visitor: Set<string> = new Set()): any {
    const hasVisited = visitor?.has(this.key);
    visitor?.add(this.key);
    if (hasVisited) {
      return {
        key: this.key,
        resolved: this.resolved,
        cycle: true,
        dependencies: [],
        frontierDependencies: [],
      };
    }
    return {
      key: this.key,
      resolved: this.resolved,
      dependencies: Array.from(this.getDependencies()).map((node) =>
        node.toJSON(visitor)
      ),
      frontierDependencies: Array.from(this.getFrontierDependencies()).map(
        (node) => node.toJSON(visitor)
      ),
    };
  }

  public override toString(): string {
    return getRangeKey(this.address.range);
  }
}
