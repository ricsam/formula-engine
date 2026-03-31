import type { DependencyManager } from "../core/managers/dependency-manager";
import { FrontierDependencyManager } from "../core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "../core/managers/workbook-manager";
import type { EvaluateAllCellsResult, RangeAddress } from "../core/types";
import { getRangeKey, keyToRangeAddress } from "../core/utils";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

  private _result: EvaluateAllCellsResult;

  constructor(
    public rangeKey: string,
    dependencyManager: DependencyManager,
    workbookManager: WorkbookManager,
    options?: { skipInitialBuild?: boolean }
  ) {
    const rangeAddress = keyToRangeAddress(rangeKey);
    super(rangeAddress, workbookManager, dependencyManager, options);

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

  public override restoreResolvedSnapshot(options: {
    dependencies: Set<import("../core/managers/dependency-node").DependencyNode>;
    result: EvaluateAllCellsResult;
  }) {
    super.restoreResolvedSnapshot({
      dependencies: options.dependencies,
    });
    this._result = options.result;
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
