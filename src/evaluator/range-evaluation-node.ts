import type { DependencyManager } from "src/core/managers/dependency-manager";
import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { EvaluateAllCellsResult, RangeAddress } from "src/core/types";
import { getRangeKey, keyToRangeAddress } from "src/core/utils";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

  private _results: EvaluateAllCellsResult[] | undefined;

  constructor(
    public rangeKey: string,
    dependencyManager: DependencyManager,
    workbookManager: WorkbookManager
  ) {
    const rangeAddress = keyToRangeAddress(rangeKey);
    super(rangeAddress, workbookManager, dependencyManager);

    this.address = rangeAddress;
    this.key = rangeKey;
  }

  setResults(results: EvaluateAllCellsResult[]): void {
    if (!this.resolved) {
      throw new Error(
        "Cannot set results on an unresolved range evaluation node"
      );
    }
    this._results = results;
  }

  // todo maybe add lookupOrder
  getResults(): EvaluateAllCellsResult[] | undefined {
    return this._results;
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
