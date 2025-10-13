import type { DependencyManager } from "src/core/managers/dependency-manager";
import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { RangeAddress } from "src/core/types";
import { keyToRangeAddress } from "src/core/utils";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

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
}
