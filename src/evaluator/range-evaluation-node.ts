import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { RangeAddress, SpreadsheetRange } from "src/core/types";
import {
  cellAddressToKey,
  keyToRangeAddress,
  rangeAddressToKey,
} from "src/core/utils";
import type { DependencyManager } from "src/core/managers/dependency-manager";
import type { CacheManager } from "src/core/managers/cache-manager";
import type { LookupOrder } from "src/core/managers";
import { EmptyCellEvaluationNode } from "./empty-cell-evaluation-node";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

  constructor(
    public rangeKey: string,
    private cacheManager: CacheManager,
    private dependencyManager: DependencyManager,
    workbookManager: WorkbookManager
  ) {
    const rangeAddress = keyToRangeAddress(rangeKey);
    super(workbookManager, dependencyManager);

    this.address = rangeAddress;
    this.key = rangeKey;
  }

  public getRangeEvalOrder(lookupOrder: LookupOrder) {
    const cacheKey = this.key + "@" + lookupOrder;
    const cachedRangeEvalOrder = this.cacheManager.getRangeEvalOrder(cacheKey);
    if (cachedRangeEvalOrder) {
      return cachedRangeEvalOrder;
    }
    const rangeEvalOrder = this.workbookManager.buildRangeEvalOrder(
      lookupOrder,
      this.address
    );
    this.cacheManager.setRangeEvalOrder(cacheKey, rangeEvalOrder);
    // set up the dependencies for the range node
    for (const entry of rangeEvalOrder) {
      if (entry.type === "value") {
        const cellKey = cellAddressToKey(entry.address);
        const cellNode = this.dependencyManager.getCellNode(cellKey);
        this.addDependency(cellNode);
      } else if (entry.type === "empty_cell" || entry.type === "empty_range") {
        for (const candidate of entry.candidates) {
          const candidateKey = cellAddressToKey(candidate);
          const candidateNode =
            this.dependencyManager.getCellNode(candidateKey);
          if (candidateNode instanceof EmptyCellEvaluationNode) {
            throw new Error("A frontier dependency can not be an empty cell");
          }
          this.addFrontierDependency(candidateNode);
        }
      }
    }
    return rangeEvalOrder;
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
      }
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
