import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { RangeAddress, SpreadsheetRange } from "src/core/types";
import { keyToRangeAddress, rangeAddressToKey } from "src/core/utils";
import type { DependencyManager } from "src/core/managers/dependency-manager";
import type { CacheManager } from "src/core/managers/cache-manager";
import type { LookupOrder } from "src/core/managers";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

  constructor(
    public rangeKey: string,
    private cacheManager: CacheManager,
    dependencyManager: DependencyManager,
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
    return rangeEvalOrder;
  }
}
