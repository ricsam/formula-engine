import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { RangeAddress, SpreadsheetRange } from "src/core/types";
import { keyToRangeAddress, rangeAddressToKey } from "src/core/utils";
import type { DependencyManager } from "src/core/managers/dependency-manager";
import type { CacheManager } from "src/core/managers/cache-manager";

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

  public getCellsInRange() {
    const cachedCellsInRange = this.cacheManager.getCellsInRange(this.key);
    if (cachedCellsInRange) {
      return cachedCellsInRange;
    }
    const cellsInRange = this.workbookManager.getCellsInRange(this.address);
    this.cacheManager.setCellsInRange(this.key, cellsInRange);
    return cellsInRange;
  }

  public getFrontierCandidates() {
    const cachedFrontierCandidates = this.cacheManager.getFrontierCandidates(
      this.key
    );
    if (cachedFrontierCandidates) {
      return cachedFrontierCandidates;
    }
    const frontierCandidates = this.workbookManager.getFrontierCandidates(
      this.address
    );
    this.cacheManager.setFrontierCandidates(this.key, frontierCandidates);
    return frontierCandidates;
  }
}
