import { FrontierDependencyManager } from "src/core/managers/frontier-dependency-manager";
import type { WorkbookManager } from "src/core/managers/workbook-manager";
import type { RangeAddress, SpreadsheetRange } from "src/core/types";
import { keyToRangeAddress, rangeAddressToKey } from "src/core/utils";
import type { DependencyManager } from "src/core/managers/dependency-manager";

export class RangeEvaluationNode extends FrontierDependencyManager {
  public key: string;
  public address: RangeAddress;

  constructor(
    public rangeKey: string,
    evaluationManager: DependencyManager,
    workbookManager: WorkbookManager
  ) {
    const rangeAddress = keyToRangeAddress(rangeKey);
    super(workbookManager, evaluationManager);

    this.address = rangeAddress;
    this.key = rangeKey;
  }
}
