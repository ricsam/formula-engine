import type { DependencyNode } from "src/core/managers/dependency-node";
import type { CellAddress, FormulaError } from "src/core/types";
import { cellAddressToKey } from "src/core/utils";

export class EvaluationError extends Error {
  constructor(
    public readonly type: FormulaError,
    message: string,
    public readonly errAddress?: DependencyNode
  ) {
    super(message);
  }
}

export class AwaitingEvaluationError extends Error {
  constructor(
    public readonly errAddress: DependencyNode,
    public readonly waitingFor: DependencyNode
  ) {
    super(errAddress.key + " is awaiting evaluation of " + waitingFor.key);
  }
}

export class SheetNotFoundError extends Error {
  constructor(public readonly sheetName: string) {
    super("Sheet not found: " + sheetName);
  }
}

export class WorkbookNotFoundError extends Error {
  constructor(public readonly workbookName: string) {
    super("Workbook not found: " + workbookName);
  }
}

