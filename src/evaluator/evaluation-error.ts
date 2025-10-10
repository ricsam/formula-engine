import type { CellAddress, FormulaError } from "src/core/types";
import { cellAddressToKey } from "src/core/utils";

export class EvaluationError extends Error {
  constructor(
    public readonly type: FormulaError,
    public readonly errAddress: CellAddress,
    message: string
  ) {
    super(message);
  }
}

export class AwaitingEvaluationError extends Error {
  constructor(
    public readonly errAddress: CellAddress,
    public readonly waitingFor: CellAddress
  ) {
    super(
      cellAddressToKey(errAddress) +
        " is awaiting evaluation of " +
        cellAddressToKey(waitingFor)
    );
  }
}
