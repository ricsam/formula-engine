import type { CellAddress, FormulaError } from "src/core/types";

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
    public readonly cellAddress: CellAddress,
    message: string
  ) {
    super(message);
  }
}
