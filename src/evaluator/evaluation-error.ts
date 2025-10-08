import type { FormulaError } from "src/core/types";

export class EvaluationError extends Error {
  constructor(
    public readonly type: FormulaError,
    message: string
  ) {
    super(message);
  }
}
