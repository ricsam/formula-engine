import {
  FormulaError,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type CellInfinity,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";

/**
 * Result type for processInfinity in max functions
 */
export type ProcessInfinityResult<T> =
  | { type: "infinity"; infinity: CellInfinity }
  | { type: "state"; state: T };

/**
 * Perform maximum calculation over an iterator of SingleEvaluationResults
 * Handles the common logic shared by MAX, MAXIF, MAXIFS functions
 *
 * @param results - Iterator of SingleEvaluationResults to find maximum from
 * @returns FunctionEvaluationResult with the maximum, infinity, or error
 */
export function performMaximum(
  results: Iterable<SingleEvaluationResult>,
  context: EvaluationContext
): FunctionEvaluationResult {
  let maxValue = -Infinity;
  let hasValues = false;

  for (const result of results) {
    if (result.type === "error") {
      // Propagate errors immediately
      return result;
    }

    if (result.type === "value") {
      if (result.result.type === "number") {
        maxValue = Math.max(maxValue, result.result.value);
        hasValues = true;
      } else if (result.result.type === "infinity") {
        if (result.result.sign === "positive") {
          // Positive infinity is always the maximum - return immediately
          return {
            type: "value",
            result: result.result,
          };
        }
        // Negative infinity doesn't change the maximum
      }
      // Non-numeric values (strings, booleans) are ignored in maximum calculation
    }
  }

  if (!hasValues) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "No numeric values found",
      errAddress: context.originCell.cellAddress,
    };
  }

  return {
    type: "value",
    result: { type: "number", value: maxValue },
  };
}

