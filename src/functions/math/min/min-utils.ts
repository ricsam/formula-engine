import {
  FormulaError,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type CellInfinity,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { FunctionNode } from "src/parser/ast";

/**
 * Result type for processInfinity in min functions
 */
export type ProcessInfinityResult<T> =
  | { type: "infinity"; infinity: CellInfinity }
  | { type: "state"; state: T };

/**
 * Perform minimum calculation over an iterator of SingleEvaluationResults
 * Handles the common logic shared by MIN, MINIF, MINIFS functions
 *
 * @param results - Iterator of SingleEvaluationResults to find minimum from
 * @returns FunctionEvaluationResult with the minimum, infinity, or error
 */
export function performMinimum(
  results: Iterable<SingleEvaluationResult>,
  context: EvaluationContext
): FunctionEvaluationResult {
  let minValue = Infinity;
  let hasValues = false;

  for (const result of results) {
    if (result.type === "error") {
      // Propagate errors immediately
      return result;
    }

    if (result.type === "value") {
      if (result.result.type === "number") {
        minValue = Math.min(minValue, result.result.value);
        hasValues = true;
      } else if (result.result.type === "infinity") {
        if (result.result.sign === "negative") {
          // Negative infinity is always the minimum - return immediately
          return {
            type: "value",
            result: result.result,
          };
        }
        // Positive infinity doesn't change the minimum
      }
      // Non-numeric values (strings, booleans) are ignored in minimum calculation
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
    result: { type: "number", value: minValue },
  };
}

