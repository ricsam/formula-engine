import {
  FormulaError,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import { flags } from "src/debug/flags";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * Performs averaging of an iterable of SingleEvaluationResult values
 * Handles numbers, infinities, and errors according to Excel's AVERAGE behavior
 *
 * @param results - Iterable of SingleEvaluationResult to average
 * @returns FunctionEvaluationResult with the average or appropriate error
 */
export function performAverage(
  results: Iterable<SingleEvaluationResult>
): FunctionEvaluationResult {
  let sum = 0;
  let count = 0;

  for (const result of results) {
    if (result.type === "error") {
      // Propagate errors immediately
      return result;
    }
    if (result.type === "awaiting-evaluation") {
      return result;
    }

    if (result.type === "value") {
      if (result.result.type === "number") {
        sum += result.result.value;
        count++;
      } else if (result.result.type === "infinity") {
        // Infinity dominates - return immediately
        return {
          type: "value",
          result: result.result,
        };
      }
      // Non-numeric values (strings, booleans) are ignored
    }
  }

  if (count === 0) {
    return {
      type: "error",
      err: FormulaError.DIV0,
      message: "Cannot calculate average of empty range",
    };
  }

  return {
    type: "value",
    result: { type: "number", value: sum / count },
  };
}
