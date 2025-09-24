import {
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";

/**
 * Performs counting of numeric values from an iterable of SingleEvaluationResult values
 * Handles numbers and infinities according to Excel's COUNT behavior
 * - Propagates errors immediately
 * - Only counts numeric values (numbers and infinities)
 * - Ignores non-numeric values (strings, booleans, empty cells)
 *
 * @param results - Iterable of SingleEvaluationResult to count
 * @returns FunctionEvaluationResult with the count
 */
export function performCount(
  results: Iterable<SingleEvaluationResult>
): FunctionEvaluationResult {
  let count = 0;

  for (const result of results) {
    if (result.type === "error") {
      // Propagate errors immediately
      return result;
    }

    if (result.type === "value") {
      if (result.result.type === "number" || result.result.type === "infinity") {
        count++;
      }
      // Non-numeric values (strings, booleans) are ignored
    }
  }

  return {
    type: "value",
    result: { type: "number", value: count },
  };
}
