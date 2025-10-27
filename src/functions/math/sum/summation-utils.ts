import {
  FormulaError,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type CellInfinity,
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { FunctionNode } from "../../../parser/ast";

/**
 * Result type for processInfinity in sum functions
 */
export type ProcessInfinityResult<T> =
  | { type: "infinity"; infinity: CellInfinity }
  | { type: "state"; state: T };

/**
 * Perform summation with error propagation
 * Used when errors should be propagated instead of skipped (e.g., SUM with direct error arguments)
 *
 * @param results - Iterator of SingleEvaluationResults to sum
 * @returns FunctionEvaluationResult with the sum, infinity, or first error encountered
 */
export function performSummation(
  results: Iterable<SingleEvaluationResult>
): FunctionEvaluationResult {
  let sum = 0;
  let hasValues = false;

  for (const result of results) {
    if (result.type === "value") {
      if (result.result.type === "number") {
        sum += result.result.value;
        hasValues = true;
      } else if (result.result.type === "infinity") {
        // Infinity dominates - return immediately
        return {
          type: "value",
          result: result.result,
        };
      } else {
        // Non-numeric values (strings, booleans) are ignored in summation
      }
    }
  }

  return {
    type: "value",
    result: { type: "number", value: sum },
  };
}
