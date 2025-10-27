import {
  type CellAddress,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../../../core/types";
import type { ParsedCriteria } from "../../criteria-parser";

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
    if (result.type === "error" || result.type === "awaiting-evaluation") {
      // Propagate errors immediately
      return result;
    }

    if (result.type === "value") {
      if (
        result.result.type === "number" ||
        result.result.type === "infinity"
      ) {
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

export function countEmptyCells(
  criteriaRangeResult: FunctionEvaluationResult,
  parsedCriteria: ParsedCriteria,
  currentCell: CellAddress
): FunctionEvaluationResult | undefined {
  // Special case: counting empty cells over infinite ranges
  if (
    criteriaRangeResult.type === "spilled-values" &&
    parsedCriteria.type === "exact" &&
    parsedCriteria.value.type === "string" &&
    parsedCriteria.value.value === ""
  ) {
    const spillArea = criteriaRangeResult.spillArea(currentCell);

    // Check if this is an infinite range
    if (
      spillArea.end.col.type === "infinity" ||
      spillArea.end.row.type === "infinity"
    ) {
      // Return infinity for infinite empty cell count
      return {
        type: "value",
        result: { type: "infinity", sign: "positive" },
      };
    }
  }
}
