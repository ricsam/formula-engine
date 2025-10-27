import {
  type FunctionDefinition,
  type FunctionEvaluationResult
} from "../../../core/types";
import {
  parseCriteriaPairs,
  processMultiCriteriaValues,
  validateMultiCriteriaArgs,
} from "../../criteria-utils";
import { performSummation } from "./summation-utils";

/**
 * SUMIFS function - Sums cells in a range that meet multiple criteria
 *
 * Usage: SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 *
 * sum_range: The range of cells to sum
 * criteria_range1: The first range to evaluate against criteria1
 * criteria1: The first criteria to match against
 * criteria_range2, criteria2: Optional additional criteria pairs
 *
 * Examples:
 *   SUMIFS(B1:B10, A1:A10, "Apple") - sums B1:B10 where A1:A10 = "Apple"
 *   SUMIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - sums C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 *
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are included in the sum
 * - Returns 0 if no values match all criteria (unlike other IFS functions that return error)
 */
export const SUMIFS: FunctionDefinition = {
  name: "SUMIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments
    const argError = validateMultiCriteriaArgs("SUMIFS", node.args.length, context);
    if (argError) {
      return argError;
    }

    // Evaluate sum range (first argument)
    const sumRangeResult = this.evaluateNode(node.args[0]!, context);
    if (sumRangeResult.type === "error") {
      return sumRangeResult;
    }

    // Parse criteria pairs starting from argument 1
    const criteriaPairsResult = parseCriteriaPairs.call(
      this,
      node,
      context,
      this.evaluateNode,
      1
    );
    if ("type" in criteriaPairsResult) {
      return criteriaPairsResult;
    }

    // Use shared summation utility for standard cases
    const matchingValues = processMultiCriteriaValues(
      this,
      sumRangeResult,
      criteriaPairsResult,
      context,
      "col-major"
    );

    if (matchingValues.type === "error" || matchingValues.type === "awaiting-evaluation") {
      return matchingValues;
    }

    return performSummation(matchingValues.values);
  },
};
