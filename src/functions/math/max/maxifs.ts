import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import {
  parseCriteriaPairs,
  processMultiCriteriaValues,
  validateMultiCriteriaArgs,
} from "../../criteria-utils";
import { performMaximum } from "./max-utils";

/**
 * MAXIFS function - Returns the maximum value among cells specified by multiple criteria
 *
 * Usage: MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 *
 * max_range: The range of cells from which you want the maximum
 * criteria_range1: The first range in which to evaluate criteria
 * criteria1: The criteria to apply to criteria_range1
 * criteria_range2, criteria2: Optional additional criteria pairs
 *
 * Examples:
 *   MAXIFS(B1:B10, A1:A10, "Apple") - max of B1:B10 where A1:A10 = "Apple"
 *   MAXIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - max of C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 *
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are considered for the maximum
 * - Returns error if no values match all criteria
 */
export const MAXIFS: FunctionDefinition = {
  name: "MAXIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments using utility
    const argError = validateMultiCriteriaArgs("MAXIFS", node.args.length);
    if (argError) {
      return argError;
    }

    // Evaluate max range
    const maxRangeResult = this.evaluateNode(node.args[0]!, context);
    if (maxRangeResult.type === "error") {
      return maxRangeResult;
    }

    // Parse criteria pairs using utility
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
    const criteriaPairs = criteriaPairsResult;

    // Use shared maximum utility for standard cases
    const matchingValues = processMultiCriteriaValues(
      this,
      maxRangeResult,
      criteriaPairs,
      context,
      "col-major"
    );

    return performMaximum(matchingValues);
  },
};
