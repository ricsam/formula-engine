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
import { performMinimum } from "./min-utils";
import { parseCriteria, matchesParsedCriteria } from "../../criteria-parser";

/**
 * MINIFS function - Returns the minimum value among cells specified by multiple criteria
 *
 * Usage: MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 *
 * min_range: The range of cells from which you want the minimum
 * criteria_range1: The first range in which to evaluate criteria
 * criteria1: The criteria to apply to criteria_range1
 * criteria_range2, criteria2: Optional additional criteria pairs
 *
 * Examples:
 *   MINIFS(B1:B10, A1:A10, "Apple") - min of B1:B10 where A1:A10 = "Apple"
 *   MINIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - min of C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 *
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are considered for the minimum
 * - Returns error if no values match all criteria
 */
export const MINIFS: FunctionDefinition = {
  name: "MINIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments using utility
    const argError = validateMultiCriteriaArgs("MINIFS", node.args.length);
    if (argError) {
      return argError;
    }

    // Evaluate min range
    const minRangeResult = this.evaluateNode(node.args[0]!, context);
    if (minRangeResult.type === "error") {
      return minRangeResult;
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

    // Use shared minimum utility for standard cases
    const matchingValues = processMultiCriteriaValues(
      this,
      minRangeResult,
      criteriaPairs,
      context,
      "col-major"
    );

    return performMinimum(matchingValues);
  },
};
