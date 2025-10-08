import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import {
  parseCriteriaPairs,
  processMultiCriteriaValues,
  validateMultiCriteriaArgs,
} from "../../criteria-utils";
import { performAverage } from "./average-utils";

/**
 * AVERAGEIFS function - Calculates the average of cells in a range that meet multiple criteria
 *
 * Usage: AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 *
 * average_range: The range of cells to average
 * criteria_range1: The first range to evaluate against criteria1
 * criteria1: The first criteria to match against
 * criteria_range2, criteria2: Optional additional criteria pairs
 *
 * Examples:
 *   AVERAGEIFS(B1:B10, A1:A10, "Apple") - averages B1:B10 where A1:A10 = "Apple"
 *   AVERAGEIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10") - averages C1:C10 where A1:A10 = "Apple" AND B1:B10 > 10
 *
 * Note:
 * - All criteria must be satisfied for a value to be included
 * - Only numeric values are included in the average calculation
 * - Returns error if no values match all criteria
 */
export const AVERAGEIFS: FunctionDefinition = {
  name: "AVERAGEIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments using utility
    const argError = validateMultiCriteriaArgs("AVERAGEIFS", node.args.length, context);
    if (argError) {
      return argError;
    }

    // Evaluate average range
    const averageRangeResult = this.evaluateNode(node.args[0]!, context);
    if (averageRangeResult.type === "error") {
      return averageRangeResult;
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

    // Use shared average utility for standard cases
    const matchingValues = processMultiCriteriaValues(
      this,
      averageRangeResult,
      criteriaPairs,
      context,
      "col-major"
    );

    return performAverage(matchingValues, context);
  },
};
