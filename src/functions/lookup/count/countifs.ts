import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import { 
  parseCriteriaPairs, 
  processMultiCriteriaValues, 
  validateCountifsArgs 
} from "../../criteria-utils";

/**
 * COUNTIFS function - Counts cells that meet multiple criteria
 * 
 * Usage: COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)
 * 
 * criteria_range1: The first range to evaluate against criteria1
 * criteria1: The first criteria to match against
 * criteria_range2, criteria2: Optional additional criteria pairs
 * 
 * Examples:
 *   COUNTIFS(A1:A10, "Apple") - counts cells in A1:A10 that equal "Apple"
 *   COUNTIFS(A1:A10, "Apple", B1:B10, ">10") - counts where A1:A10 = "Apple" AND B1:B10 > 10
 * 
 * Note:
 * - All criteria must be satisfied for a cell to be counted
 * - Counts all matching cells, not just numeric ones
 * - Returns 0 if no cells match all criteria
 * - Unlike other IFS functions, COUNTIFS counts the first criteria range itself
 */
export const COUNTIFS: FunctionDefinition = {
  name: "COUNTIFS",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments
    const argError = validateCountifsArgs(node.args.length);
    if (argError) {
      return argError;
    }

    // Parse criteria pairs starting from argument 0 (COUNTIFS has different syntax)
    const criteriaPairsResult = parseCriteriaPairs.call(this, node, context, this.evaluateNode, 0);
    if ('type' in criteriaPairsResult) {
      return criteriaPairsResult;
    }

    // The first criteria range is what we count
    const countRangeResult = criteriaPairsResult[0]!.rangeResult;

    // Process values with criteria using generator - count all matching cells (including non-numeric)
    let count = 0;

    for (const result of processMultiCriteriaValues(
      this,
      countRangeResult,
      criteriaPairsResult,
      context,
      "col-major"
    )) {
      // COUNTIFS counts all matching cells, including errors and non-numeric values
      count++;
    }

    return {
      type: "value",
      result: { type: "number", value: count },
    };
  },
};
