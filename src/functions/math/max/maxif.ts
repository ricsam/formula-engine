import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import { 
  processMultiCriteriaValues, 
  validateSingleCriteriaArgs 
} from "../../criteria-utils";
import { parseCriteria, matchesParsedCriteria } from "../../criteria-parser";
import { performMaximum } from "./max-utils";

/**
 * MAXIF function - Returns the maximum value among cells specified by a criteria
 * 
 * Usage: MAXIF(range, criteria, [max_range])
 * 
 * range: The range of cells to evaluate against the criteria
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * max_range: Optional. The range to find maximum from. If omitted, uses the range parameter
 * 
 * Examples:
 *   MAXIF(A1:A10, "Apple") - max of cells in A1:A10 that contain "Apple"
 *   MAXIF(B1:B10, ">10") - max of cells in B1:B10 with values greater than 10
 *   MAXIF(A1:A10, "Apple", B1:B10) - max of B1:B10 where A1:A10 contains "Apple"
 * 
 * Note:
 * - Only numeric values are considered for the maximum
 * - Returns error if no values match the criteria
 */
export const MAXIF: FunctionDefinition = {
  name: "MAXIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments
    const argError = validateSingleCriteriaArgs("MAXIF", node.args.length);
    if (argError) {
      return argError;
    }

    // Evaluate criteria range (first argument)
    const criteriaRangeResult = this.evaluateNode(node.args[0]!, context);
    if (criteriaRangeResult.type === "error") {
      return criteriaRangeResult;
    }

    // Evaluate criteria (second argument)
    const criteriaResult = this.evaluateNode(node.args[1]!, context);
    if (criteriaResult.type === "error") {
      return criteriaResult;
    }

    if (criteriaResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MAXIF criteria must be a single value",
      };
    }

    // Parse criteria
    const parsedCriteria = parseCriteria(criteriaResult.result);
    if (parsedCriteria.type === "error") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: parsedCriteria.message,
      };
    }

    // Determine max range (third argument if present, otherwise same as criteria range)
    const maxRangeResult = node.args.length === 3 
      ? this.evaluateNode(node.args[2]!, context)
      : criteriaRangeResult;
    
    if (maxRangeResult.type === "error") {
      return maxRangeResult;
    }

    // Use shared maximum utility
    const matchingValues = processMultiCriteriaValues(
      this,
      maxRangeResult,
      [{ rangeResult: criteriaRangeResult, parsedCriteria }],
      context
    );

    return performMaximum(matchingValues);
  },
};
