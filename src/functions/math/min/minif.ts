import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import {
  processMultiCriteriaValues,
  validateSingleCriteriaArgs,
} from "../../criteria-utils";
import { parseCriteria, matchesParsedCriteria } from "../../criteria-parser";
import { performMinimum } from "./min-utils";

/**
 * MINIF function - Returns the minimum value among cells specified by a criteria
 *
 * Usage: MINIF(range, criteria, [min_range])
 *
 * range: The range of cells to evaluate against the criteria
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * min_range: Optional. The range to find minimum from. If omitted, uses the range parameter
 *
 * Examples:
 *   MINIF(A1:A10, "Apple") - min of cells in A1:A10 that contain "Apple"
 *   MINIF(B1:B10, ">10") - min of cells in B1:B10 with values greater than 10
 *   MINIF(A1:A10, "Apple", B1:B10) - min of B1:B10 where A1:A10 contains "Apple"
 *
 * Note:
 * - Only numeric values are considered for the minimum
 * - Returns error if no values match the criteria
 */
export const MINIF: FunctionDefinition = {
  name: "MINIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments
    const argError = validateSingleCriteriaArgs("MINIF", node.args.length, context);
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
        message: "MINIF criteria must be a single value",
        errAddress: context.dependencyNode,
      };
    }

    // Parse criteria
    const parsedCriteria = parseCriteria(criteriaResult.result);
    if (parsedCriteria.type === "error") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: parsedCriteria.message,
        errAddress: context.dependencyNode,
      };
    }

    // Determine min range (third argument if present, otherwise same as criteria range)
    const minRangeResult =
      node.args.length === 3
        ? this.evaluateNode(node.args[2]!, context)
        : criteriaRangeResult;

    if (minRangeResult.type === "error") {
      return minRangeResult;
    }

    // Use shared minimum utility for standard cases
    const matchingValues = processMultiCriteriaValues(
      this,
      minRangeResult,
      [{ rangeResult: criteriaRangeResult, parsedCriteria }],
      context,
      "col-major"
    );

    if (matchingValues.type === "error" || matchingValues.type === "awaiting-evaluation") {
      return matchingValues;
    }

    return performMinimum(matchingValues.values, context);
  },
};
