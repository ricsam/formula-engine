import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import {
  processMultiCriteriaValues,
  validateSingleCriteriaArgs,
} from "../../criteria-utils";
import { parseCriteria, matchesParsedCriteria } from "../../criteria-parser";
import { performSummation } from "./summation-utils";

/**
 * SUMIF function - Sums cells in a range that meet a criteria
 *
 * Usage: SUMIF(range, criteria, [sum_range])
 *
 * range: The range of cells to evaluate against the criteria
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * sum_range: Optional. The range to sum. If omitted, uses the range parameter
 *
 * Examples:
 *   SUMIF(A1:A10, "Apple") - sums cells in A1:A10 that contain "Apple"
 *   SUMIF(B1:B10, ">10") - sums cells in B1:B10 with values greater than 10
 *   SUMIF(A1:A10, "Apple", B1:B10) - sums B1:B10 where A1:A10 contains "Apple"
 *
 * Note:
 * - Only numeric values are included in the sum
 * - Returns 0 if no values match the criteria
 */
export const SUMIF: FunctionDefinition = {
  name: "SUMIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments
    const argError = validateSingleCriteriaArgs("SUMIF", node.args.length, context);
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
        message: "SUMIF criteria must be a single value",
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

    // Determine sum range (third argument if present, otherwise same as criteria range)
    const sumRangeResult =
      node.args.length === 3
        ? this.evaluateNode(node.args[2]!, context)
        : criteriaRangeResult;

    if (sumRangeResult.type === "error") {
      return sumRangeResult;
    }

    // Use shared summation utility
    const matchingValues = processMultiCriteriaValues(
      this,
      sumRangeResult,
      [{ rangeResult: criteriaRangeResult, parsedCriteria }],
      context,
      "col-major"
    );

    if (matchingValues.type === "error" || matchingValues.type === "awaiting-evaluation") {
      return matchingValues;
    }

    return performSummation(matchingValues.values);
  },
};
