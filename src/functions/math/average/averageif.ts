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
import { performAverage } from "./average-utils";

/**
 * AVERAGEIF function - Calculates the average of cells in a range that meet a criteria
 * 
 * Usage: AVERAGEIF(range, criteria, [average_range])
 * 
 * range: The range of cells to evaluate against the criteria
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * average_range: Optional. The range to average. If omitted, uses the range parameter
 * 
 * Examples:
 *   AVERAGEIF(A1:A10, "Apple") - averages cells in A1:A10 that contain "Apple"
 *   AVERAGEIF(B1:B10, ">10") - averages cells in B1:B10 with values greater than 10
 *   AVERAGEIF(A1:A10, "Apple", B1:B10) - averages B1:B10 where A1:A10 contains "Apple"
 * 
 * Note:
 * - Only numeric values are included in the average calculation
 * - Returns error if no values match the criteria
 */
export const AVERAGEIF: FunctionDefinition = {
  name: "AVERAGEIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments
    const argError = validateSingleCriteriaArgs("AVERAGEIF", node.args.length, context);
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
        message: "AVERAGEIF criteria must be a single value",
        errAddress: context.originCell.cellAddress,
      };
    }

    // Parse criteria
    const parsedCriteria = parseCriteria(criteriaResult.result);
    if (parsedCriteria.type === "error") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: parsedCriteria.message,
        errAddress: context.originCell.cellAddress,
      };
    }

    // Determine average range (third argument if present, otherwise same as criteria range)
    const averageRangeResult = node.args.length === 3 
      ? this.evaluateNode(node.args[2]!, context)
      : criteriaRangeResult;
    
    if (averageRangeResult.type === "error") {
      return averageRangeResult;
    }

    // Use shared average utility for standard cases
    const matchingValues = processMultiCriteriaValues(
      this,
      averageRangeResult,
      [{ rangeResult: criteriaRangeResult, parsedCriteria }],
      context,
      "col-major"
    );

    return performAverage(matchingValues, context);
  },
};
