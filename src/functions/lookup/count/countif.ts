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

/**
 * COUNTIF function - Counts cells in a range that meet a criteria
 * 
 * Usage: COUNTIF(range, criteria)
 * 
 * range: The range of cells to evaluate
 * criteria: The criteria to match against. Can be:
 *   - Exact value: "Apple", 42
 *   - Comparison: ">10", "<=5", "<>0"
 *   - Wildcards: "App*", "?ruit"
 * 
 * Examples:
 *   COUNTIF(A1:A10, "Apple") - counts cells containing "Apple"
 *   COUNTIF(B1:B10, ">10") - counts cells with values greater than 10
 *   COUNTIF(C1:C10, "App*") - counts cells starting with "App"
 * 
 * Note:
 * - Supports type coercion for comparisons
 * - Case-sensitive string matching
 * - Wildcards: * matches any sequence, ? matches any single character
 */



/**
 * COUNTIF function implementation
 */
export const COUNTIF: FunctionDefinition = {
  name: "COUNTIF",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Validate arguments - COUNTIF takes exactly 2 arguments
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "COUNTIF function takes exactly 2 arguments",
      };
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
        message: "COUNTIF criteria must be a single value",
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

    // Special case: counting empty cells over infinite ranges
    if (criteriaRangeResult.type === "spilled-values" && 
        parsedCriteria.type === "exact" && 
        parsedCriteria.value.type === "string" && 
        parsedCriteria.value.value === "") {
      
      const spillArea = criteriaRangeResult.spillArea(context.currentCell);
      
      // Check if this is an infinite range
      if (spillArea.end.col.type === "infinity" || spillArea.end.row.type === "infinity") {
        // Return infinity for infinite empty cell count
        return {
          type: "value",
          result: { type: "infinity", sign: "positive" },
        };
      }
    }

    // Use shared criteria processing - count all matching values (including non-numeric)
    let count = 0;

    for (const result of processMultiCriteriaValues(
      this,
      criteriaRangeResult,
      [{ rangeResult: criteriaRangeResult, parsedCriteria }],
      context
    )) {
      // COUNTIF counts all matching cells, including errors and non-numeric values
      count++;
    }

    return {
      type: "value",
      result: { type: "number", value: count },
    };
  },
};
