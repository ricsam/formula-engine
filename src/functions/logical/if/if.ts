import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
  type ErrorEvaluationResult,
} from "src/core/types";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * IF function - Returns one value if a condition is true and another value if it's false
 * 
 * Usage: IF(logical_test, value_if_true, [value_if_false])
 * 
 * logical_test: The condition to evaluate (required)
 * value_if_true: The value to return if the condition is true (required)
 * value_if_false: The value to return if the condition is false (optional, defaults to FALSE)
 * 
 * Examples:
 *   IF(A1>10, "High", "Low") - returns "High" if A1>10, otherwise "Low"
 *   IF(B1="", "Empty", B1) - returns "Empty" if B1 is empty, otherwise B1's value
 *   IF(C1<0, "Negative") - returns "Negative" if C1<0, otherwise FALSE
 * 
 * Note:
 * - Supports spilled values (dynamic arrays) for all arguments
 * - Logical test evaluation: 0 and empty string are FALSE, everything else is TRUE
 * - If logical_test is spilled, returns spilled results
 */

/**
 * Convert a cell value to boolean for logical evaluation
 */
function isTruthy(value: CellValue): boolean {
  switch (value.type) {
    case "boolean":
      return value.value;
    case "number":
      return value.value !== 0;
    case "string":
      return value.value !== "";
    case "infinity":
      return true; // Infinity is truthy
    default:
      return false;
  }
}

/**
 * Helper for creating spilled-values result for IF function
 */
function createIfSpilledResult(
  this: FormulaEvaluator,
  {
    logicalTestResult,
    valueIfTrueResult,
    valueIfFalseResult,
    context,
  }: {
    logicalTestResult: FunctionEvaluationResult;
    valueIfTrueResult: FunctionEvaluationResult;
    valueIfFalseResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): SpilledValuesEvaluationResult | ErrorEvaluationResult {
  const hasSpilledTest = logicalTestResult.type === "spilled-values";
  const hasSpilledTrue = valueIfTrueResult.type === "spilled-values";
  const hasSpilledFalse = valueIfFalseResult.type === "spilled-values";

  if (!hasSpilledTest && !hasSpilledTrue && !hasSpilledFalse) {
    throw new Error("createIfSpilledResult called without spilled values");
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress): SpreadsheetRange => {
      // Calculate spill area (union of all spilled ranges)
      let spillArea: SpreadsheetRange | undefined;

      const spilledResults = [
        hasSpilledTest ? logicalTestResult : null,
        hasSpilledTrue ? valueIfTrueResult : null,
        hasSpilledFalse ? valueIfFalseResult : null,
      ].filter(Boolean) as SpilledValuesEvaluationResult[];

      for (const result of spilledResults) {
        const currentSpillArea = result.spillArea(origin);
        if (!spillArea) {
          spillArea = currentSpillArea;
        } else {
          spillArea = this.unionRanges(
            this.projectRange(spillArea, origin),
            this.projectRange(currentSpillArea, origin)
          );
        }
      }

      if (!spillArea) {
        throw new Error("No spilled values found");
      }
      return spillArea;
    },
    source: "IF with spilled values",
    evaluate: (spilledCell: any, evalContext: any): SingleEvaluationResult => {
      // Evaluate all arguments at this spilled position
      const spillLogicalResult = hasSpilledTest
        ? logicalTestResult.evaluate(spilledCell, evalContext)
        : logicalTestResult;
      const spillTrueResult = hasSpilledTrue
        ? valueIfTrueResult.evaluate(spilledCell, evalContext)
        : valueIfTrueResult;
      const spillFalseResult = hasSpilledFalse
        ? valueIfFalseResult.evaluate(spilledCell, evalContext)
        : valueIfFalseResult;

      // Check for errors
      if (spillLogicalResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled logical test has not been evaluated",
        };
      }
      if (spillLogicalResult.type === "error") {
        return spillLogicalResult;
      }

      if (spillTrueResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled true value has not been evaluated",
        };
      }
      if (spillTrueResult.type === "error") {
        return spillTrueResult;
      }

      if (spillFalseResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled false value has not been evaluated",
        };
      }
      if (spillFalseResult.type === "error") {
        return spillFalseResult;
      }

      // Perform IF logic
      if (spillLogicalResult.type !== "value") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid logical test result",
        };
      }

      const isTrue = isTruthy(spillLogicalResult.result);
      return isTrue ? spillTrueResult : spillFalseResult;
    },
    evaluateAllCells: (intersectingRange: any) => {
      throw new Error("WIP: evaluateAllCells for IF is not implemented");
    },
  };
}

/**
 * IF function implementation
 */
export const IF: FunctionDefinition = {
  name: "IF",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "IF function takes 2 or 3 arguments",
      };
    }

    // Evaluate logical test
    const logicalTestResult = this.evaluateNode(node.args[0]!, context);
    if (logicalTestResult.type === "error") {
      return logicalTestResult;
    }

    // Evaluate value_if_true
    const valueIfTrueResult = this.evaluateNode(node.args[1]!, context);
    if (valueIfTrueResult.type === "error") {
      return valueIfTrueResult;
    }

    // Evaluate value_if_false (optional, defaults to FALSE)
    let valueIfFalseResult: FunctionEvaluationResult;
    if (node.args.length > 2) {
      valueIfFalseResult = this.evaluateNode(node.args[2]!, context);
      if (valueIfFalseResult.type === "error") {
        return valueIfFalseResult;
      }
    } else {
      valueIfFalseResult = {
        type: "value",
        result: { type: "boolean", value: false },
      };
    }

    // Handle spilled values
    const hasSpilledValues = 
      logicalTestResult.type === "spilled-values" ||
      valueIfTrueResult.type === "spilled-values" ||
      valueIfFalseResult.type === "spilled-values";

    if (hasSpilledValues) {
      return createIfSpilledResult.call(this, {
        logicalTestResult,
        valueIfTrueResult,
        valueIfFalseResult,
        context,
      });
    }

    // All arguments are single values
    if (logicalTestResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid logical test argument",
      };
    }

    // Evaluate the logical test
    const isTrue = isTruthy(logicalTestResult.result);

    // Return the appropriate value
    return isTrue ? valueIfTrueResult : valueIfFalseResult;
  },
};
