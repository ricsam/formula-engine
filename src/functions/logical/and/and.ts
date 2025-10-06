import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
  type ErrorEvaluationResult,
} from "src/core/types";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { EvaluationContext } from "src/evaluator/evaluation-context";

/**
 * AND function - Returns TRUE if all arguments are TRUE, FALSE otherwise
 *
 * Usage: AND(logical1, [logical2], ...)
 *
 * logical1, logical2, ...: Logical values or expressions to evaluate
 *
 * Examples:
 *   AND(TRUE, TRUE) - returns TRUE
 *   AND(TRUE, FALSE) - returns FALSE
 *   AND(A1>10, B1<20) - returns TRUE if both conditions are met
 *   AND(A1:A3) - returns TRUE if all values in A1:A3 are truthy
 *
 * Note:
 * - Supports spilled values (dynamic arrays) for arguments
 * - Logical evaluation: 0 and empty string are FALSE, everything else is TRUE
 * - If any argument is spilled, returns spilled results
 * - Short-circuit evaluation: stops at first FALSE value
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

// AND function does not create spilled results - it processes all values in ranges
// as individual logical tests and returns a single boolean result

/**
 * AND function implementation
 */
export const AND: FunctionDefinition = {
  name: "AND",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length === 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "AND function requires at least one argument",
      };
    }

    // Process arguments one by one (like SUM function)
    for (const arg of node.args) {
      const argResult = this.evaluateNode(arg, context);

      if (argResult.type === "error") {
        return argResult;
      }

      if (argResult.type === "value") {
        // Single value - check if truthy
        if (!isTruthy(argResult.result)) {
          return {
            type: "value",
            result: { type: "boolean", value: false },
          };
        }
      } else if (argResult.type === "spilled-values") {
        // Range - check all values in the range
        const cellValues = argResult.evaluateAllCells.call(this, {
          context,
          origin: context.originCell.cellAddress,
          evaluate: argResult.evaluate,
        });

        for (const cellValue of cellValues) {
          if (cellValue.result.type === "error") {
            return cellValue.result;
          }
          if (cellValue.result.type === "value") {
            if (!isTruthy(cellValue.result.result)) {
              return {
                type: "value",
                result: { type: "boolean", value: false },
              };
            }
          }
        }
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid argument type for AND function",
        };
      }
    }

    // All arguments are truthy
    return {
      type: "value",
      result: { type: "boolean", value: true },
    };
  },
};
