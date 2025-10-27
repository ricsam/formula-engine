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
} from "../../../core/types";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";

/**
 * OR function - Returns TRUE if any argument is TRUE, FALSE otherwise
 *
 * Usage: OR(logical1, [logical2], ...)
 *
 * logical1, logical2, ...: Logical values or expressions to evaluate
 *
 * Examples:
 *   OR(TRUE, FALSE) - returns TRUE
 *   OR(FALSE, FALSE) - returns FALSE
 *   OR(A1>10, B1<20) - returns TRUE if either condition is met
 *   OR(A1:A3) - returns TRUE if any value in A1:A3 is truthy
 *
 * Note:
 * - Supports spilled values (dynamic arrays) for arguments
 * - Logical evaluation: 0 and empty string are FALSE, everything else is TRUE
 * - If any argument is spilled, processes all values in the range
 * - Short-circuit evaluation: stops at first TRUE value
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

// OR function does not create spilled results - it processes all values in ranges
// as individual logical tests and returns a single boolean result

/**
 * OR function implementation
 */
export const OR: FunctionDefinition = {
  name: "OR",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length === 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "OR function requires at least one argument",
        errAddress: context.dependencyNode,
      };
    }

    // Process arguments one by one (like AND function)
    for (const arg of node.args) {
      const argResult = this.evaluateNode(arg, context);

      if (argResult.type === "error" || argResult.type === "awaiting-evaluation") {
        return argResult;
      }

      if (argResult.type === "value") {
        // Single value - check if truthy
        if (isTruthy(argResult.result)) {
          // Short-circuit: return TRUE on first truthy value
          return {
            type: "value",
            result: { type: "boolean", value: true },
          };
        }
      } else if (argResult.type === "spilled-values") {
        // Range - check all values in the range
        const cellValues = argResult.evaluateAllCells.call(this, {
          context,
          origin: context.cellAddress,
          evaluate: argResult.evaluate,
          lookupOrder: "col-major",
        });
        if (cellValues.type === "error" || cellValues.type === "awaiting-evaluation") {
          return cellValues;
        }
        for (const cellValue of cellValues.values) {
          if (cellValue.result.type === "error" || cellValue.result.type === "awaiting-evaluation") {
            return cellValue.result;
          }
          if (cellValue.result.type === "value") {
            if (isTruthy(cellValue.result.result)) {
              // Short-circuit: return TRUE on first truthy value
              return {
                type: "value",
                result: { type: "boolean", value: true },
              };
            }
          }
        }
      } else {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid argument type for OR function",
          errAddress: context.dependencyNode,
        };
      }
    }

    // All arguments are falsy
    return {
      type: "value",
      result: { type: "boolean", value: false },
    };
  },
};
