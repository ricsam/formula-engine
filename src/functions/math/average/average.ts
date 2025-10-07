import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import { performAverage } from "./average-utils";
import { createArgumentIterator } from "src/functions/function-utils";

/**
 * AVERAGE function - Calculates the arithmetic mean of all numbers in the arguments
 *
 * Usage: AVERAGE(value1, [value2], ...)
 *
 * Examples:
 *   AVERAGE(1, 2, 3) - returns 2
 *   AVERAGE(A1:A10) - averages all numbers in the range A1:A10
 *   AVERAGE(A1:A5, B1:B5) - averages all numbers in both ranges
 *
 * Note:
 * - Only numeric values are included in the calculation
 * - Text, logical values, and empty cells are ignored in ranges
 * - Direct text/logical arguments cause errors
 * - Returns #DIV/0! if no numeric values are found
 */
export const AVERAGE: FunctionDefinition = {
  name: "AVERAGE",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Create iterator over all argument values
    const argumentValues = createArgumentIterator(
      this,
      node,
      context,
      "col-major"
    );

    // Use shared averaging utility (now with unified error propagation)
    return performAverage(argumentValues);
  },
};
