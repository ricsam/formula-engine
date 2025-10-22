import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import { performMaximum } from "./max-utils";
import { createArgumentIterator } from "src/functions/function-utils";

/**
 * MAX function - Returns the largest number in the arguments
 *
 * Usage: MAX(value1, [value2], ...)
 *
 * Examples:
 *   MAX(1, 2, 3) - returns 3
 *   MAX(A1:A10) - finds maximum in the range A1:A10
 *   MAX(A1:A5, B1:B5) - finds maximum across both ranges
 *
 * Note:
 * - Only numeric values are considered
 * - Text, logical values, and empty cells are ignored in ranges
 * - Direct text/logical arguments cause errors
 * - Returns error if no numeric values are found
 */
export const MAX: FunctionDefinition = {
  name: "MAX",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Create iterator over all argument values
    const argumentValues = createArgumentIterator(this, node, context, "col-major");
    if (argumentValues.type !== "values") {
      return argumentValues;
    }
    // Use shared maximum utility (now with unified error propagation)
    return performMaximum(argumentValues.values, context);
  },
};
