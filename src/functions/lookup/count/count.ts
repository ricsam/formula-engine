import {
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import { createArgumentIterator } from "src/functions/function-utils";
import { performCount } from "./count-utils";

/**
 * COUNT function - Counts the number of cells that contain numbers
 * 
 * Usage: COUNT(value1, [value2], ...)
 * 
 * Examples:
 *   COUNT(1, 2, 3) - returns 3
 *   COUNT(A1:A10) - counts numeric values in the range A1:A10
 *   COUNT(A1:A5, B1:B5) - counts numeric values in both ranges
 * 
 * Note:
 * - Only numeric values are counted (numbers and infinities)
 * - Text, logical values, and empty cells are ignored
 * - Errors propagate immediately
 * 
 * From Excel documentation:
 * "The COUNT function counts the number of cells that contain numbers, 
 * and counts numbers within the list of arguments."
 */
export const COUNT: FunctionDefinition = {
  name: "COUNT",
  evaluate: function (node, context): FunctionEvaluationResult {
    // Create iterator over all argument values
    const argumentValues = createArgumentIterator(this, node, context, "col-major");

    if (argumentValues.type !== "values") {
      return argumentValues;
    }
    // Use shared counting utility
    return performCount(argumentValues.values);
  },
};
