import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

/**
 * Concatenation operator (&) - Concatenates two values as strings
 * Works with numbers and strings only
 */
export const concatenate: ArethmeticEvaluator = (left, right) => {
  // Convert each value to string based on type
  let leftStr: string;
  let rightStr: string;

  // Handle left operand
  switch (left.type) {
    case "string":
      leftStr = left.value;
      break;
    case "number":
      leftStr = left.value.toString();
      break;
    default:
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `Cannot concatenate ${left.type} - only numbers and strings are supported`,
      };
  }

  // Handle right operand
  switch (right.type) {
    case "string":
      rightStr = right.value;
      break;
    case "number":
      rightStr = right.value.toString();
      break;
    default:
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `Cannot concatenate ${right.type} - only numbers and strings are supported`,
      };
  }

  return { type: "string", value: leftStr + rightStr };
};
