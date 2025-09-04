import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

/**
 * Less than operator (<) - Returns TRUE if left < right
 * Only works with numbers and infinity
 */
export const lessThan: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot compare ${left.type} and ${right.type}`,
    };
  }

  // Handle infinity cases
  if (left.type === "infinity" && right.type === "infinity") {
    // -∞ < +∞ is true, otherwise false (same infinities are equal)
    if (left.sign === "negative" && right.sign === "positive") {
      return { type: "boolean", value: true };
    }
    return { type: "boolean", value: false };
  }
  
  if (left.type === "infinity") {
    // +∞ is never less than anything, -∞ is less than any number
    return { type: "boolean", value: left.sign === "negative" };
  }
  
  if (right.type === "infinity") {
    // Any number is less than +∞, no number is less than -∞
    return { type: "boolean", value: right.sign === "positive" };
  }
  
  // Both are numbers
  if (left.type === "number" && right.type === "number") {
    return { type: "boolean", value: left.value < right.value };
  }
  
  // This should never be reached
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `Cannot compare ${left.type} and ${right.type}`,
  };
};
