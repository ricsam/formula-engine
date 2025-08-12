import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

export const add: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot add ${left.type} and ${right.type}`,
    };
  }

  // Handle infinity cases
  if (left.type === "infinity" || right.type === "infinity") {
    
    // Both infinity
    if (left.type === "infinity" && right.type === "infinity") {
      if (left.sign === right.sign) {
        // +∞ + +∞ = +∞, -∞ + -∞ = -∞
        return { type: "infinity", sign: left.sign };
      } else {
        // +∞ + -∞ = NaN (undefined behavior)
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "Cannot add positive and negative infinity",
        };
      }
    }
    
    // One infinity, one number: infinity dominates
    if (left.type === "infinity" && right.type === "number") {
      return { type: "infinity", sign: left.sign };
    } else if (right.type === "infinity" && left.type === "number") {
      return { type: "infinity", sign: right.sign };
    }
  }
  
  // Both numbers
  if (left.type === "number" && right.type === "number") {
    const result = left.value + right.value;
    
    // Check for overflow to infinity
    if (result === Infinity) {
      return { type: "infinity", sign: "positive" };
    }
    if (result === -Infinity) {
      return { type: "infinity", sign: "negative" };
    }
    
    return { type: "number", value: result };
  }
  
  // This should never be reached due to the type check at the beginning
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `Cannot add ${left.type} and ${right.type}`,
  };
};
