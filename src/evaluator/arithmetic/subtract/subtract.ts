import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

export const subtract: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot subtract ${left.type} and ${right.type}`,
    };
  }

  // Handle infinity cases
  if (left.type === "infinity" || right.type === "infinity") {
    
    // Both infinity
    if (left.type === "infinity" && right.type === "infinity") {
      if (left.sign === right.sign) {
        // +∞ - +∞ = undefined, -∞ - -∞ = undefined
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "Cannot subtract infinity from same-signed infinity",
        };
      } else {
        // +∞ - -∞ = +∞, -∞ - +∞ = -∞
        return { type: "infinity", sign: left.sign };
      }
    }
    
    // Left infinity, right number: ∞ - n = ∞
    if (left.type === "infinity" && right.type === "number") {
      return { type: "infinity", sign: left.sign };
    }
    
    // Left number, right infinity: n - ∞ = -∞, n - -∞ = +∞
    if (left.type === "number" && right.type === "infinity") {
      const resultSign = right.sign === "positive" ? "negative" : "positive";
      return { type: "infinity", sign: resultSign };
    }
  }
  
  // Both numbers
  if (left.type === "number" && right.type === "number") {
    const result = left.value - right.value;
    
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
    message: `Cannot subtract ${left.type} and ${right.type}`,
  };
};