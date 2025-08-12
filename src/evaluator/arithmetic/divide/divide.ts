import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

export const divide: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot divide ${left.type} and ${right.type}`,
    };
  }

  // Handle infinity cases
  if (left.type === "infinity" || right.type === "infinity") {
    
    // Both infinity: ∞ / ∞ = indeterminate
    if (left.type === "infinity" && right.type === "infinity") {
      return {
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot divide infinity by infinity",
      };
    }
    
    // Infinity divided by number: sign rules apply
    if (left.type === "infinity" && right.type === "number") {
      if (right.value === 0) {
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "Cannot divide infinity by zero",
        };
      }
      
      if (right.value > 0) {
        return { type: "infinity", sign: left.sign };
      } else {
        const resultSign = left.sign === "positive" ? "negative" : "positive";
        return { type: "infinity", sign: resultSign };
      }
    }
    
    // Number divided by infinity: approaches zero
    if (left.type === "number" && right.type === "infinity") {
      return { type: "number", value: 0 };
    }
  }
  
  // Both numbers
  if (left.type === "number" && right.type === "number") {
    // Division by zero
    if (right.value === 0) {
      if (left.value === 0) {
        // 0 / 0 = indeterminate
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "0 / 0 is undefined",
        };
      }
      
      // n / 0 = ±∞ depending on sign of n
      if (left.value > 0) {
        return { type: "infinity", sign: "positive" };
      } else {
        return { type: "infinity", sign: "negative" };
      }
    }
    
    const result = left.value / right.value;
    
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
    message: `Cannot divide ${left.type} and ${right.type}`,
  };
};