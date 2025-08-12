import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

export const multiply: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot multiply ${left.type} and ${right.type}`,
    };
  }

  // Handle infinity cases
  if (left.type === "infinity" || right.type === "infinity") {
    
    // Handle infinity * 0 = NaN (indeterminate)
    if ((left.type === "infinity" && right.type === "number" && right.value === 0) ||
        (left.type === "number" && left.value === 0 && right.type === "infinity")) {
      return {
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot multiply infinity by zero",
      };
    }
    
    // Both infinity: sign rules apply
    if (left.type === "infinity" && right.type === "infinity") {
      const resultSign = left.sign === right.sign ? "positive" : "negative";
      return { type: "infinity", sign: resultSign };
    }
    
    // Infinity * number: determine sign based on number's sign
    if (left.type === "infinity" && right.type === "number") {
      if (right.value > 0) {
        return { type: "infinity", sign: left.sign };
      } else if (right.value < 0) {
        const resultSign = left.sign === "positive" ? "negative" : "positive";
        return { type: "infinity", sign: resultSign };
      }
      // right.value === 0 is handled above
    }
    
    // Number * infinity: determine sign based on number's sign
    if (left.type === "number" && right.type === "infinity") {
      if (left.value > 0) {
        return { type: "infinity", sign: right.sign };
      } else if (left.value < 0) {
        const resultSign = right.sign === "positive" ? "negative" : "positive";
        return { type: "infinity", sign: resultSign };
      }
      // left.value === 0 is handled above
    }
  }
  
  // Both numbers
  if (left.type === "number" && right.type === "number") {
    const result = left.value * right.value;
    
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
    message: `Cannot multiply ${left.type} and ${right.type}`,
  };
};