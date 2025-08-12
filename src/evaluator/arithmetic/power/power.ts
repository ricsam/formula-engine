import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

export const power: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot exponentiate ${left.type} and ${right.type}`,
    };
  }

  // Handle infinity cases
  if (left.type === "infinity" || right.type === "infinity") {
    
    // Special cases for infinity exponentiation
    if (left.type === "infinity" && right.type === "infinity") {
      // ∞^∞ = +∞ (mathematical convention)
      return { type: "infinity", sign: "positive" };
    }
    
    // Infinity raised to a power
    if (left.type === "infinity" && right.type === "number") {
      if (right.value === 0) {
        // ∞^0 = 1 (mathematical convention)
        return { type: "number", value: 1 };
      }
      
      if (right.value > 0) {
        // ∞^(positive) = ∞, (-∞)^(positive) depends on whether exponent is even/odd
        if (left.sign === "positive") {
          return { type: "infinity", sign: "positive" };
        } else {
          // For negative infinity, we need to check if exponent is integer and even/odd
          if (Number.isInteger(right.value)) {
            const isEven = right.value % 2 === 0;
            return { type: "infinity", sign: isEven ? "positive" : "negative" };
          } else {
            // Non-integer exponent with negative base is complex
            return {
              type: "error",
              err: FormulaError.NUM,
              message: "Cannot raise negative infinity to non-integer power",
            };
          }
        }
      } else {
        // ∞^(negative) = 0
        return { type: "number", value: 0 };
      }
    }
    
    // Number raised to infinity
    if (left.type === "number" && right.type === "infinity") {
      const absBase = Math.abs(left.value);
      
      if (absBase === 1) {
        // 1^∞ = 1, (-1)^∞ is indeterminate but we'll return 1
        return { type: "number", value: 1 };
      }
      
      if (absBase > 1) {
        if (right.sign === "positive") {
          // |base| > 1: base^(+∞) = +∞
          return { type: "infinity", sign: "positive" };
        } else {
          // |base| > 1: base^(-∞) = 0
          return { type: "number", value: 0 };
        }
      } else {
        // 0 < |base| < 1
        if (right.sign === "positive") {
          // 0 < |base| < 1: base^(+∞) = 0
          return { type: "number", value: 0 };
        } else {
          // 0 < |base| < 1: base^(-∞) = +∞
          return { type: "infinity", sign: "positive" };
        }
      }
    }
  }
  
  // Both numbers
  if (left.type === "number" && right.type === "number") {
    // Special cases
    if (left.value === 0 && right.value === 0) {
      // 0^0 is indeterminate, return 1 by convention
      return { type: "number", value: 1 };
    }
    
    if (left.value === 0 && right.value < 0) {
      // 0^(negative) = infinity
      return { type: "infinity", sign: "positive" };
    }
    
    if (left.value < 0 && !Number.isInteger(right.value)) {
      // Negative base with non-integer exponent produces complex numbers
      return {
        type: "error",
        err: FormulaError.NUM,
        message: "Cannot raise negative number to non-integer power",
      };
    }
    
    const result = Math.pow(left.value, right.value);
    
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
    message: `Cannot exponentiate ${left.type} and ${right.type}`,
  };
};