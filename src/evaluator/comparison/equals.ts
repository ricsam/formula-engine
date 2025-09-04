import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";

/**
 * Equals operator (=) - Returns TRUE if both values are exactly equal
 * Strict type checking: both type and value must match
 */
export const equals: ArethmeticEvaluator = (left, right) => {
  // Check type equality first
  if (left.type !== right.type) {
    return { type: "boolean", value: false };
  }

  // Handle each type specifically
  switch (left.type) {
    case "number":
      if (right.type === "number") {
        return { type: "boolean", value: left.value === right.value };
      }
      break;
    
    case "string":
      if (right.type === "string") {
        return { type: "boolean", value: left.value === right.value };
      }
      break;
    
    case "boolean":
      if (right.type === "boolean") {
        return { type: "boolean", value: left.value === right.value };
      }
      break;
    
    case "infinity":
      if (right.type === "infinity") {
        return { type: "boolean", value: left.sign === right.sign };
      }
      break;
  }

  // This should never be reached due to type check above
  return { type: "boolean", value: false };
};
