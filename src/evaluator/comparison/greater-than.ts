import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";
import { lessThanOrEqual } from "./less-than-or-equal";

/**
 * Greater than operator (>) - Returns TRUE if left > right
 * Implemented as NOT (left <= right)
 */
export const greaterThan: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot compare ${left.type} and ${right.type}`,
    };
  }

  const lteResult = lessThanOrEqual(left, right);
  if (lteResult.type === "error") {
    return lteResult;
  }
  
  if (lteResult.type === "boolean") {
    return { type: "boolean", value: !lteResult.value };
  }
  
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Invalid comparison result",
  };
};
