import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";
import { lessThanOrEqual } from "./less-than-or-equal";

/**
 * Greater than operator (>) - Returns TRUE if left > right
 * Implemented as NOT (left <= right)
 */
export const greaterThan: ArethmeticEvaluator = (left, right, errAddress) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot compare ${left.type} and ${right.type}`,
      errAddress: errAddress,
    };
  }

  const lteResult = lessThanOrEqual(left, right, errAddress);
  if (lteResult.type === "error" || lteResult.type === "awaiting-evaluation") {
    return lteResult;
  }
  
  if (lteResult.type === "boolean") {
    return { type: "boolean", value: !lteResult.value };
  }
  
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Invalid comparison result",
    errAddress: errAddress,
  };
};
