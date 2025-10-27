import type { ArethmeticEvaluator } from "../../core/types";
import { FormulaError } from "../../core/types";
import { lessThan } from "./less-than";

/**
 * Greater than or equal operator (>=) - Returns TRUE if left >= right
 * Implemented as NOT (left < right)
 */
export const greaterThanOrEqual: ArethmeticEvaluator = (left, right, context) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot compare ${left.type} and ${right.type}`,
      errAddress: context.dependencyNode,
    };
  }

  const ltResult = lessThan(left, right, context);
  if (ltResult.type === "error" || ltResult.type === "awaiting-evaluation") {
    return ltResult;
  }
  
  if (ltResult.type === "boolean") {
    return { type: "boolean", value: !ltResult.value };
  }
  
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Invalid comparison result",
    errAddress: context.dependencyNode,
  };
};
