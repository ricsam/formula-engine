import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";
import { lessThan } from "./less-than";
import { equals } from "./equals";

/**
 * Less than or equal operator (<=) - Returns TRUE if left <= right
 * Uses lessThan OR equals logic
 */
export const lessThanOrEqual: ArethmeticEvaluator = (left, right) => {
  // Only allow number and infinity types
  if ((left.type !== "number" && left.type !== "infinity") ||
      (right.type !== "number" && right.type !== "infinity")) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `Cannot compare ${left.type} and ${right.type}`,
    };
  }

  const ltResult = lessThan(left, right);
  if (ltResult.type === "error") {
    return ltResult;
  }
  
  if (ltResult.type === "boolean" && ltResult.value === true) {
    return { type: "boolean", value: true };
  }
  
  // Check equality
  const eqResult = equals(left, right);
  if (eqResult.type === "error") {
    return eqResult;
  }
  
  if (eqResult.type === "boolean") {
    return { type: "boolean", value: eqResult.value };
  }
  
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Invalid comparison result",
  };
};
