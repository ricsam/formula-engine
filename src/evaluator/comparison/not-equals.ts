import type { ArethmeticEvaluator } from "src/core/types";
import { FormulaError } from "src/core/types";
import { equals } from "./equals";

/**
 * Not equals operator (<>) - Returns TRUE if values are not equal
 * Uses the equals operator and negates the result
 */
export const notEquals: ArethmeticEvaluator = (left, right) => {
  const equalsResult = equals(left, right);
  
  if (equalsResult.type === "error") {
    return equalsResult;
  }
  
  if (equalsResult.type === "boolean") {
    return { type: "boolean", value: !equalsResult.value };
  }
  
  return {
    type: "error",
    err: FormulaError.VALUE,
    message: "Invalid result from equals comparison",
  };
};
