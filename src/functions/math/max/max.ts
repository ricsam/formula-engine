import {
  FormulaError,
  type CellAddress,
  type CellInfinity,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import { OpenRangeEvaluator } from "../open-range-evaluator";

/**
 * MAX function - Returns the largest number in the arguments
 */
export const MAX: FunctionDefinition = {
  name: "MAX",
  evaluate: function (node, context) {
    const parseResult = (
      result: FunctionEvaluationResult
    ):
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      if (result.type === "error") {
        return result;
      }
      if (result.type === "value") {
        if (
          result.result.type === "number" ||
          result.result.type === "infinity"
        ) {
          return result.result;
        }
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `Can't find max of non-number (${result.result.type}, ${result.result.value})`,
        };
      }
      const range = result.spillArea(context.currentCell);
      let maxValue = -Infinity;
      let hasValues = false;

      const cellValues = result.evaluateAllCells.call(this, {
        context,
        evaluate: result.evaluate,
        origin: context.currentCell,
      });

      for (const cellValue of cellValues) {
        const parsed = parseResult(cellValue);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          if (parsed.sign === "positive") {
            return parsed; // Positive infinity is always the maximum
          }
          // Negative infinity - continue checking other values
        } else {
          maxValue = Math.max(maxValue, parsed.value);
          hasValues = true;
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Cannot find maximum of empty range",
        };
      }

      return {
        type: "number",
        value: maxValue,
      };
    };

    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      let maxValue = -Infinity;
      let hasValues = false;
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);
        const parsed = parseResult(value);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          if (parsed.sign === "positive") {
            return parsed; // Positive infinity is always the maximum
          }
        } else {
          maxValue = Math.max(maxValue, parsed.value);
          hasValues = true;
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "MAX requires at least one numeric argument",
        };
      }

      return {
        type: "number",
        value: maxValue,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
