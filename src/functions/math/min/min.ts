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
 * MIN function - Returns the smallest number in the arguments
 */
export const MIN: FunctionDefinition = {
  name: "MIN",
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
          message: `Can't find min of non-number (${result.result.type}, ${result.result.value})`,
        };
      }
      let minValue = Infinity;
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
          if (parsed.sign === "negative") {
            return parsed; // Negative infinity is always the minimum
          }
          // Positive infinity - continue checking other values
        } else {
          minValue = Math.min(minValue, parsed.value);
          hasValues = true;
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Cannot find minimum of empty range",
        };
      }

      return {
        type: "number",
        value: minValue,
      };
    };

    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      let minValue = Infinity;
      let hasValues = false;
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);
        const parsed = parseResult(value);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          if (parsed.sign === "negative") {
            return parsed; // Negative infinity is always the minimum
          }
          // Positive infinity - continue checking other values
        } else {
          minValue = Math.min(minValue, parsed.value);
          hasValues = true;
        }
      }

      if (!hasValues) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "MIN requires at least one numeric argument",
        };
      }

      return {
        type: "number",
        value: minValue,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
