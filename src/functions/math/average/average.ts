import {
  FormulaError,
  type CellAddress,
  type CellInfinity,
  type CellNumber,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SpilledValuesEvaluationResult,
  type ValueEvaluationResult,
} from "src/core/types";
import { OpenRangeEvaluator } from "../open-range-evaluator";

/**
 * AVERAGE function - Calculates the arithmetic mean of all numbers in the arguments
 */
export const AVERAGE: FunctionDefinition = {
  name: "AVERAGE",
  evaluate: function (node, context) {
    const parseResult = (
      result: FunctionEvaluationResult
    ):
      | { type: "number"; value: number; count: number }
      | { type: "infinity"; sign: "positive" | "negative"; count: number }
      | { type: "error"; err: FormulaError; message: string } => {
      if (result.type === "error") {
        return result;
      }
      if (result.type === "value") {
        if (result.result.type === "number") {
          return { type: "number", value: result.result.value, count: 1 };
        }
        if (result.result.type === "infinity") {
          return {
            type: "infinity",
            sign: result.result.sign,
            count: 1,
          };
        }
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `Can't average non-number (${result.result.type})`,
        };
      }
      // Handle spilled values
      const cellValues = result.evaluateAllCells.call(this, {
        context,
        evaluate: result.evaluate,
        origin: context.currentCell,
      });

      let subTotal = 0;
      let count = 0;
      for (const cellValue of cellValues) {
        const parsed = parseResult(cellValue);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          return parsed;
        }
        subTotal += parsed.value;
        count += parsed.count;
      }

      if (count === 0) {
        return {
          type: "error",
          err: FormulaError.DIV0,
          message: "Cannot calculate average of empty range",
        };
      }

      return {
        type: "number",
        value: subTotal,
        count: count,
      };
    };

    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      let total = 0;
      let totalCount = 0;
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);
        const parsed = parseResult(value);
        if (parsed.type === "error") {
          return parsed;
        }
        if (parsed.type === "infinity") {
          return { type: "infinity", sign: parsed.sign };
        }
        total += parsed.value;
        totalCount += parsed.count;
      }

      if (totalCount === 0) {
        return {
          type: "error",
          err: FormulaError.DIV0,
          message: "AVERAGE requires at least one argument",
        };
      }

      return {
        type: "number",
        value: total / totalCount,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
