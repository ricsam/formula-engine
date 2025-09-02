import {
  FormulaError,
  type CellAddress,
  type CellInfinity,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "src/core/types";
import { parseCellReference } from "src/core/utils";
import { dependencyNodeToKey } from "src/core/utils/dependency-node-key";
import { OpenRangeEvaluator } from "../open-range-evaluator";

/**
 * SUM function - Adds all numbers in the arguments
 */
export const SUM: FunctionDefinition = {
  name: "SUM",
  evaluate: function (node, context) {
    let total = 0;
    const parseResult = (
      result: SingleEvaluationResult
    ):
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string }
      | undefined => {
      if (result.type === "error") {
        return result;
      }
      if (result.result.type === "infinity") {
        return result.result;
      }
      if (result.result.type === "number") {
        total += result.result.value;
        return undefined;
      }
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `Can't sum non-number (${result.result.type}, ${result.result.value})`,
      };
    };
    const parseArgs = ():
      | CellNumber
      | CellInfinity
      | { type: "error"; err: FormulaError; message: string } => {
      for (const arg of node.args) {
        const value = this.evaluateNode(arg, context);

        if (value.type === "spilled-values") {
          const values = value.evaluateAllCells.call(this, {
            context,
            origin: context.currentCell,
            evaluate: value.evaluate,
          });
          for (const cellValue of values) {
            const parsed = parseResult(cellValue);
            if (parsed) {
              return parsed;
            }
          }
        } else {
          const parsed = parseResult(value);
          if (parsed) {
            return parsed;
          }
        }
      }
      return {
        type: "number",
        value: total,
      };
    };

    const result = parseArgs();

    if (result.type === "error") {
      return result;
    }

    return { type: "value", result };
  },
};
