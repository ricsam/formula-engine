import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { createArgumentIterator } from "../../function-utils";

/**
 * STDEV.S(number1, [number2], ...)
 *
 * Returns the standard deviation of a sample: the square root of
 * `sum((x - mean)^2) / (n - 1)` over the numbers in the arguments.
 *
 * Examples:
 *   STDEV.S(2, 4, 4, 4, 5, 5, 7, 9) - 2.138...
 *   STDEV.S(B2:B97) - the numbers of B2:B97
 *   STDEV.S(FILTER(C2:C97, A2:A97="control")) - the numbers of the control rows
 *
 * - Numbers in the arguments count, wherever they come from: directly supplied,
 *   in a range, or in a spilled result such as FILTER or SEQUENCE. Text, logical
 *   values and empty cells are skipped, like AVERAGE skips them. Excel counts a
 *   logical value or numeric text that is supplied directly (`STDEV.S(1, TRUE)`
 *   uses 1 and 1 there); this engine skips them.
 * - Errors propagate, also from inside a range, as in Excel.
 * - Fewer than two numbers give #DIV/0!, as in Excel.
 * - INFINITY or -INFINITY among the numbers gives #NUM!: their deviation from
 *   the mean is undefined.
 * - The mean is computed first and the squared deviations from it are summed
 *   in a second pass, which keeps the result accurate for numbers with a large
 *   common offset.
 * - `STDEV` is an alias: Excel's older name for the same calculation.
 */
export const STDEV_S: FunctionDefinition = {
  name: "STDEV.S",
  aliases: ["STDEV"],
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length === 0) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "STDEV.S function requires at least 1 argument",
        errAddress: context.dependencyNode,
      };
    }

    const argumentValues = createArgumentIterator(
      this,
      node,
      context,
      "col-major"
    );
    if (argumentValues.type !== "values") {
      return argumentValues;
    }

    const numbers: number[] = [];
    for (const result of argumentValues.values) {
      if (result.type === "error" || result.type === "awaiting-evaluation") {
        return result;
      }
      const value = result.result;
      if (value.type === "infinity") {
        return {
          type: "error",
          err: FormulaError.NUM,
          message: "STDEV.S: the standard deviation of an infinity is undefined",
          errAddress: context.dependencyNode,
        };
      }
      if (value.type === "number") {
        numbers.push(value.value);
      }
    }

    if (numbers.length < 2) {
      return {
        type: "error",
        err: FormulaError.DIV0,
        message: `STDEV.S needs at least 2 numbers, got ${numbers.length}`,
        errAddress: context.dependencyNode,
      };
    }

    let sum = 0;
    for (const value of numbers) {
      sum += value;
    }
    const mean = sum / numbers.length;
    let squaredDeviations = 0;
    for (const value of numbers) {
      const deviation = value - mean;
      squaredDeviations += deviation * deviation;
    }
    const result = Math.sqrt(squaredDeviations / (numbers.length - 1));
    if (!Number.isFinite(result)) {
      return {
        type: "error",
        err: FormulaError.NUM,
        message: "STDEV.S: the result is too large for a number",
        errAddress: context.dependencyNode,
      };
    }

    return { type: "value", result: { type: "number", value: result } };
  },
};
