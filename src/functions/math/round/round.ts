import {
  FormulaError,
  type CellValue,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../../../core/types";
import {
  createBroadcastResult,
  toBroadcastOperand,
} from "../../../evaluator/broadcast";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import { getSuppliedArgument } from "../../../parser/ast";
import { roundDecimal, type DecimalRoundingMode } from "./round-utils";

type RoundingFunctionName = "ROUND" | "ROUNDUP" | "ROUNDDOWN";

const ZERO: FunctionEvaluationResult = {
  type: "value",
  result: { type: "number", value: 0 },
};

function describeValue(value: CellValue): string {
  return value.type === "string" && value.value === ""
    ? "an empty value"
    : value.type;
}

/**
 * Rounds one number to `digits` decimal places, or explains why it cannot.
 */
function roundValue(
  name: RoundingFunctionName,
  mode: DecimalRoundingMode,
  number: CellValue,
  digits: CellValue,
  context: EvaluationContext
): SingleEvaluationResult {
  const invalid = (message: string): SingleEvaluationResult => ({
    type: "error",
    err: FormulaError.VALUE,
    message: `${name} ${message}`,
    errAddress: context.dependencyNode,
  });

  if (number.type !== "number" && number.type !== "infinity") {
    return invalid(`number must be a number, got ${describeValue(number)}`);
  }
  if (digits.type !== "number" && digits.type !== "infinity") {
    return invalid(`num_digits must be a number, got ${describeValue(digits)}`);
  }
  if (number.type === "infinity") {
    // Like INT and CEILING, rounding leaves an infinity unchanged.
    return { type: "value", result: number };
  }

  const places =
    digits.type === "number"
      ? digits.value
      : digits.sign === "positive"
        ? Infinity
        : -Infinity;
  const rounded = roundDecimal(number.value, places, mode);
  if (rounded === undefined) {
    return {
      type: "error",
      err: FormulaError.NUM,
      message: `${name}: the rounded value is too large`,
      errAddress: context.dependencyNode,
    };
  }
  return { type: "value", result: { type: "number", value: rounded } };
}

/**
 * Creates ROUND, ROUNDUP or ROUNDDOWN, which differ only in how they resolve
 * the digits that are rounded off.
 */
export function createRoundingFunction(
  name: RoundingFunctionName,
  mode: DecimalRoundingMode
): FunctionDefinition {
  return {
    name,
    evaluate: function (node, context): FunctionEvaluationResult {
      if (node.args.length !== 2) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: `${name} function takes exactly 2 arguments`,
          errAddress: context.dependencyNode,
        };
      }

      // Evaluate both arguments before returning either, so that one pass
      // registers all dependencies. An omitted argument is 0, as in Excel.
      const numberNode = getSuppliedArgument(node, 0);
      const digitsNode = getSuppliedArgument(node, 1);
      const numberResult = numberNode
        ? this.evaluateNode(numberNode, context)
        : ZERO;
      const digitsResult = digitsNode
        ? this.evaluateNode(digitsNode, context)
        : ZERO;
      if (
        numberResult.type === "error" ||
        numberResult.type === "awaiting-evaluation"
      ) {
        return numberResult;
      }
      if (
        digitsResult.type === "error" ||
        digitsResult.type === "awaiting-evaluation"
      ) {
        return digitsResult;
      }

      if (numberResult.type === "value" && digitsResult.type === "value") {
        return roundValue(
          name,
          mode,
          numberResult.result,
          digitsResult.result,
          context
        );
      }

      return createBroadcastResult({
        operands: [
          toBroadcastOperand(numberResult, context.cellAddress),
          toBroadcastOperand(digitsResult, context.cellAddress),
        ],
        context,
        source: `${name} function`,
        combine: (operandCell) => {
          const number = operandCell(0);
          if (number.type !== "value") {
            return number;
          }
          const digits = operandCell(1);
          if (digits.type !== "value") {
            return digits;
          }
          return roundValue(name, mode, number.result, digits.result, context);
        },
      });
    },
  };
}

/**
 * ROUND(number, num_digits)
 *
 * Rounds a number to `num_digits` decimal places, with halves rounded away
 * from zero: `ROUND(2.5, 0)` is 3 and `ROUND(-2.5, 0)` is -3. A negative
 * `num_digits` rounds to the left of the decimal point: `ROUND(1234.5, -2)` is
 * 1200. `num_digits` is truncated to a whole number.
 *
 * - The number is rounded as Excel shows it, in decimal with 15 significant
 *   digits, so `ROUND(2.675, 2)` is 2.68 although the binary value of 2.675 is
 *   slightly below it (see `round-utils.ts`).
 * - Both arguments must be numbers, like INT and CEILING: text, logical values
 *   and empty cells, which the engine reads as empty text, give #VALUE!. Excel
 *   converts `ROUND("2.5", 0)` to 3, `ROUND(TRUE, 0)` to 1 and reads an empty
 *   cell as 0.
 * - INFINITY and -INFINITY are returned unchanged. A `num_digits` of INFINITY
 *   leaves the number unchanged too, and -INFINITY rounds it to a power of ten
 *   larger than any number: 0 for ROUND and ROUNDDOWN.
 * - A result too large for a number, such as `ROUNDUP(5, -400)` or ROUNDUP of
 *   a non-zero number to -INFINITY places, is #NUM!.
 * - Errors propagate; an error in `number` wins over one in `num_digits`.
 * - Both arguments are required. An omitted argument (a comma with nothing) is
 *   0, as in Excel: `ROUND(2.5,)` is 3.
 * - Array arguments are rounded element by element and broadcast against each
 *   other like the operands of an operator (see `evaluator/broadcast.ts`):
 *   `ROUND(A1:A3, 1)` spills three values and `ROUND(A1, {0,1,2})` three
 *   roundings of one number. An error in one element only affects that element.
 *   Aggregates such as `SUM(ROUND(A1:A3, 0))` read the rounded populated cells.
 */
export const ROUND = createRoundingFunction("ROUND", "half-away-from-zero");
