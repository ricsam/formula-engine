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

const ZERO: FunctionEvaluationResult = {
  type: "value",
  result: { type: "number", value: 0 },
};

/**
 * From this quotient on, a number can no longer tell consecutive multiples of
 * the divisor apart (2^53 is where whole numbers stop being exact), so the
 * remainder says nothing about the value that was entered or computed.
 */
const MAX_QUOTIENT = 2 ** 53;

function describeValue(value: CellValue): string {
  return value.type === "string" && value.value === ""
    ? "an empty value"
    : value.type;
}

/**
 * The remainder of `number / divisor` with the sign of the divisor, for finite
 * arguments, a divisor other than 0 and a quotient below MAX_QUOTIENT.
 */
function floorRemainder(number: number, divisor: number): number {
  // A number that is `divisor * k` in this engine's arithmetic, for a whole k,
  // has no remainder: MOD(0.5, 0.1) is 0 because 0.1*5 is 0.5, although the
  // binary value of 0.5 is slightly less than five times that of 0.1.
  if (divisor * Math.round(number / divisor) === number) {
    return 0;
  }
  // Any other remainder is that of the stored binary values, which `%`
  // computes exactly with the sign of `number`; adding the divisor once gives
  // it the divisor's sign. That sum is rounded, so a remainder within half a
  // unit of the divisor's last digit, as in MOD(-1E-20, 1), is the divisor.
  const remainder = number % divisor;
  if (remainder === 0) {
    // Exact multiples are caught above; this keeps any -0 from `%` out.
    return 0;
  }
  return remainder < 0 !== divisor < 0 ? remainder + divisor : remainder;
}

/**
 * The remainder of one number divided by one divisor, or why there is none.
 */
function modValue(
  number: CellValue,
  divisor: CellValue,
  context: EvaluationContext
): SingleEvaluationResult {
  const error = (err: FormulaError, message: string): SingleEvaluationResult => ({
    type: "error",
    err,
    message,
    errAddress: context.dependencyNode,
  });

  if (number.type !== "number" && number.type !== "infinity") {
    return error(
      FormulaError.VALUE,
      `MOD number must be a number, got ${describeValue(number)}`
    );
  }
  if (divisor.type !== "number" && divisor.type !== "infinity") {
    return error(
      FormulaError.VALUE,
      `MOD divisor must be a number, got ${describeValue(divisor)}`
    );
  }
  if (divisor.type === "number" && divisor.value === 0) {
    return error(FormulaError.DIV0, "MOD: the divisor is 0");
  }
  if (number.type === "infinity" || divisor.type === "infinity") {
    // n - d*INT(n/d), Excel's definition of MOD, is #NUM! in this engine for
    // an infinite n or d: it multiplies an infinity by 0 or subtracts one
    // infinity from another.
    return error(
      FormulaError.NUM,
      "MOD: the number and the divisor must be finite"
    );
  }
  if (!(Math.abs(number.value / divisor.value) < MAX_QUOTIENT)) {
    return error(
      FormulaError.NUM,
      "MOD: the number is too large for its remainder by the divisor to be exact"
    );
  }

  return {
    type: "value",
    result: {
      type: "number",
      value: floorRemainder(number.value, divisor.value),
    },
  };
}

/**
 * MOD(number, divisor)
 *
 * Returns the remainder after dividing `number` by `divisor`. The result has
 * the sign of the divisor, as in Excel: `MOD(3, 2)` and `MOD(-3, 2)` are 1,
 * `MOD(3, -2)` and `MOD(-3, -2)` are -1. `MOD(A1, 1)` is the fractional part of
 * a positive number, such as the time of a date-time serial number.
 *
 * - A number that is the divisor times a whole number in this engine's
 *   arithmetic has a remainder of 0: `MOD(0.5, 0.1)` is 0 because `0.1*5` is
 *   0.5. Other remainders are those of the stored binary values, so
 *   `MOD(0.3, 0.1)` is 0.09999999999999998, just as `0.1*3=0.3` is FALSE.
 * - A divisor of 0 is #DIV/0!, as in Excel, also for an infinite number.
 * - Both arguments must be numbers, like ROUND, INT and CEILING: text, logical
 *   values and empty cells, which the engine reads as empty text, give
 *   #VALUE!. Excel converts `MOD("7", 2)` and `MOD(TRUE, 2)` to 1, reads an
 *   empty number as 0 and an empty divisor as #DIV/0!.
 * - INFINITY and -INFINITY in either argument give #NUM!, which is what this
 *   engine computes for Excel's definition `n - d*INT(n/d)` with them.
 * - A quotient `number / divisor` of 2^53 or more in size is #NUM!, where a
 *   number no longer tells consecutive multiples of the divisor apart. Excel
 *   also returns #NUM! for very large numbers.
 * - Errors propagate; an error in `number` wins over one in `divisor`.
 * - Both arguments are required. An omitted argument (a comma with nothing) is
 *   0, as in Excel: `MOD(, 3)` is 0 and `MOD(5, )` is #DIV/0!.
 * - Array arguments are divided element by element and broadcast against each
 *   other like the operands of an operator (see `evaluator/broadcast.ts`):
 *   `MOD(SEQUENCE(6), 3)` spills 1, 2, 0, 1, 2, 0 and `MOD(10, {3,4})` spills
 *   1 and 2. An error in one element only affects that element. Aggregates
 *   such as `SUM(MOD(A1:A9, 2))` read the remainders of the populated cells.
 */
export const MOD: FunctionDefinition = {
  name: "MOD",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MOD function takes exactly 2 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate both arguments before returning either, so that one pass
    // registers all dependencies. An omitted argument is 0, as in Excel.
    const numberNode = getSuppliedArgument(node, 0);
    const divisorNode = getSuppliedArgument(node, 1);
    const numberResult = numberNode
      ? this.evaluateNode(numberNode, context)
      : ZERO;
    const divisorResult = divisorNode
      ? this.evaluateNode(divisorNode, context)
      : ZERO;
    if (
      numberResult.type === "error" ||
      numberResult.type === "awaiting-evaluation"
    ) {
      return numberResult;
    }
    if (
      divisorResult.type === "error" ||
      divisorResult.type === "awaiting-evaluation"
    ) {
      return divisorResult;
    }

    if (numberResult.type === "value" && divisorResult.type === "value") {
      return modValue(numberResult.result, divisorResult.result, context);
    }

    return createBroadcastResult({
      operands: [
        toBroadcastOperand(numberResult, context.cellAddress),
        toBroadcastOperand(divisorResult, context.cellAddress),
      ],
      context,
      source: "MOD function",
      combine: (operandCell) => {
        const number = operandCell(0);
        if (number.type !== "value") {
          return number;
        }
        const divisor = operandCell(1);
        if (divisor.type !== "value") {
          return divisor;
        }
        return modValue(number.result, divisor.result, context);
      },
    });
  },
};
