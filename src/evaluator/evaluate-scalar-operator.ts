import type { FormulaEvaluator } from "./formula-evaluator";
import {
  type ArethmeticEvaluator,
  type CellValue,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../core/types";
import type { EvaluationContext } from "./evaluation-context";
import { createBroadcastResult, toBroadcastOperand } from "./broadcast";

export type EvaluateScalarOperatorOptions = {
  evaluateScalar: ArethmeticEvaluator;
  context: EvaluationContext;
  /**
   * for debugging messages
   */
  name: string;
};

function evaluateSingleScalarOperator(
  leftValue: CellValue,
  rightValue: CellValue,
  evaluateScalar: ArethmeticEvaluator,
  context: EvaluationContext
): SingleEvaluationResult {
  const result = evaluateScalar(leftValue, rightValue, context);
  if (result.type === "error" || result.type === "awaiting-evaluation") {
    return result;
  }
  return { type: "value", result };
}

/**
 * Applies a binary operator (arithmetic, comparison or `&`) element-wise.
 *
 * Two single values give a single value. Otherwise the operands are broadcast
 * like Excel arrays (see `broadcast.ts`): `SEQUENCE(3)*SEQUENCE(1,4)` is a 3x4
 * table, `A1:A3*10` multiplies every cell, and a cell beyond a shorter operand
 * (`A1:A3+B1:B2` in its third row) is `#N/A`.
 *
 * A blank cell inside an operand reads as the empty string, exactly as a blank
 * cell does in `A1*2`, so the operator decides what it means (`&` treats it as
 * empty text; arithmetic currently rejects it with `#VALUE!`).
 *
 * An error or awaiting-evaluation operand that is a single value is returned as
 * the result; errors inside an array stay per-cell values.
 */
export function evaluateScalarOperator(
  this: FormulaEvaluator,
  left: FunctionEvaluationResult,
  right: FunctionEvaluationResult,
  options: EvaluateScalarOperatorOptions
): FunctionEvaluationResult {
  const { evaluateScalar, context, name } = options;
  if (left.type === "error" || left.type === "awaiting-evaluation") {
    return left;
  }
  if (right.type === "error" || right.type === "awaiting-evaluation") {
    return right;
  }

  if (left.type === "value" && right.type === "value") {
    return evaluateSingleScalarOperator(
      left.result,
      right.result,
      evaluateScalar,
      context
    );
  }

  return createBroadcastResult({
    operands: [
      toBroadcastOperand(left, context.cellAddress),
      toBroadcastOperand(right, context.cellAddress),
    ],
    context,
    source: `scalar operator ${name}`,
    combine: (operandCell) => {
      const leftCell = operandCell(0);
      if (leftCell.type !== "value") {
        return leftCell;
      }
      const rightCell = operandCell(1);
      if (rightCell.type !== "value") {
        return rightCell;
      }
      return evaluateSingleScalarOperator(
        leftCell.result,
        rightCell.result,
        evaluateScalar,
        context
      );
    },
  });
}
