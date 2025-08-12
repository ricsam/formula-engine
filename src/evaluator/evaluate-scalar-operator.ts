import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import {
  FormulaError,
  type ArethmeticEvaluator,
  type CellValue,
  type EvaluationContext,
  type FunctionEvaluationResult,
} from "../core/types";

export type EvaluateScalarOperatorOptions = {
  evaluateScalar: ArethmeticEvaluator;
  context: EvaluationContext;
  /**
   * for debugging messages
   */
  name: string;
};

export function evaluateScalarOperator(
  this: FormulaEvaluator,
  left: FunctionEvaluationResult,
  right: FunctionEvaluationResult,
  options: EvaluateScalarOperatorOptions
): FunctionEvaluationResult {
  const { evaluateScalar, context, name } = options;
  if (left.type === "error") {
    return left;
  }
  if (right.type === "error") {
    return right;
  }

  if (left.type === "value" && right.type === "value") {
    const leftValue = left.result;
    const rightValue = right.result;
    const result = evaluateScalar(leftValue, rightValue);
    if (result.type === "error") {
      return result;
    }
    if (result) {
      return { type: "value", result };
    }
  }
  if (left.type === "spilled-values" && right.type === "value") {
    const leftRange = left.spillArea;
    const originResult = evaluateScalar(left.originResult, right.result);

    if (originResult.type === "error") {
      return originResult;
    }

    return {
      type: "spilled-values",
      spillArea: this.projectRange(leftRange, context.currentCell),
      spillOrigin: context.currentCell,
      source: `evaulate left spilled range onto right value in scalar operator ${name}`,
      originResult,
      evaluate: (spilled, context) => {
        const evaledLeft = left.evaluate(spilled, context);
        if (!evaledLeft) {
          return;
        }
        return this.evaluateScalarOperator(evaledLeft, right, options);
      },
    };
  }
  if (right.type === "spilled-values" && left.type === "value") {
    const rightRange = right.spillArea;
    const originResult = evaluateScalar(left.result, right.originResult);

    if (originResult.type === "error") {
      return originResult;
    }

    return {
      type: "spilled-values",
      spillArea: this.projectRange(rightRange, context.currentCell),
      spillOrigin: context.currentCell,
      source: `evaluate right spilled range onto left value in scalar operator ${name}`,
      originResult,
      evaluate: (spilled, context) => {
        const evaledRight = right.evaluate(spilled, context);
        if (!evaledRight) {
          return;
        }
        return this.evaluateScalarOperator(left, evaledRight, options);
      },
    };
  }

  if (left.type === "spilled-values" && right.type === "spilled-values") {
    const rightRange = right.spillArea;
    const leftRange = left.spillArea;

    const originResult = evaluateScalar(left.originResult, right.originResult);

    if (originResult.type === "error") {
      return originResult;
    }

    return {
      type: "spilled-values",
      spillArea: this.unionRanges(
        this.projectRange(leftRange, context.currentCell),
        this.projectRange(rightRange, context.currentCell)
      ),
      spillOrigin: context.currentCell,
      source: `evaluate spilled ranges in scalar operator ${name}`,
      originResult,
      evaluate: (spilled, context) => {
        const evaledLeft = left.evaluate(spilled, context);
        if (!evaledLeft) {
          return;
        }
        const evaledRight = right.evaluate(spilled, context);
        if (!evaledRight) {
          return;
        }
        return this.evaluateScalarOperator(evaledLeft, evaledRight, options);
      },
    };
  }

  const leftVal =
    left.type === "value"
      ? left.result.type === "infinity"
        ? "infinity"
        : `(${left.result.type}, ${left.result.value})`
      : left.type;
  const rightVal =
    right.type === "value"
      ? right.result.type === "infinity"
        ? "infinity"
        : `(${right.result.type}, ${right.result.value})`
      : right.type;

  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `Can't evaluate (${leftVal}, ${rightVal}) in scalar operator ${name}`,
  };
}
