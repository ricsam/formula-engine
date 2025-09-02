import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import {
  FormulaError,
  type ArethmeticEvaluator,
  type CellAddress,
  type CellValue,
  type EvaluationContext,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../core/types";

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
  evaluateScalar: ArethmeticEvaluator
): SingleEvaluationResult {
  const result = evaluateScalar(leftValue, rightValue);
  if (result.type === "error") {
    return result;
  }
  return { type: "value", result };
}

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
    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => left.spillArea(origin),
      source: `evaulate left spilled range onto right value in scalar operator ${name}`,
      evaluate: (spilled, context) => {
        const evaledLeft = left.evaluate(spilled, context);
        if (!evaledLeft) {
          return;
        }
        if (evaledLeft.type === "error") {
          return evaledLeft;
        }
        return evaluateSingleScalarOperator(
          evaledLeft.result,
          right.result,
          evaluateScalar
        );
      },
      evaluateAllCells: function* (options) {
        for (const cellValue of left.evaluateAllCells.call(this, options)) {
          if (cellValue.type === "error") {
            yield cellValue;
          } else {
            yield evaluateSingleScalarOperator(
              cellValue.result,
              right.result,
              evaluateScalar
            );
          }
        }
      },
    };
  }
  if (right.type === "spilled-values" && left.type === "value") {
    return {
      type: "spilled-values",
      spillArea: (origin: CellAddress) => right.spillArea(origin),
      source: `evaluate right spilled range onto left value in scalar operator ${name}`,
      evaluate: (spilled, context) => {
        const evaledRight = right.evaluate(spilled, context);
        if (!evaledRight) {
          return;
        }
        if (evaledRight.type === "error") {
          return evaledRight;
        }
        return evaluateSingleScalarOperator(
          left.result,
          evaledRight.result,
          evaluateScalar
        );
      },
      evaluateAllCells: function* (options) {
        for (const cellValue of right.evaluateAllCells.call(this, options)) {
          if (cellValue.type === "error") {
            yield cellValue;
          } else {
            yield evaluateSingleScalarOperator(
              left.result,
              cellValue.result,
              evaluateScalar
            );
          }
        }
      },
    };
  }

  if (left.type === "spilled-values" && right.type === "spilled-values") {
    return {
      type: "spilled-values",
      spillArea: (origin) =>
        this.unionRanges(left.spillArea(origin), right.spillArea(origin)),
      source: `evaluate spilled ranges in scalar operator ${name}`,
      evaluate: (spilled, context) => {
        const evaledLeft = left.evaluate(spilled, context);
        if (!evaledLeft) {
          return;
        }
        if (evaledLeft.type === "error") {
          return evaledLeft;
        }
        const evaledRight = right.evaluate(spilled, context);
        if (!evaledRight) {
          return;
        }
        if (evaledRight.type === "error") {
          return evaledRight;
        }
        return evaluateSingleScalarOperator(
          evaledLeft.result,
          evaledRight.result,
          evaluateScalar
        );
      },
      evaluateAllCells: function* (options) {
        const leftResults = left.evaluateAllCells.call(this, options);
        const rightResults = right.evaluateAllCells.call(this, options);
        let rightResult: SingleEvaluationResult | undefined | void;
        let leftResult: SingleEvaluationResult | undefined | void;
        do {
          rightResult = rightResults.next().value;
          leftResult = leftResults.next().value;

          if (!leftResult) {
            yield {
              type: "error",
              err: FormulaError.REF,
              message: "Left result is not found",
            };
          } else if (!rightResult) {
            yield {
              type: "error",
              err: FormulaError.REF,
              message: "Right result is not found",
            };
          } else if (leftResult.type === "error") {
            yield leftResult;
          } else if (rightResult.type === "error") {
            yield rightResult;
          } else {
            yield evaluateSingleScalarOperator(
              leftResult.result,
              rightResult.result,
              evaluateScalar
            );
          }
        } while (leftResult || rightResult);
      },
    };
  }

  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `Can't evaluate (${left.type}, ${right.type}) in scalar operator ${name}`,
  };
}
