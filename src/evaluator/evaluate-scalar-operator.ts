import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import {
  FormulaError,
  type ArethmeticEvaluator,
  type CellAddress,
  type CellValue,
  type EvaluateAllCellsResult,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
} from "../core/types";
import type { EvaluationContext } from "./evaluation-context";

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
          if (cellValue.result.type === "error") {
            yield cellValue;
          } else {
            yield {
              result: evaluateSingleScalarOperator(
                cellValue.result.result,
                right.result,
                evaluateScalar
              ),
              relativePos: cellValue.relativePos,
            };
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
          if (cellValue.result.type === "error") {
            yield cellValue;
          } else {
            const result = evaluateSingleScalarOperator(
              left.result,
              cellValue.result.result,
              evaluateScalar
            );
            yield {
              result: result,
              relativePos: cellValue.relativePos,
            };
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
        let rightResult: EvaluateAllCellsResult | undefined | void;
        let leftResult: EvaluateAllCellsResult | undefined | void;
        do {
          rightResult = rightResults.next().value;
          leftResult = leftResults.next().value;

          if (!leftResult && rightResult) {
            yield {
              result: {
                type: "error",
                err: FormulaError.REF,
                message: "Left result is not found",
              },
              relativePos: rightResult.relativePos,
            };
          } else if (!rightResult && leftResult) {
            yield {
              result: {
                type: "error",
                err: FormulaError.REF,
                message: "Right result is not found",
              },
              relativePos: leftResult.relativePos,
            };
          } else if (leftResult && leftResult.result.type === "error") {
            yield leftResult;
          } else if (rightResult && rightResult.result.type === "error") {
            yield rightResult;
          } else if (
            leftResult &&
            rightResult &&
            leftResult.result.type === "value" &&
            rightResult.result.type === "value"
          ) {
            yield {
              result: evaluateSingleScalarOperator(
                leftResult.result.result,
                rightResult.result.result,
                evaluateScalar
              ),
              relativePos: leftResult.relativePos,
            };
          } else {
            break;
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
