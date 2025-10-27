import type { FormulaEvaluator } from "./formula-evaluator";
import {
  FormulaError,
  type ArethmeticEvaluator,
  type CellAddress,
  type CellValue,
  type CellInRangeResult,
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
  evaluateScalar: ArethmeticEvaluator,
  context: EvaluationContext
): SingleEvaluationResult {
  const result = evaluateScalar(leftValue, rightValue, context);
  if (result.type === "error" || result.type === "awaiting-evaluation") {
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
  if (left.type === "error" || left.type === "awaiting-evaluation") {
    return left;
  }
  if (right.type === "error" || right.type === "awaiting-evaluation") {
    return right;
  }

  if (left.type === "value" && right.type === "value") {
    const leftValue = left.result;
    const rightValue = right.result;
    const result = evaluateScalar(leftValue, rightValue, options.context);
    if (result.type === "error" || result.type === "awaiting-evaluation") {
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
        if (
          evaledLeft.type === "error" ||
          evaledLeft.type === "awaiting-evaluation"
        ) {
          return evaledLeft;
        }
        return evaluateSingleScalarOperator(
          evaledLeft.result,
          right.result,
          evaluateScalar,
          options.context
        );
      },
      evaluateAllCells: function (options) {
        const leftResults = left.evaluateAllCells.call(this, options);
        if (
          leftResults.type === "error" ||
          leftResults.type === "awaiting-evaluation"
        ) {
          return leftResults;
        }
        const results = leftResults.values.map((cellValue) => {
          if (
            cellValue.result.type === "error" ||
            cellValue.result.type === "awaiting-evaluation"
          ) {
            return cellValue;
          } else {
            return {
              result: evaluateSingleScalarOperator(
                cellValue.result.result,
                right.result,
                evaluateScalar,
                options.context
              ),
              relativePos: cellValue.relativePos,
            };
          }
        });
        return {
          type: "values",
          values: results,
        };
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
        if (
          evaledRight.type === "error" ||
          evaledRight.type === "awaiting-evaluation"
        ) {
          return evaledRight;
        }
        return evaluateSingleScalarOperator(
          left.result,
          evaledRight.result,
          evaluateScalar,
          options.context
        );
      },
      evaluateAllCells: function (options) {
        const rightResults = right.evaluateAllCells.call(this, options);
        if (
          rightResults.type === "error" ||
          rightResults.type === "awaiting-evaluation"
        ) {
          return rightResults;
        }
        const results = rightResults.values.map((cellValue) => {
          if (
            cellValue.result.type === "error" ||
            cellValue.result.type === "awaiting-evaluation"
          ) {
            return cellValue;
          } else {
            const result = evaluateSingleScalarOperator(
              left.result,
              cellValue.result.result,
              evaluateScalar,
              options.context
            );
            return {
              result: result,
              relativePos: cellValue.relativePos,
            };
          }
        });
        return {
          type: "values",
          values: results,
        };
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
        if (
          evaledLeft.type === "error" ||
          evaledLeft.type === "awaiting-evaluation"
        ) {
          return evaledLeft;
        }
        const evaledRight = right.evaluate(spilled, context);
        if (
          evaledRight.type === "error" ||
          evaledRight.type === "awaiting-evaluation"
        ) {
          return evaledRight;
        }

        // Check if either value is an empty cell (empty string)
        const isLeftEmpty =
          evaledLeft.type === "value" &&
          evaledLeft.result.type === "string" &&
          evaledLeft.result.value === "";
        const isRightEmpty =
          evaledRight.type === "value" &&
          evaledRight.result.type === "string" &&
          evaledRight.result.value === "";

        // If either operand is empty, return #N/A
        if (isLeftEmpty || isRightEmpty) {
          return {
            type: "error",
            err: FormulaError.NA,
            message: "Empty cell in scalar operation",
            errAddress: options.context.dependencyNode,
          };
        }

        return evaluateSingleScalarOperator(
          evaledLeft.result,
          evaledRight.result,
          evaluateScalar,
          options.context
        );
      },
      evaluateAllCells: function (options) {
        const leftResults = left.evaluateAllCells.call(this, options);
        if (
          leftResults.type === "error" ||
          leftResults.type === "awaiting-evaluation"
        ) {
          return leftResults;
        }
        const rightResults = right.evaluateAllCells.call(this, options);
        if (
          rightResults.type === "error" ||
          rightResults.type === "awaiting-evaluation"
        ) {
          return rightResults;
        }

        // Create position-based maps for both left and right results
        const leftMap = new Map<string, CellInRangeResult>();
        const rightMap = new Map<string, CellInRangeResult>();
        const allPositions = new Set<string>();

        for (const result of leftResults.values) {
          const key = `${result.relativePos.x},${result.relativePos.y}`;
          leftMap.set(key, result);
          allPositions.add(key);
        }

        for (const result of rightResults.values) {
          const key = `${result.relativePos.x},${result.relativePos.y}`;
          rightMap.set(key, result);
          allPositions.add(key);
        }

        const results: CellInRangeResult[] = [];

        // Process each unique position
        for (const posKey of allPositions) {
          const leftResult = leftMap.get(posKey);
          const rightResult = rightMap.get(posKey);

          // Extract relative position from the key
          const [x, y] = posKey.split(",").map(Number);
          const relativePos = { x: x!, y: y! };

          if (!leftResult && !rightResult) {
            // Both empty - this shouldn't happen as we only iterate positions that exist
            continue;
          } else if (!leftResult && rightResult) {
            // Left is empty/missing, right has value
            if (
              rightResult.result.type === "error" ||
              rightResult.result.type === "awaiting-evaluation"
            ) {
              results.push(rightResult);
            } else if (rightResult.result.type === "value") {
              // Treat missing left as empty, which is handled by the operator
              results.push({
                result: {
                  type: "error",
                  err: FormulaError.NA,
                  message: "Left operand is empty",
                  errAddress: context.dependencyNode,
                },
                relativePos,
              });
            } else {
              results.push(rightResult);
            }
          } else if (!rightResult && leftResult) {
            // Right is empty/missing, left has value
            if (
              leftResult.result.type === "error" ||
              leftResult.result.type === "awaiting-evaluation"
            ) {
              results.push(leftResult);
            } else if (leftResult.result.type === "value") {
              // Treat missing right as empty, which is handled by the operator
              results.push({
                result: {
                  type: "error",
                  err: FormulaError.NA,
                  message: "Right operand is empty",
                  errAddress: context.dependencyNode,
                },
                relativePos,
              });
            } else {
              results.push(leftResult);
            }
          } else if (leftResult && rightResult) {
            // Both have values
            if (
              leftResult.result.type === "error" ||
              leftResult.result.type === "awaiting-evaluation"
            ) {
              results.push(leftResult);
            } else if (
              rightResult.result.type === "error" ||
              rightResult.result.type === "awaiting-evaluation"
            ) {
              results.push(rightResult);
            } else if (
              leftResult.result.type === "value" &&
              rightResult.result.type === "value"
            ) {
              results.push({
                result: evaluateSingleScalarOperator(
                  leftResult.result.result,
                  rightResult.result.result,
                  evaluateScalar,
                  options.context
                ),
                relativePos: leftResult.relativePos,
              });
            } else {
              // Both are awaiting-evaluation or some other state
              results.push({
                result: {
                  type: "error",
                  err: FormulaError.VALUE,
                  message:
                    "Cannot evaluate scalar operator on non-value results",
                  errAddress: context.dependencyNode,
                },
                relativePos,
              });
            }
          }
        }

        return { type: "values", values: results };
      },
    };
  }

  return {
    type: "error",
    err: FormulaError.VALUE,
    message: `Can't evaluate (${left.type}, ${right.type}) in scalar operator ${name}`,
    errAddress: options.context.dependencyNode,
  };
}
