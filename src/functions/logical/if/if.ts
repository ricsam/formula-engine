import {
  FormulaError,
  type CellValue,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import {
  createBroadcastResult,
  toBroadcastOperand,
} from "../../../evaluator/broadcast";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { ASTNode } from "../../../parser/ast";

/**
 * IF function - Returns one value if a condition is true and another value if it's false
 *
 * Usage: IF(logical_test, value_if_true, [value_if_false])
 *
 * logical_test: The condition to evaluate (required)
 * value_if_true: The value to return if the condition is true (required)
 * value_if_false: The value to return if the condition is false (optional, defaults to FALSE)
 *
 * Examples:
 *   IF(A1>10, "High", "Low") - returns "High" if A1>10, otherwise "Low"
 *   IF(B1="", "Empty", B1) - returns "Empty" if B1 is empty, otherwise B1's value
 *   IF(C1<0, "Negative") - returns "Negative" if C1<0, otherwise FALSE
 *   IF(SEQUENCE(1,3), A1:A2) - a 2x3 array: each row repeats A1 or A2
 *
 * Note:
 * - Logical test evaluation: 0 and empty string are FALSE, everything else is TRUE
 * - A single-value test evaluates only the chosen branch and returns it unchanged,
 *   so a range or array branch spills as it is.
 * - An array test is broadcast against the two values like Excel arrays (see
 *   `evaluator/broadcast.ts`): the result is as large as the largest of the three
 *   in each dimension, an argument of size 1 in a dimension is repeated along it,
 *   and cells beyond a shorter argument are #N/A. Each cell reads only the branch
 *   its condition chooses; an error in a single-value branch is that cell's value,
 *   so it only shows up where its branch is chosen. Aggregates that read every
 *   cell at once (SUM(IF(...))) read all three arguments, as Excel does.
 * - An omitted value (`IF(A1,,"x")` or `IF(A1,"x",)`) is 0; an omitted condition
 *   is FALSE.
 */

/**
 * Convert a cell value to boolean for logical evaluation
 */
function isTruthy(value: CellValue): boolean {
  switch (value.type) {
    case "boolean":
      return value.value;
    case "number":
      return value.value !== 0;
    case "string":
      return value.value !== "";
    case "infinity":
      return true; // Infinity is truthy
    default:
      return false;
  }
}

/**
 * Evaluates an IF argument. A missing argument (no comma) gives `whenMissing`;
 * an empty one (a comma without a value, as in `IF(A1,,1)`) gives
 * `whenEmpty`, like Excel.
 */
function evaluateArgument(
  this: FormulaEvaluator,
  arg: ASTNode | undefined,
  context: EvaluationContext,
  whenMissing: CellValue,
  whenEmpty: CellValue
): FunctionEvaluationResult {
  if (!arg) {
    return { type: "value", result: whenMissing };
  }
  if (arg.type === "empty") {
    return { type: "value", result: whenEmpty };
  }
  return this.evaluateNode(arg, context);
}

const FALSE_VALUE: CellValue = { type: "boolean", value: false };
const ZERO_VALUE: CellValue = { type: "number", value: 0 };

/**
 * IF function implementation
 */
export const IF: FunctionDefinition = {
  name: "IF",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "IF function takes 2 or 3 arguments",
        errAddress: context.dependencyNode,
      };
    }

    const logicalTestResult = evaluateArgument.call(
      this,
      node.args[0],
      context,
      FALSE_VALUE,
      FALSE_VALUE
    );
    if (
      logicalTestResult.type === "awaiting-evaluation" ||
      logicalTestResult.type === "error"
    ) {
      return logicalTestResult;
    }

    if (logicalTestResult.type === "value") {
      // LAZY EVALUATION: only evaluate the branch we need, and return it as is.
      return isTruthy(logicalTestResult.result)
        ? evaluateArgument.call(
            this,
            node.args[1],
            context,
            ZERO_VALUE,
            ZERO_VALUE
          )
        : evaluateArgument.call(
            this,
            node.args[2],
            context,
            FALSE_VALUE,
            ZERO_VALUE
          );
    }

    // An array test is broadcast against both values, so both are needed for
    // the shape. Their cells are still read lazily, per chosen branch.
    const valueIfTrueResult = evaluateArgument.call(
      this,
      node.args[1],
      context,
      ZERO_VALUE,
      ZERO_VALUE
    );
    if (valueIfTrueResult.type === "awaiting-evaluation") {
      return valueIfTrueResult;
    }
    const valueIfFalseResult = evaluateArgument.call(
      this,
      node.args[2],
      context,
      FALSE_VALUE,
      ZERO_VALUE
    );
    if (valueIfFalseResult.type === "awaiting-evaluation") {
      return valueIfFalseResult;
    }

    return createBroadcastResult({
      operands: [
        toBroadcastOperand(logicalTestResult, context.cellAddress),
        toBroadcastOperand(valueIfTrueResult, context.cellAddress),
        toBroadcastOperand(valueIfFalseResult, context.cellAddress),
      ],
      context,
      source: "IF with an array condition",
      combine: (operandCell) => {
        const test = operandCell(0);
        if (test.type !== "value") {
          return test;
        }
        return operandCell(isTruthy(test.result) ? 1 : 2);
      },
    });
  },
};
