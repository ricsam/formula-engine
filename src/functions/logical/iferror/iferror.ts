import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { captureEvaluationErrors } from "../../../core/utils";
import {
  createBroadcastResult,
  toBroadcastOperand,
} from "../../../evaluator/broadcast";
import { getSuppliedArgument } from "../../../parser/ast";

/**
 * IFERROR function - Returns a value you specify if a formula evaluates to an error
 *
 * Usage: IFERROR(value, value_if_error)
 *
 * value: The expression to check for errors (required)
 * value_if_error: The value to return if the expression is an error (required)
 *
 * Examples:
 *   IFERROR(A1/B1, "Division error") - returns "Division error" if B1 is 0
 *   IFERROR(VLOOKUP(A1, table, 2), "Not found") - returns "Not found" if lookup fails
 *   IFERROR(1/0, 0) - returns 0 instead of #DIV/0! error
 *
 * Note:
 * - If value is not an error, returns the value unchanged
 * - If value is an error, returns value_if_error
 * - Any error type (#VALUE!, #DIV/0!, #N/A, etc.) triggers the error handling
 * - value_if_error is only evaluated when value is a single error or an array
 * - An array value is checked cell by cell, and an array value_if_error is
 *   broadcast against it like Excel arrays (see `evaluator/broadcast.ts`): a
 *   dimension of size 1 is repeated and cells beyond a shorter array are #N/A.
 *   An error in a single value_if_error only shows in the cells that need it.
 * - An omitted argument (a comma with nothing) is 0, as in Excel:
 *   `IFERROR(1/"x",)` is 0
 */

const ZERO: FunctionEvaluationResult = {
  type: "value",
  result: { type: "number", value: 0 },
};

/**
 * IFERROR function implementation
 */
export const IFERROR: FunctionDefinition = {
  name: "IFERROR",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "IFERROR function takes exactly 2 arguments",
        errAddress: context.dependencyNode,
      };
    }

    const evaluateArgument = (index: number): FunctionEvaluationResult => {
      const argument = getSuppliedArgument(node, index);
      return argument
        ? captureEvaluationErrors(context.dependencyNode, () =>
            this.evaluateNode(argument, context)
          )
        : ZERO;
    };

    // Evaluate the value argument (the expression to check for errors)
    const valueResult = evaluateArgument(0);

    if (valueResult.type === "awaiting-evaluation") {
      return valueResult;
    }

    // LAZY EVALUATION: Only evaluate the error handler if we have an error
    if (valueResult.type === "error") {
      return evaluateArgument(1);
    }

    // If a single value is not an error, return it unchanged WITHOUT evaluating
    // the error handler
    if (valueResult.type === "value") {
      return valueResult;
    }

    // An array may have errors in some cells, so the error handler is needed
    const valueIfErrorResult = evaluateArgument(1);
    if (valueIfErrorResult.type === "awaiting-evaluation") {
      return valueIfErrorResult;
    }

    return createBroadcastResult({
      operands: [
        toBroadcastOperand(valueResult, context.cellAddress),
        toBroadcastOperand(valueIfErrorResult, context.cellAddress),
      ],
      context,
      source: "IFERROR with an array value",
      combine: (operandCell) => {
        const value = operandCell(0);
        return value.type === "error" ? operandCell(1) : value;
      },
    });
  },
};
