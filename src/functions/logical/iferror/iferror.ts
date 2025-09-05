import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
  type ErrorEvaluationResult,
} from "src/core/types";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

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
 * - Supports spilled values (dynamic arrays) for both arguments
 * - Any error type (#VALUE!, #DIV/0!, #N/A, etc.) triggers the error handling
 */

/**
 * Helper for creating spilled-values result for IFERROR function
 */
function createIfErrorSpilledResult(
  this: FormulaEvaluator,
  {
    valueResult,
    valueIfErrorResult,
    context,
  }: {
    valueResult: FunctionEvaluationResult;
    valueIfErrorResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): SpilledValuesEvaluationResult | ErrorEvaluationResult {
  const hasSpilledValue = valueResult.type === "spilled-values";
  const hasSpilledError = valueIfErrorResult.type === "spilled-values";

  if (!hasSpilledValue && !hasSpilledError) {
    throw new Error("createIfErrorSpilledResult called without spilled values");
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress): SpreadsheetRange => {
      // Calculate spill area (union of spilled ranges)
      if (hasSpilledValue && valueResult.type === "spilled-values") {
        const spillArea = valueResult.spillArea(origin);
        if (hasSpilledError && valueIfErrorResult.type === "spilled-values") {
          return this.unionRanges(
            this.projectRange(spillArea, origin),
            this.projectRange(valueIfErrorResult.spillArea(origin), origin)
          );
        }
        return spillArea;
      } else if (hasSpilledError && valueIfErrorResult.type === "spilled-values") {
        return valueIfErrorResult.spillArea(origin);
      } else {
        throw new Error("No spilled values found");
      }
    },
    source: "IFERROR with spilled values",
    evaluate: (spilledCell: any, evalContext: any): SingleEvaluationResult => {
      // Evaluate arguments at this spilled position
      const spillValueResult = hasSpilledValue
        ? valueResult.evaluate(spilledCell, evalContext)
        : valueResult;
      const spillErrorResult = hasSpilledError
        ? valueIfErrorResult.evaluate(spilledCell, evalContext)
        : valueIfErrorResult;

      if (spillValueResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled value has not been evaluated",
        };
      }

      if (spillErrorResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled error value has not been evaluated",
        };
      }

      // IFERROR logic: if first argument is error, return second argument
      if (spillValueResult.type === "error") {
        return spillErrorResult;
      }

      // If first argument is not an error, return it
      return spillValueResult;
    },
    evaluateAllCells: (intersectingRange: any) => {
      throw new Error("WIP: evaluateAllCells for IFERROR is not implemented");
    },
  };
}

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
      };
    }

    // Evaluate the value argument (the expression to check for errors)
    const valueResult = this.evaluateNode(node.args[0]!, context);
    
    // Evaluate the value_if_error argument
    const valueIfErrorResult = this.evaluateNode(node.args[1]!, context);
    if (valueIfErrorResult.type === "error") {
      return valueIfErrorResult; // Error in the error handler itself
    }

    // Handle spilled values
    const hasSpilledValues = 
      valueResult.type === "spilled-values" ||
      valueIfErrorResult.type === "spilled-values";

    if (hasSpilledValues) {
      return createIfErrorSpilledResult.call(this, {
        valueResult,
        valueIfErrorResult,
        context,
      });
    }

    // IFERROR logic: if first argument is error, return second argument
    if (valueResult.type === "error") {
      return valueIfErrorResult;
    }

    // If first argument is not an error, return it unchanged
    return valueResult;
  },
};
