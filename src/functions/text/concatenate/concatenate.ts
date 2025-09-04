import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type CellString,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type ErrorEvaluationResult,
  type EvaluationContext,
  type SingleEvaluationResult,
} from "src/core/types";
import { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { FunctionNode } from "src/parser/ast";

/**
 * CONCATENATE function - Joins several text strings into one text string
 *
 * Usage: CONCATENATE(text1, text2, ...)
 *
 * text1, text2, ...: Text strings to be joined together.
 *
 * Example: CONCATENATE("Hello", " ", "World") returns "Hello World"
 *
 * Note:
 * - Takes 1 or more arguments
 * - Supports type coercion: strings, numbers, and booleans are converted to strings
 * - Numbers: converted to string representation (123 -> "123", Infinity -> "INFINITY")
 * - Booleans: TRUE -> "TRUE", FALSE -> "FALSE"
 * - Supports dynamic arrays (spilled values) for arguments
 * - If any argument is a spilled value, the result will be spilled
 */

/**
 * Helper function to convert a cell value to string with type coercion
 */
function coerceToString(
  result: FunctionEvaluationResult
): string | ErrorEvaluationResult {
  if (result.type !== "value") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "Invalid argument type",
    };
  }

  switch (result.result.type) {
    case "string":
      return result.result.value;
    case "number":
      // Convert number to string
      if (result.result.value === Infinity) {
        return "INFINITY";
      } else if (result.result.value === -Infinity) {
        return "-INFINITY";
      } else if (isNaN(result.result.value)) {
        return "NaN";
      } else {
        return result.result.value.toString();
      }
    case "boolean":
      // Convert boolean to string
      return result.result.value ? "TRUE" : "FALSE";
    default:
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Cannot convert argument to string",
      };
  }
}

/**
 * Helper function to perform CONCATENATE operation on evaluated arguments
 */
function concatenateOperation(
  textResults: FunctionEvaluationResult[]
): { type: "value"; result: CellString } | ErrorEvaluationResult {
  let result = "";
  for (const textResult of textResults) {
    const stringValue = coerceToString(textResult);
    if (typeof stringValue === "object" && stringValue.type === "error") {
      return stringValue;
    }
    result += stringValue;
  }
  return { type: "value", result: { type: "string", value: result } };
}

/**
 * Helper for creating spilled-values result for CONCATENATE function
 */
function createConcatenateSpilledResult(
  this: FormulaEvaluator,
  {
    textResults,
    context,
  }: {
    textResults: FunctionEvaluationResult[];
    context: EvaluationContext;
  }
): SpilledValuesEvaluationResult | ErrorEvaluationResult {
  const spilledResults = textResults.filter(
    (result) => result.type === "spilled-values"
  );

  if (spilledResults.length === 0) {
    throw new Error(
      "createConcatenateSpilledResult called without spilled values"
    );
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress): SpreadsheetRange => {
      // Calculate spill area (union of all spilled ranges)
      let spillArea: SpreadsheetRange | undefined;

      for (const result of textResults) {
        if (result.type === "spilled-values") {
          const currentSpillArea = result.spillArea(origin);
          if (!spillArea) {
            spillArea = currentSpillArea;
          } else {
            spillArea = this.unionRanges(
              this.projectRange(spillArea, origin),
              this.projectRange(currentSpillArea, origin)
            );
          }
        }
      }

      if (!spillArea) {
        throw new Error("No spilled values found");
      }
      return spillArea;
    },
    source: "CONCATENATE with spilled values",
    evaluate: (spilledCell, evalContext): SingleEvaluationResult => {
      // Evaluate all arguments at this spilled position
      const spilledTextResults: FunctionEvaluationResult[] = [];

      for (const textResult of textResults) {
        if (textResult.type === "spilled-values") {
          const spillResult = textResult.evaluate(spilledCell, evalContext);
          if (spillResult === undefined) {
            return {
              type: "error",
              err: FormulaError.REF,
              message: "The spilled results have not been evaluated",
            };
          }
          if (spillResult.type === "error") {
            return spillResult;
          }
          spilledTextResults.push(spillResult);
        } else {
          spilledTextResults.push(textResult);
        }
      }

      return concatenateOperation(spilledTextResults);
    },
    evaluateAllCells: (intersectingRange) => {
      throw new Error(
        "WIP: evaluateAllCells for CONCATENATE is not implemented"
      );
    },
  };
}

/**
 * CONCATENATE function implementation
 */
export const CONCATENATE: FunctionDefinition = {
  name: "CONCATENATE",
  evaluate: function (
    node: FunctionNode,
    context: EvaluationContext
  ): FunctionEvaluationResult {
    if (node.args.length < 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "CONCATENATE function requires at least 1 argument",
      };
    }

    // Evaluate all arguments
    const textResults: FunctionEvaluationResult[] = [];
    for (const arg of node.args) {
      const result = this.evaluateNode(arg, context);
      if (result.type === "error") {
        return result;
      }
      textResults.push(result);
    }

    // Check if any arguments are spilled values
    const hasSpilledValues = textResults.some(
      (result) => result.type === "spilled-values"
    );

    if (hasSpilledValues) {
      return createConcatenateSpilledResult.call(this, {
        textResults,
        context,
      });
    }

    // All arguments are single values - type coercion will be handled in concatenateOperation

    // Perform concatenation
    return concatenateOperation(textResults);
  },
};
