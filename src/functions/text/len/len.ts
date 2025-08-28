import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type CellNumber,
  type CellAddress,
  type EvaluationContext,
  type SpreadsheetRange,
} from "src/core/types";
import { convertToString } from "../text-helpers";

/**
 * LEN function - Returns the length of a text string
 *
 * Usage: LEN(text)
 *
 * text: The text string whose length you want to find.
 *
 * Example: LEN("Hello World") returns 11
 *
 * Note:
 * - Supports dynamic arrays (spilled values) for the text argument
 * - Strict type checking: text must be string
 * - Returns the number of characters in the text string
 */
export const LEN: FunctionDefinition = {
  name: "LEN",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 1) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "LEN function takes exactly 1 argument",
      };
    }

    // Evaluate the text argument
    const textResult = this.evaluateNode(node.args[0]!, context);
    if (textResult.type === "error") {
      return textResult;
    }

    // Handle spilled-values input
    if (textResult.type === "spilled-values") {
      return createLenSpilledResult.call(this, {
        textResult,
        context,
      });
    }

    // Single value
    if (textResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid text argument",
      };
    }

    // Strict type checking - no coercion
    if (textResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Text argument must be a string",
      };
    }

    try {
      return {
        type: "value",
        result: lenOperation(textResult),
      };
    } catch (error) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "LEN operation failed",
      };
    }
  },
};

/**
 * Core LEN operation
 */
function lenOperation(textResult: FunctionEvaluationResult): CellNumber {
  const textStr = convertToString(textResult);
  return { type: "number", value: textStr.length };
}

/**
 * Helper for creating spilled-values result for LEN function
 */
function createLenSpilledResult(
  this: any,
  {
    textResult,
    context,
  }: {
    textResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): FunctionEvaluationResult {
  if (textResult.type !== "spilled-values") {
    throw new Error("createLenSpilledResult called without spilled values");
  }

  // Calculate origin result
  const originTextResult = {
    type: "value",
    result: textResult.originResult,
  } as FunctionEvaluationResult;

  let originCellValue: CellNumber;
  try {
    originCellValue = lenOperation(originTextResult);
  } catch (error) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "LEN operation failed",
    };
  }

  return {
    type: "spilled-values",
    spillArea: textResult.spillArea,
    spillOrigin: context.currentCell,
    source: "LEN with spilled text values",
    originResult: originCellValue,
    evaluate: (spilledCell: { address: CellAddress; spillOffset: { x: number; y: number } }, evalContext: EvaluationContext) => {
      const spillTextResult = textResult.evaluate(spilledCell, evalContext);
      if (!spillTextResult || spillTextResult.type === "error") {
        return spillTextResult;
      }

      try {
        return {
          type: "value",
          result: lenOperation(spillTextResult),
        };
      } catch (error) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "LEN operation failed",
        };
      }
    },
  };
}
