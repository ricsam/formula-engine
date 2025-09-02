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
import {
  midOperation,
  createMidSpilledResult,
  convertToString,
  extractNumericValue,
} from "../text-helpers";

/**
 * RIGHT function - Returns the rightmost characters from a text string
 *
 * Usage: RIGHT(text, num_chars)
 *
 * text: The text string to extract characters from.
 * num_chars: The number of characters to extract from the right side of the text.
 *
 * Example: RIGHT("Hello, World!", 6) returns "World!"
 *
 * Note:
 * - If num_chars is less than 0, the function returns an error.
 * - If num_chars is greater than the length of text, the function returns the entire text.
 * - Supports dynamic arrays (spilled values) for both arguments
 * - Strict type checking: text must be string, num_chars must be number
 * - Implemented as MID(text, LEN(text) - num_chars + 1, num_chars)
 */

/**
 * Helper function to calculate RIGHT operation using MID
 */
function rightOperation(
  textResult: FunctionEvaluationResult,
  numCharsResult: FunctionEvaluationResult
): { type: "value"; result: CellString } | ErrorEvaluationResult {
  const textStr = convertToString(textResult);
  const numChars = extractNumericValue(numCharsResult);

  // Validate numChars
  if (numChars < 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "NumChars argument must be a positive number",
    };
  }

  // Calculate start position: LEN(text) - num_chars + 1
  const textLength = textStr.length;
  const startPos = Math.max(1, textLength - Math.floor(numChars) + 1);

  // Create start_num result
  const startNumResult: FunctionEvaluationResult = {
    type: "value",
    result: { type: "number", value: startPos },
  };

  // Use MID operation
  return midOperation(textResult, startNumResult, numCharsResult);
}

export const RIGHT: FunctionDefinition = {
  name: "RIGHT",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 1 || node.args.length > 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "RIGHT function takes 1 or 2 arguments",
      };
    }

    // Evaluate the text argument
    const textResult = this.evaluateNode(node.args[0]!, context);
    if (textResult.type === "error") {
      return textResult;
    }

    // Evaluate the numChars argument (optional, defaults to 1)
    let numCharsResult: FunctionEvaluationResult;
    if (node.args.length > 1) {
      numCharsResult = this.evaluateNode(node.args[1]!, context);
      if (numCharsResult.type === "error") {
        return numCharsResult;
      }
    } else {
      numCharsResult = {
        type: "value",
        result: { type: "number", value: 1 },
      };
    }

    // Handle spilled-values inputs
    if (
      textResult.type === "spilled-values" ||
      numCharsResult.type === "spilled-values"
    ) {
      // For RIGHT with spilled values, we need a custom spilled result handler
      // because we need to calculate start_num dynamically for each text value
      return createRightSpilledResult.call(this, {
        textResult,
        numCharsResult,
        context,
      });
    }

    // Both arguments are single values
    if (textResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid text argument",
      };
    }

    if (numCharsResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid numChars argument",
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

    if (numCharsResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "NumChars argument must be a number",
      };
    }

    // Use RIGHT operation: RIGHT(text, num_chars) = MID(text, LEN(text) - num_chars + 1, num_chars)
    return rightOperation(textResult, numCharsResult);
  },
};

/**
 * Helper for creating spilled-values result for RIGHT function
 */
function createRightSpilledResult(
  this: FormulaEvaluator,
  {
    textResult,
    numCharsResult,
    context,
  }: {
    textResult: FunctionEvaluationResult;
    numCharsResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): SpilledValuesEvaluationResult | ErrorEvaluationResult {
  if (
    textResult.type !== "spilled-values" &&
    numCharsResult.type !== "spilled-values"
  ) {
    throw new Error("createRightSpilledResult called without spilled values");
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress): SpreadsheetRange => {
      // Calculate spill area
      let spillArea;
      if (textResult.type === "spilled-values") {
        spillArea = textResult.spillArea(origin);
      } else if (numCharsResult.type === "spilled-values") {
        spillArea = numCharsResult.spillArea(origin);
      } else {
        throw new Error("No spilled values found");
      }

      // Union spill areas if both are spilled
      if (
        textResult.type === "spilled-values" &&
        numCharsResult.type === "spilled-values"
      ) {
        spillArea = this.unionRanges(
          this.projectRange(textResult.spillArea(origin), origin),
          this.projectRange(numCharsResult.spillArea(origin), origin)
        );
      }
      return spillArea;
    },
    source: "RIGHT with spilled values",
    evaluate: (spilledCell, evalContext): SingleEvaluationResult => {
      // Evaluate arguments at this spilled position
      const spillTextResult =
        textResult.type === "spilled-values"
          ? textResult.evaluate(spilledCell, evalContext)
          : textResult;
      const spillNumResult =
        numCharsResult.type === "spilled-values"
          ? numCharsResult.evaluate(spilledCell, evalContext)
          : numCharsResult;

      if (spillTextResult === undefined || spillNumResult === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled results have not been evaluated",
        };
      }
      if (spillNumResult.type === "error") {
        return spillNumResult;
      }
      if (spillTextResult.type === "error") {
        return spillTextResult;
      }

      return rightOperation(spillTextResult, spillNumResult);
    },
    evaluateAllCells: (intersectingRange) => {
      throw new Error("WIP: evaluateAllCells for RIGHT is not implemented");
    },
  };
}
