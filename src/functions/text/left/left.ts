import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import { midOperation, createMidSpilledResult } from "../text-helpers";

/**
 * LEFT function - Returns the leftmost characters from a text string
 *
 * Usage: LEFT(text, num_chars)
 *
 * text: The text string to extract characters from.
 * num_chars: The number of characters to extract from the left side of the text.
 *
 * Example: LEFT("Hello, World!", 5) returns "Hello"
 *
 * Note:
 * - If num_chars is less than 0, the function returns an error.
 * - If num_chars is greater than the length of text, the function returns the entire text.
 * - Supports dynamic arrays (spilled values) for both arguments
 * - Strict type checking: text must be string, num_chars must be number
 * - Implemented as MID(text, 1, num_chars)
 */
export const LEFT: FunctionDefinition = {
  name: "LEFT",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 1 || node.args.length > 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "LEFT function takes 1 or 2 arguments",
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

    // Create start_num argument (always 1 for LEFT)
    const startNumResult: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 1 },
    };

    // Handle spilled-values inputs using MID's spilled result handler
    if (textResult.type === "spilled-values" || numCharsResult.type === "spilled-values") {
      return createMidSpilledResult.call(this, {
        textResult,
        startNumResult,
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

    // Use MID operation: LEFT(text, num_chars) = MID(text, 1, num_chars)
    try {
      return {
        type: "value",
        result: midOperation(textResult, startNumResult, numCharsResult),
      };
    } catch (error) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "LEFT operation failed",
      };
    }
  },
};