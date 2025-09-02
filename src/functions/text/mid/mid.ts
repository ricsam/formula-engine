import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "src/core/types";
import { midOperation, createMidSpilledResult } from "../text-helpers";

/**
 * MID function - Returns characters from the middle of a text string
 *
 * Usage: MID(text, start_num, num_chars)
 *
 * text: The text string to extract characters from.
 * start_num: The position of the first character to extract (1-based).
 * num_chars: The number of characters to extract.
 *
 * Example: MID("Hello, World!", 8, 5) returns "World"
 *
 * Note:
 * - If start_num is less than 1, the function returns an error.
 * - If num_chars is less than 0, the function returns an error.
 * - If start_num is greater than the length of text, the function returns an empty string.
 * - Supports dynamic arrays (spilled values) for all arguments
 * - Strict type checking: text must be string, start_num and num_chars must be numbers
 */
export const MID: FunctionDefinition = {
  name: "MID",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MID function takes exactly 3 arguments",
      };
    }

    // Evaluate the text argument
    const textResult = this.evaluateNode(node.args[0]!, context);
    if (textResult.type === "error") {
      return textResult;
    }

    // Evaluate the start_num argument
    const startNumResult = this.evaluateNode(node.args[1]!, context);
    if (startNumResult.type === "error") {
      return startNumResult;
    }

    // Evaluate the num_chars argument
    const numCharsResult = this.evaluateNode(node.args[2]!, context);
    if (numCharsResult.type === "error") {
      return numCharsResult;
    }

    // Handle spilled-values inputs
    if (
      textResult.type === "spilled-values" ||
      startNumResult.type === "spilled-values" ||
      numCharsResult.type === "spilled-values"
    ) {
      return createMidSpilledResult.call(this, {
        textResult,
        startNumResult,
        numCharsResult,
        context,
      });
    }

    // All arguments are single values
    if (textResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid text argument",
      };
    }

    if (startNumResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid start_num argument",
      };
    }

    if (numCharsResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid num_chars argument",
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

    if (startNumResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Start_num argument must be a number",
      };
    }

    if (numCharsResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Num_chars argument must be a number",
      };
    }

    const result = midOperation(textResult, startNumResult, numCharsResult);
    if (result.type === "error") {
      return result;
    }
    return {
      type: "value",
      result: result.result,
    };
  },
};
