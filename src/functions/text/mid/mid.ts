import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import { getSuppliedArgument } from "../../../parser/ast";
import { createMidSpilledResult, midOperation } from "../text-helpers";

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
 * - An omitted start_num or num_chars (a comma with nothing) is 0, as in Excel:
 *   `MID("abc",2,)` is empty text and `MID("abc",,1)` is #VALUE!
 */
export const MID: FunctionDefinition = {
  name: "MID",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MID function takes exactly 3 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate the text argument
    const textResult = this.evaluateNode(node.args[0]!, context);
    if (
      textResult.type === "error" ||
      textResult.type === "awaiting-evaluation"
    ) {
      return textResult;
    }

    // Evaluate the start_num argument (0 when omitted)
    const zero: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 0 },
    };
    const startNumNode = getSuppliedArgument(node, 1);
    const startNumResult = startNumNode
      ? this.evaluateNode(startNumNode, context)
      : zero;
    if (
      startNumResult.type === "error" ||
      startNumResult.type === "awaiting-evaluation"
    ) {
      return startNumResult;
    }

    // Evaluate the num_chars argument (0 when omitted)
    const numCharsNode = getSuppliedArgument(node, 2);
    const numCharsResult = numCharsNode
      ? this.evaluateNode(numCharsNode, context)
      : zero;
    if (
      numCharsResult.type === "error" ||
      numCharsResult.type === "awaiting-evaluation"
    ) {
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

    // Strict type checking - no coercion
    if (textResult.result.type !== "string") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Text argument must be a string",
        errAddress: context.dependencyNode,
      };
    }

    if (startNumResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Start_num argument must be a number",
        errAddress: context.dependencyNode,
      };
    }

    if (numCharsResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Num_chars argument must be a number",
        errAddress: context.dependencyNode,
      };
    }

    const result = midOperation(
      textResult,
      startNumResult,
      numCharsResult,
      context
    );
    if (result.type === "error" || result.type === "awaiting-evaluation") {
      return result;
    }
    return {
      type: "value",
      result: result.result,
    };
  },
};
