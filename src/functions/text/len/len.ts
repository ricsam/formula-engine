import {
  FormulaError,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type CellNumber,
  type CellAddress,
  type SpreadsheetRange,
} from "src/core/types";
import { convertToString } from "../text-helpers";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import type { EvaluationContext } from "src/evaluator/evaluation-context";

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
        errAddress: context.dependencyNode,
      };
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

    if (textResult.result.type === "string") {
      return {
        type: "value",
        result: { type: "number", value: textResult.result.value.length },
      };
    }

    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "LEN operation failed",
      errAddress: context.dependencyNode,
    };
  },
};

/**
 * Helper for creating spilled-values result for LEN function
 */
function createLenSpilledResult(
  this: FormulaEvaluator,
  {
    textResult,
    context,
  }: {
    textResult: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): FunctionEvaluationResult {
  if (textResult.type !== "spilled-values") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "createLenSpilledResult called without spilled values",
      errAddress: context.dependencyNode,
    };
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress) => textResult.spillArea(origin),
    source: "LEN with spilled text values",
    evaluate: (spillOffset, evalContext) => {
      const spillTextResult = textResult.evaluate(spillOffset, evalContext);
      if (
        !spillTextResult ||
        spillTextResult.type === "error" ||
        spillTextResult.type === "awaiting-evaluation"
      ) {
        return spillTextResult;
      }

      if (
        spillTextResult.type !== "value" ||
        spillTextResult.result.type !== "string"
      ) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "LEN operation failed",
          errAddress: context.dependencyNode,
        };
      }

      return {
        type: "value",
        result: { type: "number", value: spillTextResult.result.value.length },
      };
    },
    evaluateAllCells: (intersectingRange) => {
      throw new Error("WIP: evaluateAllCells for LEN is not implemented");
    },
  };
}
