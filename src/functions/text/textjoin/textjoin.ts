import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
  type SingleEvaluationResult,
  type ErrorEvaluationResult,
  type CellString,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * TEXTJOIN function - Joins text from multiple ranges and/or strings with a specified delimiter
 *
 * Usage: TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
 *
 * delimiter: String to insert between text items (can be empty string "")
 * ignore_empty: If TRUE, empty cells/strings are ignored; if FALSE, they are included
 * text1, text2, ...: Text items to join (can be strings, ranges, or arrays)
 *
 * Examples:
 *   TEXTJOIN(", ", TRUE, A1:A3) - joins A1, A2, A3 with ", " separator, ignoring empty cells
 *   TEXTJOIN("-", FALSE, "a", "b", "c") - returns "a-b-c"
 *   TEXTJOIN(", ", TRUE, "Red", "", "Blue") - returns "Red, Blue" (empty string ignored)
 *
 * Note:
 * - Supports type coercion: numbers and booleans are converted to strings
 * - Returns #VALUE! if result exceeds 32,767 characters
 */

const MAX_TEXT_LENGTH = 32767;

/**
 * Helper function to convert a cell value to string with type coercion
 */
function coerceToString(value: CellValue): string {
  switch (value.type) {
    case "string":
      return value.value;
    case "number":
      // Convert number to string
      if (value.value === Infinity) {
        return "INFINITY";
      } else if (value.value === -Infinity) {
        return "-INFINITY";
      } else if (isNaN(value.value)) {
        return "NaN";
      } else {
        return value.value.toString();
      }
    case "boolean":
      // Convert boolean to string
      return value.value ? "TRUE" : "FALSE";
    case "infinity":
      return value.sign === "positive" ? "INFINITY" : "-INFINITY";
    default:
      return "";
  }
}

/**
 * Helper function to collect all text items from arguments, handling ranges
 */
function collectTextItems(
  this: FormulaEvaluator,
  args: FunctionEvaluationResult[],
  ignoreEmpty: boolean,
  context: EvaluationContext
): string[] | ErrorEvaluationResult {
  const textItems: string[] = [];

  for (const arg of args) {
    if (arg.type === "error") {
      return arg;
    }

    if (arg.type === "value") {
      const text = coerceToString(arg.result);
      if (!ignoreEmpty || text !== "") {
        textItems.push(text);
      }
    } else if (arg.type === "spilled-values") {
      // Extract all values from the range
      const cellValues = arg.evaluateAllCells.call(this, {
        context,
        origin: context.cellAddress,
        evaluate: arg.evaluate,
        lookupOrder: "col-major",
      });

      if (cellValues.type !== "values") {
        return cellValues;
      }

      for (const cellValue of cellValues.values) {
        if (
          cellValue.result.type === "error" ||
          cellValue.result.type === "awaiting-evaluation"
        ) {
          return cellValue.result;
        }
        if (cellValue.result.type === "value") {
          const text = coerceToString(cellValue.result.result);
          if (!ignoreEmpty || text !== "") {
            textItems.push(text);
          }
        }
      }
    }
  }

  return textItems;
}

/**
 * TEXTJOIN function implementation
 */
export const TEXTJOIN: FunctionDefinition = {
  name: "TEXTJOIN",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "TEXTJOIN function requires at least 3 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate delimiter (first argument)
    const delimiterResult = this.evaluateNode(node.args[0]!, context);
    if (delimiterResult.type === "error") {
      return delimiterResult;
    }
    if (delimiterResult.type === "awaiting-evaluation") {
      return delimiterResult;
    }

    // Delimiter must be a single value
    if (delimiterResult.type === "spilled-values") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "TEXTJOIN delimiter must be a single value",
        errAddress: context.dependencyNode,
      };
    }

    const delimiter = coerceToString(delimiterResult.result);

    // Evaluate ignore_empty (second argument)
    const ignoreEmptyResult = this.evaluateNode(node.args[1]!, context);
    if (ignoreEmptyResult.type === "error") {
      return ignoreEmptyResult;
    }
    if (ignoreEmptyResult.type === "awaiting-evaluation") {
      return ignoreEmptyResult;
    }

    // ignore_empty must be a single value
    if (ignoreEmptyResult.type === "spilled-values") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "TEXTJOIN ignore_empty must be a single value",
        errAddress: context.dependencyNode,
      };
    }

    // Convert ignore_empty to boolean
    let ignoreEmpty = false;
    if (ignoreEmptyResult.result.type === "boolean") {
      ignoreEmpty = ignoreEmptyResult.result.value;
    } else if (ignoreEmptyResult.result.type === "number") {
      ignoreEmpty = ignoreEmptyResult.result.value !== 0;
    } else if (ignoreEmptyResult.result.type === "string") {
      ignoreEmpty = ignoreEmptyResult.result.value !== "";
    }

    // Evaluate text arguments (remaining arguments)
    const textArgs: FunctionEvaluationResult[] = [];
    for (let i = 2; i < node.args.length; i++) {
      const result = this.evaluateNode(node.args[i]!, context);
      if (result.type === "error" || result.type === "awaiting-evaluation") {
        return result;
      }
      textArgs.push(result);
    }

    // Collect all text items from arguments
    const textItems = collectTextItems.call(
      this,
      textArgs,
      ignoreEmpty,
      context
    );

    // Check if we got an error
    if (!Array.isArray(textItems)) {
      return textItems;
    }

    // Join text items with delimiter
    const result = textItems.join(delimiter);

    // Check length limit
    if (result.length > MAX_TEXT_LENGTH) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `TEXTJOIN result exceeds ${MAX_TEXT_LENGTH} character limit`,
        errAddress: context.dependencyNode,
      };
    }

    return {
      type: "value",
      result: { type: "string", value: result },
    };
  },
};
