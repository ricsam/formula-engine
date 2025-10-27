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
} from "../../../core/types";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";

/**
 * EXACT function - Compares two text strings and returns TRUE if they are exactly the same, FALSE otherwise
 * 
 * Usage: EXACT(text1, text2)
 * 
 * text1: Required. The first text string
 * text2: Required. The second text string
 * 
 * Examples:
 *   EXACT("Apple", "Apple") - returns TRUE
 *   EXACT("Apple", "apple") - returns FALSE (case-sensitive)
 *   EXACT("123", "123") - returns TRUE
 *   EXACT("", "") - returns TRUE (empty strings are equal)
 * 
 * Note:
 * - The comparison is case-sensitive
 * - Non-text values are converted to text before comparison
 * - Supports spilled values (dynamic arrays) for both arguments
 */

/**
 * Convert a cell value to string for text comparison
 */
function cellValueToString(value: CellValue): string {
  switch (value.type) {
    case "string":
      return value.value;
    case "number":
      return value.value.toString();
    case "boolean":
      return value.value ? "TRUE" : "FALSE";
    case "infinity":
      return value.sign === "positive" ? "INFINITY" : "-INFINITY";
    default:
      return "";
  }
}

/**
 * Helper for creating spilled-values result for EXACT function
 */
function createExactSpilledResult(
  this: FormulaEvaluator,
  {
    text1Result,
    text2Result,
    context,
  }: {
    text1Result: FunctionEvaluationResult;
    text2Result: FunctionEvaluationResult;
    context: EvaluationContext;
  }
): SpilledValuesEvaluationResult | ErrorEvaluationResult {
  const hasSpilledText1 = text1Result.type === "spilled-values";
  const hasSpilledText2 = text2Result.type === "spilled-values";

  if (!hasSpilledText1 && !hasSpilledText2) {
    throw new Error("createExactSpilledResult called without spilled values");
  }

  return {
    type: "spilled-values",
    spillArea: (origin: CellAddress): SpreadsheetRange => {
      // Calculate spill area (union of spilled ranges)
      let spillArea: SpreadsheetRange | undefined;

      const spilledResults = [
        hasSpilledText1 ? text1Result : null,
        hasSpilledText2 ? text2Result : null,
      ].filter(Boolean) as SpilledValuesEvaluationResult[];

      for (const result of spilledResults) {
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

      if (!spillArea) {
        throw new Error("No spilled values found");
      }
      return spillArea;
    },
    source: "EXACT with spilled values",
    evaluate: (spilledCell: any, evalContext: any): SingleEvaluationResult => {
      // Evaluate both arguments at this spilled position
      const spillText1Result = hasSpilledText1
        ? text1Result.evaluate(spilledCell, evalContext)
        : text1Result;
      const spillText2Result = hasSpilledText2
        ? text2Result.evaluate(spilledCell, evalContext)
        : text2Result;

      // Check for errors
      if (spillText1Result === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled text1 argument has not been evaluated",
          errAddress: context.dependencyNode,
        };
      }
      if (spillText1Result.type === "error" || spillText1Result.type === "awaiting-evaluation") {
        return spillText1Result;
      }

      if (spillText2Result === undefined) {
        return {
          type: "error",
          err: FormulaError.REF,
          message: "The spilled text2 argument has not been evaluated",
          errAddress: context.dependencyNode,
        };
      }
      if (spillText2Result.type === "error" || spillText2Result.type === "awaiting-evaluation") {
        return spillText2Result;
      }

      // Perform EXACT comparison
      if (spillText1Result.type !== "value" || spillText2Result.type !== "value") {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message: "Invalid argument types for EXACT function #1",
          errAddress: context.dependencyNode,
        };
      }

      const text1 = cellValueToString(spillText1Result.result);
      const text2 = cellValueToString(spillText2Result.result);

      return {
        type: "value",
        result: { type: "boolean", value: text1 === text2 },
      };
    },
    evaluateAllCells: (intersectingRange: any) => {
      throw new Error("WIP: evaluateAllCells for EXACT is not implemented");
    },
  };
}

/**
 * EXACT function implementation
 */
export const EXACT: FunctionDefinition = {
  name: "EXACT",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length !== 2) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "EXACT function takes exactly 2 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate both text arguments
    const text1Result = this.evaluateNode(node.args[0]!, context);
    if (text1Result.type === "error" || text1Result.type === "awaiting-evaluation") {
      return text1Result;
    }

    const text2Result = this.evaluateNode(node.args[1]!, context);
    if (text2Result.type === "error" || text2Result.type === "awaiting-evaluation") {
      return text2Result;
    }

    // Handle spilled values
    const hasSpilledValues = 
      text1Result.type === "spilled-values" ||
      text2Result.type === "spilled-values";

    if (hasSpilledValues) {
      return createExactSpilledResult.call(this, {
        text1Result,
        text2Result,
        context,
      });
    }

    // Both arguments are single values
    if (text1Result.type !== "value" || text2Result.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "Invalid argument types for EXACT function #2",
        errAddress: context.dependencyNode,
      };
    }

    // Convert both values to strings and compare
    const text1 = cellValueToString(text1Result.result);
    const text2 = cellValueToString(text2Result.result);

    return {
      type: "value",
      result: { type: "boolean", value: text1 === text2 },
    };
  },
};
