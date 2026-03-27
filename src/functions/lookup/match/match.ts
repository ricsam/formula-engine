import {
  FormulaError,
  type CellValue,
  type EvaluateAllCellsResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
} from "../../../core/types";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";

/**
 * MATCH function - Returns the position of a value in an array
 * MATCH(lookup_value, lookup_array, [match_type])
 * match_type: 1 = less than or equal (default), 0 = exact match, -1 = greater than or equal
 *
 * STRICT TYPE CHECKING:
 * - lookup_value: string or number only
 * - lookup_array: array of strings or numbers only
 * - match_type: number only (must be -1, 0, or 1)
 */

// Helper function to perform MATCH operation
function matchOperation(
  lookupValue: CellValue,
  lookupArray: EvaluateAllCellsResult,
  matchType: number,
  context: EvaluationContext,
  isHorizontal: boolean
): FunctionEvaluationResult {
  // Strict type checking for lookup_value
  if (lookupValue.type !== "string" && lookupValue.type !== "number") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `MATCH lookup_value must be string or number, got ${lookupValue.type}`,
      errAddress: context.dependencyNode,
    };
  }

  // Validate match_type
  if (matchType !== -1 && matchType !== 0 && matchType !== 1) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `MATCH match_type must be -1, 0, or 1, got ${matchType}`,
      errAddress: context.dependencyNode,
    };
  }
  if (
    lookupArray.type === "awaiting-evaluation" ||
    lookupArray.type === "error"
  ) {
    return lookupArray;
  }

  for (const value of lookupArray.values) {
    if (value.result.type === "value") {
      if (matchType === 0) {
        // Exact match (case-insensitive for strings, matching Excel behavior)
        const arrayValue = value.result.result;

        const isMatch =
          arrayValue.type === lookupValue.type &&
          (arrayValue.type === "string" && lookupValue.type === "string"
            ? arrayValue.value.toLowerCase() === lookupValue.value.toLowerCase()
            : arrayValue.value === lookupValue.value);

        if (isMatch) {
          // For horizontal arrays (single row), use x position (column index)
          // For vertical arrays (single/multiple columns), use y position (row index)
          const position = isHorizontal
            ? value.relativePos.x + 1
            : value.relativePos.y + 1;

          return {
            type: "value",
            result: { type: "number", value: position },
          }; // 1-based index
        }
      } else {
        // Approximate match (1 or -1) - requires sorted array
        // For now, throw an error until we add sorting validation
        // TODO: Add proper approximate match logic with sorted array validation
        throw new Error("MATCH: approximate match not fully implemented");
      }
    }
  }

  return {
    type: "error",
    err: FormulaError.NA,
    message: "MATCH: lookup_value not found in lookup_array",
    errAddress: context.dependencyNode,
  };
}

export const MATCH: FunctionDefinition = {
  name: "MATCH",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MATCH function takes 2 or 3 arguments",
        errAddress: context.dependencyNode,
      };
    }

    // Evaluate lookup_value
    const lookupValueResult = this.evaluateNode(node.args[0]!, context);
    if (
      lookupValueResult.type === "error" ||
      lookupValueResult.type === "awaiting-evaluation"
    ) {
      return lookupValueResult;
    }

    // Evaluate lookup_array
    const lookupArrayResult = this.evaluateNode(node.args[1]!, context);
    if (
      lookupArrayResult.type === "error" ||
      lookupArrayResult.type === "awaiting-evaluation"
    ) {
      return lookupArrayResult;
    }

    // Evaluate match_type (optional, defaults to 1)
    let matchTypeResult: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 1 },
    };

    if (node.args[2]) {
      matchTypeResult = this.evaluateNode(node.args[2], context);
      if (
        matchTypeResult.type === "error" ||
        matchTypeResult.type === "awaiting-evaluation"
      ) {
        return matchTypeResult;
      }
    }

    // Handle spilled arrays for lookup_value and match_type (not lookup_array which is expected to be a range)
    if (
      lookupValueResult.type === "spilled-values" ||
      matchTypeResult.type === "spilled-values"
    ) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MATCH: Spilled array arguments not yet implemented",
        errAddress: context.dependencyNode,
      };
    }

    // Strict type checking for match_type
    if (matchTypeResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `MATCH match_type must be number, got ${matchTypeResult.result.type}`,
        errAddress: context.dependencyNode,
      };
    }

    // Extract lookup_array values
    let lookupArray: EvaluateAllCellsResult = {
      type: "awaiting-evaluation",
      waitingFor: context.dependencyNode,
      errAddress: context.dependencyNode,
    };
    let isHorizontal = false;

    // Handle direct range arguments (like A:A) before extracting lookup_array values

    // Extract lookup_array values for non-infinite ranges
    if (lookupArrayResult.type === "value") {
      // Single value case - treat as array with one element
      lookupArray = {
        type: "values",
        values: [{ result: lookupArrayResult, relativePos: { x: 0, y: 0 } }],
      };
      isHorizontal = false; // Single value, treat as vertical
    } else if (lookupArrayResult.type === "spilled-values") {
      // Validate that lookup_array is 1D (either single row OR single column)
      const spillArea = lookupArrayResult.spillArea(context.cellAddress);
      const startRow = spillArea.start.row;
      const endRow = spillArea.end.row;
      const startCol = spillArea.start.col;
      const endCol = spillArea.end.col;

      // Check if it's a single row (horizontal)
      const isSingleRow = endRow.type === "number" && startRow === endRow.value;

      // Check if it's a single column (vertical)
      const isSingleCol = endCol.type === "number" && startCol === endCol.value;

      // MATCH requires a 1D array - either single row OR single column, not both or neither
      if (!isSingleRow && !isSingleCol) {
        return {
          type: "error",
          err: FormulaError.VALUE,
          message:
            "MATCH lookup_array must be a single row or single column (1D array)",
          errAddress: context.dependencyNode,
        };
      }

      // Cannot be both single row AND single column (that would be a single cell, which is handled)
      if (isSingleRow && isSingleCol) {
        // This is actually a single cell, which is fine
        isHorizontal = false; // Treat single cell as vertical
      } else {
        // Horizontal if it's a single row
        isHorizontal = isSingleRow;
      }

      // Extract values from spilled array
      lookupArray = lookupArrayResult.evaluateAllCells.call(this, {
        context,
        evaluate: lookupArrayResult.evaluate,
        origin: context.cellAddress,
        lookupOrder: "col-major",
      });
    }

    // Perform MATCH operation
    return matchOperation(
      lookupValueResult.result,
      lookupArray,
      Math.floor(matchTypeResult.result.value), // Floor to handle decimal inputs
      context,
      isHorizontal
    );
  },
};
