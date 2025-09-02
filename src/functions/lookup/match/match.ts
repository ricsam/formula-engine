import {
  FormulaError,
  type CellValue,
  type CellNumber,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type ValueEvaluationResult,
  type EvaluationContext,
  type CellAddress,
  type SpreadsheetRange,
  type SpilledValuesEvaluationResult,
} from "src/core/types";
import type { FormulaEngine } from "src/core/engine";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import { OpenRangeEvaluator } from "../../math/open-range-evaluator";

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
  lookupArray: Iterable<FunctionEvaluationResult>,
  matchType: number
): FunctionEvaluationResult {
  // Strict type checking for lookup_value
  if (lookupValue.type !== "string" && lookupValue.type !== "number") {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `MATCH lookup_value must be string or number, got ${lookupValue.type}`,
    };
  }

  // Validate match_type
  if (matchType !== -1 && matchType !== 0 && matchType !== 1) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: `MATCH match_type must be -1, 0, or 1, got ${matchType}`,
    };
  }

  const values: CellValue[] = [];

  for (const value of lookupArray) {
    if (value.type === "value") {
      values.push(value.result);
    }
  }

  if (values.length === 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      message: "MATCH lookup_array cannot be empty",
    };
  }

  if (matchType === 0) {
    // Exact match
    for (let i = 0; i < values.length; i++) {
      const arrayValue = values[i]!;
      if (
        arrayValue.type === lookupValue.type &&
        arrayValue.value === lookupValue.value
      ) {
        return { type: "value", result: { type: "number", value: i + 1 } }; // 1-based index
      }
    }
    return {
      type: "error",
      err: FormulaError.NA,
      message: "MATCH: lookup_value not found in lookup_array",
    };
  } else {
    // Approximate match (1 or -1) - requires sorted array
    // For now, implement exact match behavior until we add sorting validation
    // TODO: Add proper approximate match logic with sorted array validation
    for (let i = 0; i < values.length; i++) {
      const arrayValue = values[i]!;
      if (
        arrayValue.type === lookupValue.type &&
        arrayValue.value === lookupValue.value
      ) {
        return { type: "value", result: { type: "number", value: i + 1 } }; // 1-based index
      }
    }
    return {
      type: "error",
      err: FormulaError.NA,
      message:
        "MATCH: lookup_value not found in lookup_array (approximate match not fully implemented)",
    };
  }
}

export const MATCH: FunctionDefinition = {
  name: "MATCH",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 2 || node.args.length > 3) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MATCH function takes 2 or 3 arguments",
      };
    }

    // Evaluate lookup_value
    const lookupValueResult = this.evaluateNode(node.args[0]!, context);
    if (lookupValueResult.type === "error") {
      return lookupValueResult;
    }

    // Evaluate lookup_array
    const lookupArrayResult = this.evaluateNode(node.args[1]!, context);
    if (lookupArrayResult.type === "error") {
      return lookupArrayResult;
    }

    // Evaluate match_type (optional, defaults to 1)
    let matchTypeResult: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 1 },
    };

    if (node.args[2]) {
      matchTypeResult = this.evaluateNode(node.args[2], context);
      if (matchTypeResult.type === "error") {
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
      };
    }

    // Extract values for normal (non-spilled) case
    if (lookupValueResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MATCH: Invalid lookup_value result type",
      };
    }

    if (matchTypeResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: "MATCH: Invalid match_type result type",
      };
    }

    // Strict type checking for match_type
    if (matchTypeResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        message: `MATCH match_type must be number, got ${matchTypeResult.result.type}`,
      };
    }

    // Extract lookup_array values
    let lookupArray: Iterable<FunctionEvaluationResult> = [];

    // Handle direct range arguments (like A:A) before extracting lookup_array values

    // Extract lookup_array values for non-infinite ranges
    if (lookupArrayResult.type === "value") {
      // Single value case - treat as array with one element
      lookupArray = [lookupArrayResult];
    } else if (lookupArrayResult.type === "spilled-values") {
      // Extract values from spilled array
      lookupArray = lookupArrayResult.evaluateAllCells.call(this, {
        context,
        evaluate: lookupArrayResult.evaluate,
        origin: context.currentCell,
      });
    }

    // Perform MATCH operation
    return matchOperation(
      lookupValueResult.result,
      lookupArray,
      Math.floor(matchTypeResult.result.value) // Floor to handle decimal inputs
    );
  },
};
