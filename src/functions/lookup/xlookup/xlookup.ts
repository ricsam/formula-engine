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
  type EvaluateAllCellsResult,
} from "src/core/types";
import type { EvaluationContext } from "src/evaluator/evaluation-context";
import type { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * XLOOKUP function - Searches for a value and returns a corresponding value from another array
 * XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
 *
 * Parameters:
 * - lookup_value: The value to search for (any type)
 * - lookup_array: The array or range to search
 * - return_array: The array or range to return values from
 * - if_not_found: Optional value to return when no match is found (default: #N/A)
 * - match_mode: Optional integer specifying match type (default: 0)
 *   - 0: Exact match; if none, return #N/A
 *   - -1: Exact match; if none, return next smaller item
 *   - 1: Exact match; if none, return next larger item
 *   - 2: Wildcard match (*, ?, ~)
 * - search_mode: Optional integer specifying search direction (default: 1)
 *   - 1: Search from first item (forward)
 *   - -1: Search from last item (reverse)
 *   - 2: Binary search (ascending order required)
 *   - -2: Binary search (descending order required)
 */

// Helper to compare values for XLOOKUP
function compareValues(a: CellValue, b: CellValue): number {
  // Handle different type comparisons
  if (a.type !== b.type) {
    // Numbers < Strings < Booleans in Excel comparison order
    const typeOrder = { number: 0, string: 1, boolean: 2, infinity: 0 };
    return (typeOrder[a.type] || 3) - (typeOrder[b.type] || 3);
  }

  // Same type comparison
  if (a.type === "number" && b.type === "number") {
    return a.value - b.value;
  }
  if (a.type === "string" && b.type === "string") {
    return a.value.localeCompare(b.value);
  }
  if (a.type === "boolean" && b.type === "boolean") {
    return (a.value ? 1 : 0) - (b.value ? 1 : 0);
  }

  return 0;
}

// Helper to check if values match (exact match)
function valuesMatch(a: CellValue, b: CellValue): boolean {
  if (a.type !== b.type) return false;
  if (a.type === "number" && b.type === "number") {
    return a.value === b.value;
  }
  if (a.type === "string" && b.type === "string") {
    return a.value === b.value;
  }
  if (a.type === "boolean" && b.type === "boolean") {
    return a.value === b.value;
  }
  return false;
}

// Helper to check wildcard match
function wildcardMatch(pattern: string, text: string): boolean {
  // Convert wildcard pattern to regex
  // ~ is escape character, * matches any sequence, ? matches single char
  let regex = "";
  let i = 0;
  while (i < pattern.length) {
    const char = pattern[i];
    if (!char) break;
    
    if (char === "~" && i + 1 < pattern.length) {
      // Escape next character
      const nextChar = pattern[i + 1];
      if (nextChar) {
        regex += nextChar.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      }
      i += 2;
    } else if (char === "*") {
      regex += ".*";
      i++;
    } else if (char === "?") {
      regex += ".";
      i++;
    } else {
      regex += char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      i++;
    }
  }

  try {
    const re = new RegExp(`^${regex}$`, "i"); // Case-insensitive
    return re.test(text);
  } catch {
    return false;
  }
}

// Helper function to perform XLOOKUP operation
function xlookupOperation(
  this: FormulaEvaluator,
  lookupValue: CellValue,
  lookupArray: EvaluateAllCellsResult[],
  returnArray: EvaluateAllCellsResult[],
  ifNotFound: CellValue | null,
  hasIfNotFound: boolean,
  matchMode: number,
  searchMode: number,
  context: EvaluationContext
): FunctionEvaluationResult {
  // Validate match_mode
  if (![0, -1, 1, 2].includes(matchMode)) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      errAddress: context.originCell.cellAddress,
      message: `XLOOKUP match_mode must be -1, 0, 1, or 2, got ${matchMode}`,
    };
  }

  // Validate search_mode
  if (![1, -1, 2, -2].includes(searchMode)) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      errAddress: context.originCell.cellAddress,
      message: `XLOOKUP search_mode must be 1, -1, 2, or -2, got ${searchMode}`,
    };
  }

  // Check arrays have same length
  if (lookupArray.length !== returnArray.length) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      errAddress: context.originCell.cellAddress,
      message: "XLOOKUP lookup_array and return_array must have same dimensions",
    };
  }

  if (lookupArray.length === 0) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      errAddress: context.originCell.cellAddress,
      message: "XLOOKUP lookup_array cannot be empty",
    };
  }

  // Binary search modes (not yet implemented)
  if (searchMode === 2 || searchMode === -2) {
    return {
      type: "error",
      err: FormulaError.VALUE,
      errAddress: context.originCell.cellAddress,
      message: "XLOOKUP binary search mode not yet implemented",
    };
  }

  // Determine search order
  const searchArray =
    searchMode === -1 ? [...lookupArray].reverse() : lookupArray;
  const searchReturnArray =
    searchMode === -1 ? [...returnArray].reverse() : returnArray;

  let bestMatchIndex = -1;
  let bestMatchValue: CellValue | null = null;

  // Search through the array
  for (let i = 0; i < searchArray.length; i++) {
    const item = searchArray[i];
    if (!item || item.result.type !== "value") continue;

    const arrayValue = item.result.result;

    // Match mode 0: Exact match
    if (matchMode === 0) {
      if (valuesMatch(lookupValue, arrayValue)) {
        const returnItem = searchReturnArray[i];
        if (!returnItem) continue;
        
        if (returnItem.result.type === "value") {
          return {
            type: "value",
            result: returnItem.result.result,
          };
        } else if (returnItem.result.type === "error") {
          return returnItem.result;
        }
      }
    }
    // Match mode 2: Wildcard match
    else if (matchMode === 2) {
      if (lookupValue.type === "string" && arrayValue.type === "string") {
        if (wildcardMatch(lookupValue.value, arrayValue.value)) {
          const returnItem = searchReturnArray[i];
          if (!returnItem) continue;
          
          if (returnItem.result.type === "value") {
            return {
              type: "value",
              result: returnItem.result.result,
            };
          } else if (returnItem.result.type === "error") {
            return returnItem.result;
          }
        }
      }
    }
    // Match mode -1: Exact or next smaller
    else if (matchMode === -1) {
      const cmp = compareValues(arrayValue, lookupValue);
      if (cmp === 0) {
        // Exact match found
        const returnItem = searchReturnArray[i];
        if (!returnItem) continue;
        
        if (returnItem.result.type === "value") {
          return {
            type: "value",
            result: returnItem.result.result,
          };
        } else if (returnItem.result.type === "error") {
          return returnItem.result;
        }
      } else if (cmp < 0) {
        // arrayValue < lookupValue
        if (bestMatchValue === null || compareValues(arrayValue, bestMatchValue) > 0) {
          bestMatchIndex = i;
          bestMatchValue = arrayValue;
        }
      }
    }
    // Match mode 1: Exact or next larger
    else if (matchMode === 1) {
      const cmp = compareValues(arrayValue, lookupValue);
      if (cmp === 0) {
        // Exact match found
        const returnItem = searchReturnArray[i];
        if (!returnItem) continue;
        
        if (returnItem.result.type === "value") {
          return {
            type: "value",
            result: returnItem.result.result,
          };
        } else if (returnItem.result.type === "error") {
          return returnItem.result;
        }
      } else if (cmp > 0) {
        // arrayValue > lookupValue
        if (bestMatchValue === null || compareValues(arrayValue, bestMatchValue) < 0) {
          bestMatchIndex = i;
          bestMatchValue = arrayValue;
        }
      }
    }
  }

  // Check if we found an approximate match
  if (bestMatchIndex !== -1) {
    const returnItem = searchReturnArray[bestMatchIndex];
    if (returnItem && returnItem.result.type === "value") {
      return {
        type: "value",
        result: returnItem.result.result,
      };
    } else if (returnItem && returnItem.result.type === "error") {
      return returnItem.result;
    }
  }

  // No match found - return if_not_found or #N/A
  if (hasIfNotFound && ifNotFound !== null) {
    return {
      type: "value",
      result: ifNotFound,
    };
  }

  return {
    type: "error",
    err: FormulaError.NA,
    errAddress: context.originCell.cellAddress,
    message: "XLOOKUP: lookup_value not found",
  };
}

export const XLOOKUP: FunctionDefinition = {
  name: "XLOOKUP",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 3 || node.args.length > 6) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP function takes 3 to 6 arguments",
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

    // Evaluate return_array
    const returnArrayResult = this.evaluateNode(node.args[2]!, context);
    if (returnArrayResult.type === "error") {
      return returnArrayResult;
    }

    // Evaluate if_not_found (optional)
    let ifNotFoundResult: FunctionEvaluationResult | null = null;
    if (node.args[3]) {
      ifNotFoundResult = this.evaluateNode(node.args[3], context);
      if (ifNotFoundResult.type === "error") {
        return ifNotFoundResult;
      }
    }

    // Evaluate match_mode (optional, defaults to 0)
    let matchModeResult: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 0 },
    };
    if (node.args[4]) {
      matchModeResult = this.evaluateNode(node.args[4], context);
      if (matchModeResult.type === "error") {
        return matchModeResult;
      }
    }

    // Evaluate search_mode (optional, defaults to 1)
    let searchModeResult: FunctionEvaluationResult = {
      type: "value",
      result: { type: "number", value: 1 },
    };
    if (node.args[5]) {
      searchModeResult = this.evaluateNode(node.args[5], context);
      if (searchModeResult.type === "error") {
        return searchModeResult;
      }
    }

    // Handle spilled arrays for simple arguments
    if (
      lookupValueResult.type === "spilled-values" ||
      (ifNotFoundResult && ifNotFoundResult.type === "spilled-values") ||
      matchModeResult.type === "spilled-values" ||
      searchModeResult.type === "spilled-values"
    ) {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP: Spilled array arguments not yet supported",
      };
    }

    // Validate result types
    if (lookupValueResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP: Invalid lookup_value result type",
      };
    }

    if (matchModeResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP: Invalid match_mode result type",
      };
    }

    if (searchModeResult.type !== "value") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP: Invalid search_mode result type",
      };
    }

    // Type check match_mode
    if (matchModeResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: `XLOOKUP match_mode must be number, got ${matchModeResult.result.type}`,
      };
    }

    // Type check search_mode
    if (searchModeResult.result.type !== "number") {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: `XLOOKUP search_mode must be number, got ${searchModeResult.result.type}`,
      };
    }

    // Extract if_not_found value
    let ifNotFound: CellValue | null = null;
    const hasIfNotFound = !!ifNotFoundResult;
    if (ifNotFoundResult && ifNotFoundResult.type === "value") {
      ifNotFound = ifNotFoundResult.result;
    }

    // Extract match_mode and search_mode as integers
    const matchMode = Math.floor(matchModeResult.result.value);
    const searchMode = Math.floor(searchModeResult.result.value);

    // Extract lookup_array values
    let lookupArray: EvaluateAllCellsResult[] = [];
    if (lookupArrayResult.type === "value") {
      lookupArray = [
        { result: lookupArrayResult, relativePos: { x: 0, y: 0 } },
      ];
    } else if (lookupArrayResult.type === "spilled-values") {
      lookupArray = Array.from(
        lookupArrayResult.evaluateAllCells.call(this, {
          context,
          evaluate: lookupArrayResult.evaluate,
          origin: context.originCell.cellAddress,
          lookupOrder: "col-major",
        })
      );
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP: Invalid lookup_array type",
      };
    }

    // Extract return_array values
    let returnArray: EvaluateAllCellsResult[] = [];
    if (returnArrayResult.type === "value") {
      returnArray = [
        { result: returnArrayResult, relativePos: { x: 0, y: 0 } },
      ];
    } else if (returnArrayResult.type === "spilled-values") {
      returnArray = Array.from(
        returnArrayResult.evaluateAllCells.call(this, {
          context,
          evaluate: returnArrayResult.evaluate,
          origin: context.originCell.cellAddress,
          lookupOrder: "col-major",
        })
      );
    } else {
      return {
        type: "error",
        err: FormulaError.VALUE,
        errAddress: context.originCell.cellAddress,
        message: "XLOOKUP: Invalid return_array type",
      };
    }

    // Perform XLOOKUP operation
    return xlookupOperation.call(
      this,
      lookupValueResult.result,
      lookupArray,
      returnArray,
      ifNotFound,
      hasIfNotFound,
      matchMode,
      searchMode,
      context
    );
  },
};
