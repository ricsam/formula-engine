import {
  FormulaError,
  type CellValue,
  type ErrorEvaluationResult,
  type FunctionDefinition,
  type FunctionEvaluationResult,
  type SingleEvaluationResult,
  type SpilledValuesEvaluationResult,
  type ValueEvaluationResult,
} from "../../../core/types";
import { getSpillAreaShape } from "../../../evaluator/array-shape";
import type { EvaluationContext } from "../../../evaluator/evaluation-context";
import type { FormulaEvaluator } from "../../../evaluator/formula-evaluator";
import { getSuppliedArgument, type FunctionNode } from "../../../parser/ast";

/**
 * XLOOKUP function - Searches for a value and returns a corresponding value from another array
 * XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
 *
 * Parameters:
 * - lookup_value: The value to search for (any type)
 * - lookup_array: The row or column to search
 * - return_array: The row or column to return values from, the same size as lookup_array
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
 *
 * Omitted arguments (a comma with nothing) follow Excel: an omitted
 * if_not_found, match_mode or search_mode takes its default, so
 * `XLOOKUP(x, a, b,, 1)` returns #N/A when nothing matches. An omitted
 * lookup_value finds the first blank cell of lookup_array, as Microsoft
 * documents; a blank cell is empty text in this engine, so a lookup value of
 * empty text finds blank cells too.
 *
 * The arrays are compared by position, so blank cells inside a range (which a
 * range does not report as values) keep their place. Only a blank lookup value
 * matches blank cells.
 */

const BLANK: CellValue = { type: "string", value: "" };

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

/** A lookup or return array read by position along its single row or column. */
type LookupVector = {
  rows: number;
  cols: number;
  /** Cells along the vector; Infinity for an unbounded range such as A:A. */
  length: number;
  /** Reported cells by position. A missing position is a blank cell. */
  cells: Map<number, SingleEvaluationResult>;
  /** The reported positions in ascending order. */
  positions: number[];
};

function valueError(
  message: string,
  context: EvaluationContext
): ErrorEvaluationResult {
  return {
    type: "error",
    err: FormulaError.VALUE,
    errAddress: context.dependencyNode,
    message,
  };
}

/**
 * Reads lookup_array or return_array. Positions come from the cells' places,
 * so the blank cells a range leaves out still separate the cells around them.
 */
function readLookupVector(
  this: FormulaEvaluator,
  result: ValueEvaluationResult | SpilledValuesEvaluationResult,
  name: string,
  context: EvaluationContext
): LookupVector | ErrorEvaluationResult {
  if (result.type === "value") {
    return {
      rows: 1,
      cols: 1,
      length: 1,
      cells: new Map([[0, result]]),
      positions: [0],
    };
  }

  const { rows, cols } = getSpillAreaShape(
    result.spillArea(context.cellAddress)
  );
  if (rows !== 1 && cols !== 1) {
    return valueError(`XLOOKUP ${name} must be a single row or column`, context);
  }
  const vertical = rows !== 1;

  const cellValues = result.evaluateAllCells.call(this, {
    context,
    evaluate: result.evaluate,
    origin: context.cellAddress,
    lookupOrder: vertical ? "col-major" : "row-major",
  });
  if (cellValues.type !== "values") {
    return cellValues;
  }

  const cells = new Map<number, SingleEvaluationResult>();
  for (const { result: cell, relativePos } of cellValues.values) {
    const position = vertical ? relativePos.y : relativePos.x;
    const across = vertical ? relativePos.x : relativePos.y;
    if (position < 0 || across !== 0 || cells.has(position)) continue;
    cells.set(position, cell);
  }

  return {
    rows,
    cols,
    length: vertical ? rows : cols,
    cells,
    positions: Array.from(cells.keys()).sort((a, b) => a - b),
  };
}

function cellAt(vector: LookupVector, position: number): SingleEvaluationResult {
  return vector.cells.get(position) ?? { type: "value", result: BLANK };
}

// Helper function to perform XLOOKUP operation
function xlookupOperation(
  lookupValue: CellValue,
  lookupVector: LookupVector,
  returnVector: LookupVector,
  ifNotFound: SingleEvaluationResult | undefined,
  matchMode: number,
  searchMode: number,
  context: EvaluationContext
): FunctionEvaluationResult {
  // Validate match_mode
  if (![0, -1, 1, 2].includes(matchMode)) {
    return valueError(
      `XLOOKUP match_mode must be -1, 0, 1, or 2, got ${matchMode}`,
      context
    );
  }

  // Validate search_mode
  if (![1, -1, 2, -2].includes(searchMode)) {
    return valueError(
      `XLOOKUP search_mode must be 1, -1, 2, or -2, got ${searchMode}`,
      context
    );
  }

  // Check arrays have same size
  if (
    lookupVector.rows !== returnVector.rows ||
    lookupVector.cols !== returnVector.cols
  ) {
    return valueError(
      "XLOOKUP lookup_array and return_array must have same dimensions",
      context
    );
  }

  // Binary search modes (not yet implemented)
  if (searchMode === 2 || searchMode === -2) {
    return valueError("XLOOKUP binary search mode not yet implemented", context);
  }

  // Blank cells only take part when looking up a blank. Then every position is
  // searched; an unbounded vector is searched one position past its last
  // reported cell, the first of its endless blank tail.
  const lookupBlank = lookupValue.type === "string" && lookupValue.value === "";
  const lastPosition = (vector: LookupVector) =>
    vector.positions[vector.positions.length - 1] ?? -1;
  const searchedCount = lookupBlank
    ? Number.isFinite(lookupVector.length)
      ? lookupVector.length
      : Math.max(lastPosition(lookupVector), lastPosition(returnVector)) + 2
    : lookupVector.positions.length;
  const positionAt = (step: number) => {
    const index = searchMode === -1 ? searchedCount - 1 - step : step;
    return lookupBlank ? index : lookupVector.positions[index]!;
  };

  const returnAt = (position: number): FunctionEvaluationResult => {
    const returnCell = cellAt(returnVector, position);
    return returnCell.type === "value"
      ? { type: "value", result: returnCell.result }
      : returnCell;
  };

  let bestMatchPosition = -1;
  let bestMatchValue: CellValue | null = null;

  // Search through the array
  for (let step = 0; step < searchedCount; step++) {
    const position = positionAt(step);
    const item = cellAt(lookupVector, position);
    if (item.type === "awaiting-evaluation") return item;
    // Error values never match.
    if (item.type !== "value") continue;

    const arrayValue = item.result;

    // Match mode 0: Exact match
    if (matchMode === 0) {
      if (valuesMatch(lookupValue, arrayValue)) {
        return returnAt(position);
      }
    }
    // Match mode 2: Wildcard match
    else if (matchMode === 2) {
      if (
        lookupValue.type === "string" &&
        arrayValue.type === "string" &&
        wildcardMatch(lookupValue.value, arrayValue.value)
      ) {
        return returnAt(position);
      }
    }
    // Match mode -1: Exact or next smaller
    else if (matchMode === -1) {
      const cmp = compareValues(arrayValue, lookupValue);
      if (cmp === 0) {
        return returnAt(position);
      } else if (cmp < 0) {
        // arrayValue < lookupValue
        if (
          bestMatchValue === null ||
          compareValues(arrayValue, bestMatchValue) > 0
        ) {
          bestMatchPosition = position;
          bestMatchValue = arrayValue;
        }
      }
    }
    // Match mode 1: Exact or next larger
    else if (matchMode === 1) {
      const cmp = compareValues(arrayValue, lookupValue);
      if (cmp === 0) {
        return returnAt(position);
      } else if (cmp > 0) {
        // arrayValue > lookupValue
        if (
          bestMatchValue === null ||
          compareValues(arrayValue, bestMatchValue) < 0
        ) {
          bestMatchPosition = position;
          bestMatchValue = arrayValue;
        }
      }
    }
  }

  // Check if we found an approximate match
  if (bestMatchPosition !== -1) {
    return returnAt(bestMatchPosition);
  }

  // No match found - return if_not_found or #N/A
  if (ifNotFound !== undefined) {
    return ifNotFound;
  }

  return {
    type: "error",
    err: FormulaError.NA,
    errAddress: context.dependencyNode,
    message: "XLOOKUP: lookup_value not found",
  };
}

/**
 * Evaluates an optional numeric argument, or returns `defaultValue` when the
 * call omits it.
 */
function evaluateOptionalNumber(
  this: FormulaEvaluator,
  node: FunctionNode,
  index: number,
  name: string,
  defaultValue: number,
  context: EvaluationContext
): number | ErrorEvaluationResult {
  const argument = getSuppliedArgument(node, index);
  if (!argument) {
    return defaultValue;
  }
  const result = this.evaluateNode(argument, context);
  if (result.type === "error" || result.type === "awaiting-evaluation") {
    return result;
  }
  if (result.type === "spilled-values") {
    return valueError("XLOOKUP: Spilled array arguments not yet supported", context);
  }
  if (result.result.type !== "number") {
    return valueError(
      `XLOOKUP ${name} must be number, got ${result.result.type}`,
      context
    );
  }
  return Math.floor(result.result.value);
}

export const XLOOKUP: FunctionDefinition = {
  name: "XLOOKUP",
  evaluate: function (node, context): FunctionEvaluationResult {
    if (node.args.length < 3 || node.args.length > 6) {
      return valueError("XLOOKUP function takes 3 to 6 arguments", context);
    }

    // Evaluate lookup_value. An omitted one evaluates to a blank.
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

    // Evaluate return_array
    const returnArrayResult = this.evaluateNode(node.args[2]!, context);
    if (
      returnArrayResult.type === "error" ||
      returnArrayResult.type === "awaiting-evaluation"
    ) {
      return returnArrayResult;
    }

    // Evaluate if_not_found (optional; without it a failed lookup is #N/A).
    // Its value, even an error, only matters when nothing matches.
    let ifNotFound: SingleEvaluationResult | undefined;
    const ifNotFoundNode = getSuppliedArgument(node, 3);
    if (ifNotFoundNode) {
      const ifNotFoundResult = this.evaluateNode(ifNotFoundNode, context);
      if (ifNotFoundResult.type === "awaiting-evaluation") {
        return ifNotFoundResult;
      }
      if (ifNotFoundResult.type === "spilled-values") {
        return valueError(
          "XLOOKUP: Spilled array arguments not yet supported",
          context
        );
      }
      ifNotFound =
        ifNotFoundResult.type === "value"
          ? { type: "value", result: ifNotFoundResult.result }
          : ifNotFoundResult;
    }

    // Evaluate match_mode (optional, defaults to 0)
    const matchMode = evaluateOptionalNumber.call(
      this,
      node,
      4,
      "match_mode",
      0,
      context
    );
    if (typeof matchMode !== "number") {
      return matchMode;
    }

    // Evaluate search_mode (optional, defaults to 1)
    const searchMode = evaluateOptionalNumber.call(
      this,
      node,
      5,
      "search_mode",
      1,
      context
    );
    if (typeof searchMode !== "number") {
      return searchMode;
    }

    if (lookupValueResult.type === "spilled-values") {
      return valueError(
        "XLOOKUP: Spilled array arguments not yet supported",
        context
      );
    }

    const lookupVector = readLookupVector.call(
      this,
      lookupArrayResult,
      "lookup_array",
      context
    );
    if ("type" in lookupVector) {
      return lookupVector;
    }

    const returnVector = readLookupVector.call(
      this,
      returnArrayResult,
      "return_array",
      context
    );
    if ("type" in returnVector) {
      return returnVector;
    }

    // Perform XLOOKUP operation
    return xlookupOperation(
      lookupValueResult.result,
      lookupVector,
      returnVector,
      ifNotFound,
      matchMode,
      searchMode,
      context
    );
  },
};
