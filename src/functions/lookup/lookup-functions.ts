import type { ASTNode, RangeNode, ReferenceNode } from "src/parser/ast";
import {
  offsetRange,
  parseCellAddress,
  parseCellRange,
} from "../../core/address";
import {
  FormulaError,
  type SimpleCellRange,
  type CellValue,
} from "../../core/types";
import type {
  FunctionDefinition,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import {
  coerceToNumber,
  getArrayFromEvalResult,
  isFormulaError,
  propagateErrorFromEvalResults,
  safeGetScalarValue,
} from "../utils";

/**
 * Helper to get a 2D array from arguments
 * Arrays from the evaluator are already 2D arrays
 */
function getArrayFromArg(arg: CellValue): CellValue[][] {
  if (Array.isArray(arg)) {
    // Arrays from the evaluator are already 2D
    // Check if it's a valid 2D array structure
    if (arg.length > 0 && Array.isArray(arg[0])) {
      return arg as CellValue[][];
    }
    // Safety fallback: if somehow we get a 1D array, wrap it
    return [arg];
  }
  // Single value becomes a 1x1 array
  return [[arg]];
}

/**
 * Helper to flatten a 2D array to 1D
 */
function flatten2D(array: CellValue[][]): CellValue[] {
  const result: CellValue[] = [];
  for (const row of array) {
    for (const cell of row) {
      result.push(cell);
    }
  }
  return result;
}

/**
 * INDEX function - Returns a value from a table or array
 * INDEX(array, row_num, [column_num])
 */
const INDEX: FunctionDefinition = {
  name: "INDEX",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    // Get the array/range (with bounds check)
    if (!argEvaluatedValues[0])
      return { type: "value", value: FormulaError.VALUE };
    if (!argEvaluatedValues[0])
      return { type: "value", value: FormulaError.VALUE };
    const array = getArrayFromEvalResult(argEvaluatedValues[0]);

    // Get row number (1-based)
    let rowNum: number;
    try {
      rowNum = Math.floor(
        coerceToNumber(safeGetScalarValue(argEvaluatedValues, 1, 1))
      );
    } catch {
      return { type: "value", value: FormulaError.VALUE };
    }

    // Get column number (1-based, optional)
    let colNum = 1;
    if (argEvaluatedValues.length >= 3) {
      try {
        colNum = Math.floor(
          coerceToNumber(safeGetScalarValue(argEvaluatedValues, 2, 1))
        );
      } catch {
        return { type: "value", value: FormulaError.VALUE };
      }
    }

    // Validate indices
    if (rowNum < 1 || rowNum > array.length) {
      return { type: "value", value: FormulaError.REF };
    }

    const row = array[rowNum - 1];
    if (!row || colNum < 1 || colNum > row.length) {
      return { type: "value", value: FormulaError.REF };
    }

    return { type: "value", value: row[colNum - 1] };
  },
};

/**
 * MATCH function - Returns the position of a value in an array
 * MATCH(lookup_value, lookup_array, [match_type])
 * match_type: 1 = less than or equal (default), 0 = exact match, -1 = greater than or equal
 */
const MATCH: FunctionDefinition = {
  name: "MATCH",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const lookupValue = safeGetScalarValue(argEvaluatedValues, 0, 0);
    if (!argEvaluatedValues[1])
      return { type: "value", value: FormulaError.VALUE };
    const lookupArray = getArrayFromEvalResult(argEvaluatedValues[1]);
    const flatArray = flatten2D(lookupArray);

    // Get match type (default is 1)
    let matchType = 1;
    if (argEvaluatedValues.length >= 3) {
      try {
        matchType = Math.floor(
          coerceToNumber(safeGetScalarValue(argEvaluatedValues, 2, 1))
        );
        if (matchType !== -1 && matchType !== 0 && matchType !== 1) {
          return { type: "value", value: FormulaError.VALUE };
        }
      } catch {
        return { type: "value", value: FormulaError.VALUE };
      }
    }

    // Exact match
    if (matchType === 0) {
      for (let i = 0; i < flatArray.length; i++) {
        if (compareValues(flatArray[i], lookupValue) === 0) {
          return { type: "value", value: i + 1 }; // 1-based index
        }
      }
      return { type: "value", value: FormulaError.NA };
    }

    // For matchType 1 or -1, array must be sorted
    // matchType 1: find largest value <= lookupValue (array ascending)
    // matchType -1: find smallest value >= lookupValue (array descending)

    if (matchType === 1) {
      let lastValidIndex = -1;
      for (let i = 0; i < flatArray.length; i++) {
        const cmp = compareValues(flatArray[i], lookupValue);
        if (cmp <= 0) {
          lastValidIndex = i;
        } else {
          break; // Array is sorted ascending, so we can stop
        }
      }
      return {
        type: "value",
        value: lastValidIndex >= 0 ? lastValidIndex + 1 : FormulaError.NA,
      };
    } else {
      // matchType === -1
      // For descending arrays, find the first value >= lookupValue
      for (let i = 0; i < flatArray.length; i++) {
        const cmp = compareValues(flatArray[i], lookupValue);
        if (cmp >= 0) {
          return { type: "value", value: i + 1 };
        }
      }
      return { type: "value", value: FormulaError.NA };
    }
  },
};

/**
 * VLOOKUP function - Vertical lookup
 * VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
 */
const VLOOKUP: FunctionDefinition = {
  name: "VLOOKUP",
  minArgs: 3,
  maxArgs: 4,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error)
      return {
        type: "value",
        value: error,
      };

    const lookupValue = safeGetScalarValue(argEvaluatedValues, 0, 0);
    if (!argEvaluatedValues[1])
      return { type: "value", value: FormulaError.VALUE };
    const tableArray = getArrayFromEvalResult(argEvaluatedValues[1]);

    let colIndex: number;
    try {
      colIndex = Math.floor(
        coerceToNumber(safeGetScalarValue(argEvaluatedValues, 2, 1))
      );
    } catch {
      return { type: "value", value: FormulaError.VALUE };
    }

    // Validate column index
    if (
      colIndex < 1 ||
      (tableArray.length > 0 &&
        tableArray[0] &&
        colIndex > tableArray[0].length)
    ) {
      return { type: "value", value: FormulaError.REF };
    }

    // Get range lookup flag (default is TRUE for approximate match)
    let rangeLookup = true;
    if (argEvaluatedValues.length >= 4) {
      const rl = safeGetScalarValue(argEvaluatedValues, 3, false);
      if (typeof rl === "boolean") {
        rangeLookup = rl;
      } else if (typeof rl === "number") {
        rangeLookup = rl !== 0;
      } else if (typeof rl === "string") {
        rangeLookup = rl.toUpperCase() !== "FALSE";
      }
    }

    // Search in first column
    if (rangeLookup) {
      // Approximate match - find largest value <= lookupValue
      let lastValidRow = -1;
      for (let i = 0; i < tableArray.length; i++) {
        const row = tableArray[i];
        if (!row || row.length === 0) continue;
        const cmp = compareValues(row[0], lookupValue);
        if (cmp <= 0) {
          lastValidRow = i;
        } else {
          break; // Assuming sorted data
        }
      }

      if (lastValidRow >= 0 && tableArray[lastValidRow]) {
        const row = tableArray[lastValidRow]!;
        return {
          type: "value",
          value: row[colIndex - 1],
        };
      }
      return { type: "value", value: FormulaError.NA };
    } else {
      // Exact match
      for (let i = 0; i < tableArray.length; i++) {
        const row = tableArray[i];
        if (!row || row.length === 0) continue;
        if (compareValues(row[0], lookupValue) === 0) {
          return {
            type: "value",
            value: row[colIndex - 1],
          };
        }
      }
      return { type: "value", value: FormulaError.NA };
    }
  },
};

/**
 * HLOOKUP function - Horizontal lookup
 * HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
 */
const HLOOKUP: FunctionDefinition = {
  name: "HLOOKUP",
  minArgs: 3,
  maxArgs: 4,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error)
      return {
        type: "value",
        value: error,
      };

    const lookupValue = safeGetScalarValue(argEvaluatedValues, 0, 0);
    if (!argEvaluatedValues[1])
      return { type: "value", value: FormulaError.VALUE };
    const tableArray = getArrayFromEvalResult(argEvaluatedValues[1]);

    let rowIndex: number;
    try {
      rowIndex = Math.floor(
        coerceToNumber(safeGetScalarValue(argEvaluatedValues, 2, 1))
      );
    } catch {
      return { type: "value", value: FormulaError.VALUE };
    }

    // Validate row index
    if (rowIndex < 1 || rowIndex > tableArray.length) {
      return { type: "value", value: FormulaError.REF };
    }

    // Get range lookup flag (default is TRUE)
    let rangeLookup = true;
    if (argEvaluatedValues.length >= 4) {
      const rl = safeGetScalarValue(argEvaluatedValues, 3, false);
      if (typeof rl === "boolean") {
        rangeLookup = rl;
      } else if (typeof rl === "number") {
        rangeLookup = rl !== 0;
      } else if (typeof rl === "string") {
        rangeLookup = rl.toUpperCase() !== "FALSE";
      }
    }

    if (tableArray.length === 0)
      return { type: "value", value: FormulaError.NA };
    const firstRow = tableArray[0];
    if (!firstRow) return { type: "value", value: FormulaError.NA };

    // Search in first row
    if (rangeLookup) {
      // Approximate match
      let lastValidCol = -1;
      for (let i = 0; i < firstRow.length; i++) {
        const cmp = compareValues(firstRow[i], lookupValue);
        if (cmp <= 0) {
          lastValidCol = i;
        } else {
          break;
        }
      }

      if (lastValidCol >= 0 && tableArray[rowIndex - 1]) {
        const row = tableArray[rowIndex - 1]!;
        return {
          type: "value",
          value: row[lastValidCol],
        };
      }
      return { type: "value", value: FormulaError.NA };
    } else {
      // Exact match
      for (let i = 0; i < firstRow.length; i++) {
        if (compareValues(firstRow[i], lookupValue) === 0) {
          const targetRow = tableArray[rowIndex - 1];
          return {
            type: "value",
            value: targetRow ? targetRow[i] : FormulaError.NA,
          };
        }
      }
      return { type: "value", value: FormulaError.NA };
    }
  },
};

/**
 * XLOOKUP function - Modern lookup function
 * XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
 */
const XLOOKUP: FunctionDefinition = {
  name: "XLOOKUP",
  minArgs: 3,
  maxArgs: 6,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error)
      return {
        type: "value",
        value: error,
      };

    const lookupValue = safeGetScalarValue(argEvaluatedValues, 0, 0);
    if (!argEvaluatedValues[1])
      return { type: "value", value: FormulaError.VALUE };
    const lookupArray = getArrayFromEvalResult(argEvaluatedValues[1]);
    if (!argEvaluatedValues[2])
      return { type: "value", value: FormulaError.VALUE };
    const returnArray = getArrayFromEvalResult(argEvaluatedValues[2]);

    // Get if_not_found value (optional)
    const ifNotFound =
      argEvaluatedValues.length >= 4
        ? safeGetScalarValue(argEvaluatedValues, 3, "#N/A")
        : "#N/A";

    // Get match mode (default is 0 for exact match)
    let matchMode = 0;
    if (argEvaluatedValues.length >= 5) {
      try {
        matchMode = Math.floor(
          coerceToNumber(safeGetScalarValue(argEvaluatedValues, 4, 0))
        );
        if (matchMode < -1 || matchMode > 2) {
          return { type: "value", value: FormulaError.VALUE };
        }
      } catch {
        return { type: "value", value: FormulaError.VALUE };
      }
    }

    // Get search mode (default is 1 for first to last)
    let searchMode = 1;
    if (argEvaluatedValues.length >= 6) {
      try {
        searchMode = Math.floor(
          coerceToNumber(safeGetScalarValue(argEvaluatedValues, 5, 1))
        );
        if (![1, -1, 2, -2].includes(searchMode)) {
          return { type: "value", value: FormulaError.VALUE };
        }
      } catch {
        return { type: "value", value: FormulaError.VALUE };
      }
    }

    // Flatten arrays for searching
    const flatLookup = flatten2D(lookupArray);
    const flatReturn = flatten2D(returnArray);

    // Arrays must be same size
    if (flatLookup.length !== flatReturn.length) {
      return { type: "value", value: FormulaError.VALUE };
    }

    // Determine search direction and method
    const isReverse = searchMode < 0;
    const isBinary = Math.abs(searchMode) === 2;

    if (isBinary && matchMode !== 0) {
      // Binary search only works with exact match
      return { type: "value", value: FormulaError.VALUE };
    }

    // Search based on match mode
    let foundIndex = -1;

    if (matchMode === 0) {
      // Exact match
      if (isReverse) {
        for (let i = flatLookup.length - 1; i >= 0; i--) {
          if (compareValues(flatLookup[i], lookupValue) === 0) {
            foundIndex = i;
            break;
          }
        }
      } else {
        for (let i = 0; i < flatLookup.length; i++) {
          if (compareValues(flatLookup[i], lookupValue) === 0) {
            foundIndex = i;
            break;
          }
        }
      }
    } else if (matchMode === -1) {
      // Exact match or next smallest
      let bestIndex = -1;
      for (let i = 0; i < flatLookup.length; i++) {
        const cmp = compareValues(flatLookup[i], lookupValue);
        if (cmp <= 0) {
          bestIndex = i;
          if (cmp === 0) break; // Exact match found
        }
      }
      foundIndex = bestIndex;
    } else if (matchMode === 1) {
      // Exact match or next largest
      for (let i = 0; i < flatLookup.length; i++) {
        const cmp = compareValues(flatLookup[i], lookupValue);
        if (cmp >= 0) {
          foundIndex = i;
          break;
        }
      }
    } else if (matchMode === 2) {
      // Wildcard match (only for strings)
      if (typeof lookupValue !== "string") {
        return { type: "value", value: FormulaError.VALUE };
      }

      const pattern = wildcardToRegex(lookupValue);
      for (let i = 0; i < flatLookup.length; i++) {
        if (
          typeof flatLookup[i] === "string" &&
          pattern.test(flatLookup[i] as string)
        ) {
          foundIndex = i;
          if (!isReverse) break;
        }
      }
    }

    return {
      type: "value",
      value: foundIndex >= 0 ? flatReturn[foundIndex] : ifNotFound,
    };
  },
};

/**
 * ROW function - Returns the row number of the current cell when no argument is provided
 * ROW([reference]) - For now, we support no-arg form only and return current row (1-based)
 */
const ROW_FN: FunctionDefinition = {
  name: "ROW",
  minArgs: 0,
  maxArgs: 1,
  evaluate: ({ argEvaluatedValues, context }): FunctionEvaluationResult => {
    // Support only no-arg form reliably
    if (argEvaluatedValues.length === 0) {
      const row = context.currentCell?.row ?? 0;
      return { type: "value", value: row + 1 };
    }
    // Reference-aware ROW(reference) not yet supported due to lack of address metadata in values
    return { type: "value", value: FormulaError.VALUE };
  },
};

/**
 * COLUMN function - Returns the column number of the current cell when no argument is provided
 * COLUMN([reference]) - For now, we support no-arg form only and return current column (1-based)
 */
const COLUMN_FN: FunctionDefinition = {
  name: "COLUMN",
  minArgs: 0,
  maxArgs: 1,
  evaluate: ({ argEvaluatedValues, context }): FunctionEvaluationResult => {
    if (argEvaluatedValues.length === 0) {
      const col = context.currentCell?.col ?? 0;
      return { type: "value", value: col + 1 };
    }
    return { type: "value", value: FormulaError.VALUE };
  },
};

/**
 * ROWS function - Returns the number of rows in an array or reference
 */
const ROWS_FN: FunctionDefinition = {
  name: "ROWS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (!argEvaluatedValues[0])
      return { type: "value", value: FormulaError.VALUE };
    const array = getArrayFromEvalResult(argEvaluatedValues[0]);
    return { type: "value", value: array.length };
  },
};

/**
 * COLUMNS function - Returns the number of columns in an array or reference
 */
const COLUMNS_FN: FunctionDefinition = {
  name: "COLUMNS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (!argEvaluatedValues[0])
      return { type: "value", value: FormulaError.VALUE };
    const array = getArrayFromEvalResult(argEvaluatedValues[0]);
    const firstRow = array[0] ?? [];
    return { type: "value", value: (firstRow as CellValue[]).length };
  },
};

/**
 * CHOOSE function - Returns a value from a list of values based on an index
 * CHOOSE(index, value1, value2, ...)
 */
const CHOOSE_FN: FunctionDefinition = {
  name: "CHOOSE",
  minArgs: 2,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    // First arg is index (1-based)
    const indexRaw = safeGetScalarValue(argEvaluatedValues, 0, 1);
    if (typeof indexRaw !== "number") {
      return { type: "value", value: FormulaError.VALUE };
    }
    const index = Math.floor(indexRaw);
    if (index < 1 || index >= argEvaluatedValues.length) {
      return { type: "value", value: FormulaError.VALUE };
    }
    return {
      type: "value",
      value: safeGetScalarValue(argEvaluatedValues, index, 0),
    };
  },
};

/**
 * INDIRECT function - Returns the reference specified by a text string
 * INDIRECT(ref_text)
 * - Supports cell address like "A1" and ranges like "A1:B2"
 * - Returns value for single cell; for ranges returns 2D array of values
 * - Does not support R1C1 or external workbook references in this minimal impl
 */
const INDIRECT_FN: FunctionDefinition = {
  name: "INDIRECT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argEvaluatedValues, context }): FunctionEvaluationResult => {
    const ref = safeGetScalarValue(argEvaluatedValues, 0, "");
    if (typeof ref !== "string" || ref.trim() === "")
      return { type: "value", value: FormulaError.REF };
    const refStr = ref.trim();

    // Prefer single address first to avoid ambiguous single-cell ranges
    const addr = parseCellAddress(refStr, context.currentSheet ?? 0);
    if (addr) {
      return { type: "value", value: context.getCellValue(addr) };
    }

    // Try range next
    const range = parseCellRange(refStr, context.currentSheet ?? 0);
    if (range) {
      const values = context.getRangeValues(range, context.evaluationStack);
      // If 1x1, return scalar when available
      if (values.length === 1 && (values[0]?.length ?? 0) === 1) {
        return { type: "value", value: values[0]![0]! };
      }
      return { type: "value", value: values as unknown as CellValue };
    }

    return { type: "value", value: FormulaError.REF };
  },
};

function astNodeIsReference(node: ASTNode): node is ReferenceNode {
  return node.type === "reference";
}

function astNodeIsRange(node: ASTNode): node is RangeNode {
  return node.type === "range";
}

/**
 * OFFSET(reference, rows, cols, [height], [width])
 * Returns a range offset from a starting reference. If height/width provided, returns a resized range.
 */
const OFFSET_FN: FunctionDefinition = {
  name: "OFFSET",
  minArgs: 3,
  maxArgs: 5,
  evaluate: ({
    argEvaluatedValues,
    argNodes,
    context,
  }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    // Reference can be a raw node (reference/range), text address, or array value
    const [refVal, rowsVal, colsVal, heightVal, widthVal] =
      argEvaluatedValues.map((r) => r.value);
    const [refNode] = argNodes;
    const rows = typeof rowsVal === "number" ? Math.floor(rowsVal) : 0;
    const cols = typeof colsVal === "number" ? Math.floor(colsVal) : 0;
    const height =
      typeof heightVal === "number"
        ? Math.max(1, Math.floor(heightVal))
        : undefined;
    const width =
      typeof widthVal === "number"
        ? Math.max(1, Math.floor(widthVal))
        : undefined;

    let baseRange: SimpleCellRange | undefined = undefined;

    if (refNode && astNodeIsReference(refNode)) {
      baseRange = { start: refNode.address, end: refNode.address };
    } else if (refNode && astNodeIsRange(refNode)) {
      baseRange = refNode.range;
    } else if (typeof refVal === "string") {
      const range = parseCellRange(refVal, context.currentSheet ?? 0);
      if (range) baseRange = range;
      else {
        const addr = parseCellAddress(refVal, context.currentSheet ?? 0);
        if (addr) baseRange = { start: addr, end: addr };
      }
    }

    if (!baseRange && Array.isArray(refVal)) {
      // Use current cell as anchor, create range of the array dimensions
      const rowsCount = (refVal as CellValue[][]).length || 1;
      const colsCount = Array.isArray(refVal[0])
        ? (refVal[0] as CellValue[]).length || 1
        : 1;
      const start = context.currentCell ?? {
        sheet: context.currentSheet ?? 0,
        col: 0,
        row: 0,
      };
      baseRange = {
        start,
        end: {
          sheet: start.sheet,
          col: start.col + colsCount - 1,
          row: start.row + rowsCount - 1,
        },
      };
    }

    if (!baseRange) return { type: "value", value: FormulaError.REF };

    // Apply offset
    let target = offsetRange(baseRange, cols, rows);

    // Apply resize if provided
    if (height !== undefined || width !== undefined) {
      const h = height ?? target.end.row - target.start.row + 1;
      const w = width ?? target.end.col - target.start.col + 1;
      target = {
        start: target.start,
        end: {
          sheet: target.start.sheet,
          col: target.start.col + w - 1,
          row: target.start.row + h - 1,
        },
      };
    }

    const values = context.getRangeValues(target, context.evaluationStack);
    if (values.length === 1 && (values[0]?.length ?? 0) === 1)
      return { type: "value", value: values[0]![0]! };
    return { type: "value", value: values as unknown as CellValue };
  },
};

/**
 * Compare two values for sorting/matching
 * Returns: negative if a < b, 0 if equal, positive if a > b
 */
function compareValues(a: CellValue, b: CellValue): number {
  // Handle errors
  if (isFormulaError(a) || isFormulaError(b)) {
    return 0; // Treat errors as equal for comparison
  }

  // Handle undefined/null
  if (a === undefined || a === null) a = 0;
  if (b === undefined || b === null) b = 0;

  // If types are different, convert to numbers
  if (typeof a !== typeof b) {
    try {
      const numA = coerceToNumber(a);
      const numB = coerceToNumber(b);
      return numA - numB;
    } catch {
      // If can't convert, compare as strings
      const strA = String(a);
      const strB = String(b);
      return strA.localeCompare(strB);
    }
  }

  // Same types
  if (typeof a === "number" && typeof b === "number") {
    return a - b;
  }

  if (typeof a === "string" && typeof b === "string") {
    return a.localeCompare(b);
  }

  if (typeof a === "boolean" && typeof b === "boolean") {
    return (a ? 1 : 0) - (b ? 1 : 0);
  }

  return 0;
}

/**
 * Convert wildcard pattern to regex
 * ? matches any single character
 * * matches any sequence of characters
 */
function wildcardToRegex(pattern: string): RegExp {
  const escaped = pattern
    .replace(/[.+^${}()|[\]\\]/g, "\\$&") // Escape regex special chars
    .replace(/\?/g, ".") // ? -> .
    .replace(/\*/g, ".*"); // * -> .*
  return new RegExp(`^${escaped}$`, "i");
}

// Export all lookup functions
export const lookupFunctions: FunctionDefinition[] = [
  INDEX,
  MATCH,
  VLOOKUP,
  HLOOKUP,
  XLOOKUP,
  ROW_FN,
  COLUMN_FN,
  ROWS_FN,
  COLUMNS_FN,
  CHOOSE_FN,
  INDIRECT_FN,
  OFFSET_FN,
];
