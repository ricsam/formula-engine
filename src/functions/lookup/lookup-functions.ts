import type { CellValue, SimpleCellRange } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
} from "../../evaluator/evaluator";
import {
  parseCellAddress,
  parseCellRange,
  offsetRange,
} from "../../core/address";
import { coerceToNumber, isFormulaError, propagateError } from "../index";
import type { ASTNode, RangeNode, ReferenceNode } from "src/parser/ast";

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
export const INDEX: FunctionDefinition = {
  name: "INDEX",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    // Get the array/range
    const array = getArrayFromArg(args[0]);

    // Get row number (1-based)
    let rowNum: number;
    try {
      rowNum = Math.floor(coerceToNumber(args[1]));
    } catch {
      return "#VALUE!";
    }

    // Get column number (1-based, optional)
    let colNum = 1;
    if (args.length >= 3) {
      try {
        colNum = Math.floor(coerceToNumber(args[2]));
      } catch {
        return "#VALUE!";
      }
    }

    // Validate indices
    if (rowNum < 1 || rowNum > array.length) {
      return "#REF!";
    }

    const row = array[rowNum - 1];
    if (!row || colNum < 1 || colNum > row.length) {
      return "#REF!";
    }

    return row[colNum - 1];
  },
};

/**
 * MATCH function - Returns the position of a value in an array
 * MATCH(lookup_value, lookup_array, [match_type])
 * match_type: 1 = less than or equal (default), 0 = exact match, -1 = greater than or equal
 */
export const MATCH: FunctionDefinition = {
  name: "MATCH",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const lookupValue = args[0];
    const lookupArray = getArrayFromArg(args[1]);
    const flatArray = flatten2D(lookupArray);

    // Get match type (default is 1)
    let matchType = 1;
    if (args.length >= 3) {
      try {
        matchType = Math.floor(coerceToNumber(args[2]));
        if (matchType !== -1 && matchType !== 0 && matchType !== 1) {
          return "#VALUE!";
        }
      } catch {
        return "#VALUE!";
      }
    }

    // Exact match
    if (matchType === 0) {
      for (let i = 0; i < flatArray.length; i++) {
        if (compareValues(flatArray[i], lookupValue) === 0) {
          return i + 1; // 1-based index
        }
      }
      return "#N/A";
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
      return lastValidIndex >= 0 ? lastValidIndex + 1 : "#N/A";
    } else {
      // matchType === -1
      // For descending arrays, find the first value >= lookupValue
      for (let i = 0; i < flatArray.length; i++) {
        const cmp = compareValues(flatArray[i], lookupValue);
        if (cmp >= 0) {
          return i + 1;
        }
      }
      return "#N/A";
    }
  },
};

/**
 * VLOOKUP function - Vertical lookup
 * VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
 */
export const VLOOKUP: FunctionDefinition = {
  name: "VLOOKUP",
  minArgs: 3,
  maxArgs: 4,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const lookupValue = args[0];
    const tableArray = getArrayFromArg(args[1]);

    let colIndex: number;
    try {
      colIndex = Math.floor(coerceToNumber(args[2]));
    } catch {
      return "#VALUE!";
    }

    // Validate column index
    if (
      colIndex < 1 ||
      (tableArray.length > 0 &&
        tableArray[0] &&
        colIndex > tableArray[0].length)
    ) {
      return "#REF!";
    }

    // Get range lookup flag (default is TRUE for approximate match)
    let rangeLookup = true;
    if (args.length >= 4) {
      const rl = args[3];
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
        return row[colIndex - 1];
      }
      return "#N/A";
    } else {
      // Exact match
      for (let i = 0; i < tableArray.length; i++) {
        const row = tableArray[i];
        if (!row || row.length === 0) continue;
        if (compareValues(row[0], lookupValue) === 0) {
          return row[colIndex - 1];
        }
      }
      return "#N/A";
    }
  },
};

/**
 * HLOOKUP function - Horizontal lookup
 * HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
 */
export const HLOOKUP: FunctionDefinition = {
  name: "HLOOKUP",
  minArgs: 3,
  maxArgs: 4,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const lookupValue = args[0];
    const tableArray = getArrayFromArg(args[1]);

    let rowIndex: number;
    try {
      rowIndex = Math.floor(coerceToNumber(args[2]));
    } catch {
      return "#VALUE!";
    }

    // Validate row index
    if (rowIndex < 1 || rowIndex > tableArray.length) {
      return "#REF!";
    }

    // Get range lookup flag (default is TRUE)
    let rangeLookup = true;
    if (args.length >= 4) {
      const rl = args[3];
      if (typeof rl === "boolean") {
        rangeLookup = rl;
      } else if (typeof rl === "number") {
        rangeLookup = rl !== 0;
      } else if (typeof rl === "string") {
        rangeLookup = rl.toUpperCase() !== "FALSE";
      }
    }

    if (tableArray.length === 0) return "#N/A";
    const firstRow = tableArray[0];
    if (!firstRow) return "#N/A";

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
        return row[lastValidCol];
      }
      return "#N/A";
    } else {
      // Exact match
      for (let i = 0; i < firstRow.length; i++) {
        if (compareValues(firstRow[i], lookupValue) === 0) {
          const targetRow = tableArray[rowIndex - 1];
          return targetRow ? targetRow[i] : "#N/A";
        }
      }
      return "#N/A";
    }
  },
};

/**
 * XLOOKUP function - Modern lookup function
 * XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
 */
export const XLOOKUP: FunctionDefinition = {
  name: "XLOOKUP",
  minArgs: 3,
  maxArgs: 6,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const lookupValue = args[0];
    const lookupArray = getArrayFromArg(args[1]);
    const returnArray = getArrayFromArg(args[2]);

    // Get if_not_found value (optional)
    const ifNotFound = args.length >= 4 ? args[3] : "#N/A";

    // Get match mode (default is 0 for exact match)
    let matchMode = 0;
    if (args.length >= 5) {
      try {
        matchMode = Math.floor(coerceToNumber(args[4]));
        if (matchMode < -1 || matchMode > 2) {
          return "#VALUE!";
        }
      } catch {
        return "#VALUE!";
      }
    }

    // Get search mode (default is 1 for first to last)
    let searchMode = 1;
    if (args.length >= 6) {
      try {
        searchMode = Math.floor(coerceToNumber(args[5]));
        if (![1, -1, 2, -2].includes(searchMode)) {
          return "#VALUE!";
        }
      } catch {
        return "#VALUE!";
      }
    }

    // Flatten arrays for searching
    const flatLookup = flatten2D(lookupArray);
    const flatReturn = flatten2D(returnArray);

    // Arrays must be same size
    if (flatLookup.length !== flatReturn.length) {
      return "#VALUE!";
    }

    // Determine search direction and method
    const isReverse = searchMode < 0;
    const isBinary = Math.abs(searchMode) === 2;

    if (isBinary && matchMode !== 0) {
      // Binary search only works with exact match
      return "#VALUE!";
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
        return "#VALUE!";
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

    return foundIndex >= 0 ? flatReturn[foundIndex] : ifNotFound;
  },
};

/**
 * ROW function - Returns the row number of the current cell when no argument is provided
 * ROW([reference]) - For now, we support no-arg form only and return current row (1-based)
 */
export const ROW_FN: FunctionDefinition = {
  name: "ROW",
  minArgs: 0,
  maxArgs: 1,
  evaluate: ({ argValues: args, context }): CellValue => {
    // Support only no-arg form reliably
    if (args.length === 0) {
      const row = context.currentCell?.row ?? 0;
      return row + 1;
    }
    // Reference-aware ROW(reference) not yet supported due to lack of address metadata in values
    return "#VALUE!";
  },
};

/**
 * COLUMN function - Returns the column number of the current cell when no argument is provided
 * COLUMN([reference]) - For now, we support no-arg form only and return current column (1-based)
 */
export const COLUMN_FN: FunctionDefinition = {
  name: "COLUMN",
  minArgs: 0,
  maxArgs: 1,
  evaluate: ({ argValues: args, context }): CellValue => {
    if (args.length === 0) {
      const col = context.currentCell?.col ?? 0;
      return col + 1;
    }
    return "#VALUE!";
  },
};

/**
 * ROWS function - Returns the number of rows in an array or reference
 */
export const ROWS_FN: FunctionDefinition = {
  name: "ROWS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const array = getArrayFromArg(args[0]);
    return array.length;
  },
};

/**
 * COLUMNS function - Returns the number of columns in an array or reference
 */
export const COLUMNS_FN: FunctionDefinition = {
  name: "COLUMNS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const array = getArrayFromArg(args[0]);
    const firstRow = array[0] ?? [];
    return (firstRow as CellValue[]).length;
  },
};

/**
 * CHOOSE function - Returns a value from a list of values based on an index
 * CHOOSE(index, value1, value2, ...)
 */
export const CHOOSE_FN: FunctionDefinition = {
  name: "CHOOSE",
  minArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    // First arg is index (1-based)
    const indexRaw = args[0];
    if (typeof indexRaw !== "number") {
      return "#VALUE!";
    }
    const index = Math.floor(indexRaw);
    if (index < 1 || index >= args.length) {
      return "#VALUE!";
    }
    return args[index];
  },
};

/**
 * INDIRECT function - Returns the reference specified by a text string
 * INDIRECT(ref_text)
 * - Supports cell address like "A1" and ranges like "A1:B2"
 * - Returns value for single cell; for ranges returns 2D array of values
 * - Does not support R1C1 or external workbook references in this minimal impl
 */
export const INDIRECT_FN: FunctionDefinition = {
  name: "INDIRECT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args, context }): CellValue => {
    const ref = args[0];
    if (typeof ref !== "string" || ref.trim() === "") return "#REF!";
    const refStr = ref.trim();

    // Prefer single address first to avoid ambiguous single-cell ranges
    const addr = parseCellAddress(refStr, context.currentSheet ?? 0);
    if (addr) {
      return context.getCellValue(addr);
    }

    // Try range next
    const range = parseCellRange(refStr, context.currentSheet ?? 0);
    if (range) {
      const values = context.getRangeValues(range, context.evaluationStack);
      // If 1x1, return scalar when available
      if (values.length === 1 && (values[0]?.length ?? 0) === 1) {
        return values[0]![0]!;
      }
      return values as unknown as CellValue;
    }

    return "#REF!";
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
export const OFFSET_FN: FunctionDefinition = {
  name: "OFFSET",
  minArgs: 3,
  maxArgs: 5,
  evaluate: ({ argValues: args, argNodes, context }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    // Reference can be a raw node (reference/range), text address, or array value
    const [refVal, rowsVal, colsVal, heightVal, widthVal] = args;
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

    if (!baseRange) return "#REF!";

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
      return values[0]![0]!;
    return values as unknown as CellValue;
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
