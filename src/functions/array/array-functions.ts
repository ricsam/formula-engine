import { FormulaError, type ASTNode, type CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
  EvaluationResult,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import {
  isFormulaError,
  propagateError,
  propagateErrorFromEvalResults,
  getArrayFromEvalResult,
  getScalarValue,
  safeGetScalarValue,
  assertScalarArg,
  assertArrayArg,
} from "../utils";

// Helper function to coerce values to boolean
function coerceToBoolean(value: CellValue): boolean {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  if (typeof value === "string") {
    const upper = value.toUpperCase();
    if (upper === "TRUE") return true;
    if (upper === "FALSE") return false;
    return value.length > 0;
  }
  if (value === undefined || value === null) {
    return false;
  }
  return false;
}

// Helper to flatten a value into a 1D array
function flattenToArray(result: EvaluationResult): CellValue[] {
  if (result.type === "2d-array") {
    return result.value.flat();
  } else if (result.type === "value") {
    return [result.value];
  } else {
    throw new Error("#VALUE!");
  }
}

// Helper to get dimensions of an array
function getArrayDimensions(result: EvaluationResult): {
  rows: number;
  cols: number;
  isRowRange: boolean;
  isColumnRange: boolean;
} {
  if (result.type === "value") {
    return { rows: 1, cols: 1, isRowRange: false, isColumnRange: false };
  }
  const value = result.value;

  if (Array.isArray(value[0])) {
    // 2D array
    const rows = value.length;
    const cols = Math.max(
      ...value.map((row) => (Array.isArray(row) ? row.length : 1))
    );

    // Row range: 1 row, multiple columns
    const isRowRange = rows === 1 && cols > 1;

    // Column range: multiple rows, 1 column each
    const isColumnRange =
      rows > 1 && value.every((row) => Array.isArray(row) && row.length === 1);

    return { rows, cols, isRowRange, isColumnRange };
  }

  // 1D array - treat as column range
  return {
    rows: value.length,
    cols: 1,
    isRowRange: false,
    isColumnRange: true,
  };
}

// Helper function to filter row ranges (columns within a row)
function filterRowRange(
  sourceArray: CellValue[][],
  conditionArrays: CellValue[][][]
): CellValue | CellValue[][] {
  const sourceRow = sourceArray[0]; // Single row with multiple columns
  if (!sourceRow) {
    throw new Error("#VALUE!");
  }

  const result: CellValue[] = [];

  // Validate all condition arrays are also row ranges with same column count
  for (const condArray of conditionArrays) {
    if (
      !Array.isArray(condArray) ||
      condArray.length !== 1 ||
      !Array.isArray(condArray[0]) ||
      condArray[0].length !== sourceRow.length
    ) {
      throw new Error("#VALUE!");
    }
  }

  // Filter columns based on conditions
  for (let colIndex = 0; colIndex < sourceRow.length; colIndex++) {
    let includeColumn = true;

    // Check all conditions for this column
    for (const condArray of conditionArrays) {
      const conditionRow = condArray[0];
      if (!conditionRow) {
        continue;
      }
      if (!coerceToBoolean(conditionRow[colIndex])) {
        includeColumn = false;
        break;
      }
    }

    if (includeColumn) {
      result.push(sourceRow[colIndex]);
    }
  }

  // If no results, return #N/A
  if (result.length === 0) {
    return FormulaError.NA;
  }

  // Return as column vector (standard FILTER output format)
  return result.map((value) => [value]);
}

// FILTER(sourceArray, ...boolArrays)
const FILTER: FunctionDefinition = {
  name: "FILTER",
  evaluate: ({
    argEvaluatedValues,
  }): ReturnType<FunctionDefinition["evaluate"]> => {
    if (argEvaluatedValues.length < 2) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const firstArg = argEvaluatedValues[0];
    if (!firstArg) {
      throw new Error("#VALUE!");
    }

    if (firstArg.type === "value") {
      // Single value - check if all conditions are true
      for (const val of argEvaluatedValues.slice(1)) {
        const condition = getScalarValue(val);
        if (
          !coerceToBoolean(Array.isArray(condition) ? condition[0] : condition)
        ) {
          return { type: "value", value: FormulaError.NA };
        }
      }
      return {
        type: "2d-array",
        value: [[firstArg.value]],
        dimensions: { rows: 1, cols: 1 },
      }; // Return as 2D array
    }

    // Get dimensions of source array
    const sourceDims = getArrayDimensions(firstArg);

    // Detect if this is a row range (1 row, multiple columns)
    if (sourceDims.rows === 1 && sourceDims.cols > 1) {
      const conditionArrays: CellValue[][][] = [];
      for (const val of argEvaluatedValues.slice(1)) {
        if (val.type === "2d-array") {
          conditionArrays.push(val.value);
        } else {
          // Convert scalar to single-cell 2D array
          conditionArrays.push([[val.value]]);
        }
      }
      return {
        type: "2d-array",
        value: filterRowRange(firstArg.value, conditionArrays) as CellValue[][],
        dimensions: { rows: 1, cols: firstArg.value[0]?.length ?? 0 },
      };
    }

    // Validate all condition arrays have compatible dimensions
    for (const condArray of argEvaluatedValues.slice(1)) {
      const condDims = getArrayDimensions(condArray);

      // Condition must have same number of rows
      if (condDims.rows !== sourceDims.rows) {
        throw new Error("#VALUE!");
      }

      // For 2D source arrays, condition can be either 1D (applied row-wise) or 2D (exact match)
      if (condDims.cols !== 1 && condDims.cols !== sourceDims.cols) {
        throw new Error("#VALUE!");
      }
    }

    // Filter the array
    const result: CellValue[][] = [];

    // 2D array filtering
    for (let r = 0; r < sourceDims.rows; r++) {
      let includeRow = true;

      // Check all conditions for this row
      for (const condArray of argEvaluatedValues.slice(1)) {
        if (condArray.type === "value") {
          // 1D condition array - check the single value for this row
          const condValue = Array.isArray(condArray)
            ? Array.isArray(condArray[r])
              ? condArray[r][0]
              : condArray[r]
            : condArray;
          if (!coerceToBoolean(condValue)) {
            includeRow = false;
            break;
          }
        } else {
          const condDims = getArrayDimensions(condArray);
          // 2D condition array - check if any value in this row is true
          const condRow = condArray.value[r];
          let rowHasTrue = false;
          if (!condRow) {
            continue;
          }
          for (let c = 0; c < condDims.cols; c++) {
            const condValue = condRow[c];
            if (coerceToBoolean(condValue)) {
              rowHasTrue = true;
              break;
            }
          }
          if (!rowHasTrue) {
            includeRow = false;
            break;
          }
        }
      }

      const val = firstArg.value[r];
      if (includeRow && val) {
        result.push(val);
      }
    }

    // If no results, return #N/A
    if (result.length === 0) {
      return { type: "value", value: FormulaError.NA };
    }

    return {
      type: "2d-array",
      value: result,
      dimensions: { rows: result.length, cols: result[0]?.length ?? 1 },
    };
  },
};

// SORT(array, [sort_index], [sort_order], [by_col])
const SORT: FunctionDefinition = {
  name: "SORT",
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (argEvaluatedValues.length < 1) {
      throw new Error("#VALUE!");
    }

    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const firstArg = argEvaluatedValues[0];
    if (!firstArg) {
      throw new Error("#VALUE!");
    }

    const source = getArrayFromEvalResult(firstArg);
    const sortIndexRaw = safeGetScalarValue(argEvaluatedValues, 1, 1);
    const sortOrderRaw = safeGetScalarValue(argEvaluatedValues, 2, 1); // 1 asc, -1 desc
    const byColRaw = safeGetScalarValue(argEvaluatedValues, 3, false);

    const byCol =
      typeof byColRaw === "boolean"
        ? byColRaw
        : typeof byColRaw === "number"
          ? byColRaw !== 0
          : false;
    const sortIndex =
      typeof sortIndexRaw === "number"
        ? Math.max(1, Math.floor(sortIndexRaw))
        : 1;
    const sortOrder =
      typeof sortOrderRaw === "number" && sortOrderRaw < 0 ? -1 : 1;

    const rows = source.length;
    const cols =
      rows > 0 ? (Array.isArray(source[0]) ? source[0].length : 1) : 0;

    // Edge cases
    if (rows === 0 || cols === 0)
      return {
        type: "2d-array",
        value: [[]],
        dimensions: { rows: 1, cols: 1 },
      };

    // Clone data for sorting
    const clone2D = source.map((r) =>
      Array.isArray(r) ? [...r] : [r]
    ) as CellValue[][];

    // Comparator
    const cmp = (a: CellValue, b: CellValue): number => {
      // Treat undefined as smallest
      if (a === undefined && b === undefined) return 0;
      if (a === undefined) return -1 * sortOrder;
      if (b === undefined) return 1 * sortOrder;
      if (typeof a === "number" && typeof b === "number")
        return (a - b) * sortOrder;
      const sa = typeof a === "string" ? a : String(a);
      const sb = typeof b === "string" ? b : String(b);
      return sa.localeCompare(sb) * sortOrder;
    };

    if (byCol) {
      // Sort columns by row sortIndex
      const idx = Math.min(sortIndex - 1, rows - 1);
      // Transpose sort
      const columns: CellValue[][] = [];
      for (let c = 0; c < cols; c++) {
        const col: CellValue[] = [];
        for (let r = 0; r < rows; r++) col.push(clone2D[r]?.[c]);
        columns.push(col);
      }
      columns.sort((c1, c2) => cmp(c1[idx], c2[idx]));
      // Reconstruct rows
      const sorted: CellValue[][] = Array.from({ length: rows }, () => []);
      for (let c = 0; c < columns.length; c++) {
        for (let r = 0; r < rows; r++) sorted[r]?.push(columns[c]?.[r]);
      }
      return {
        type: "2d-array",
        value: sorted,
        dimensions: { rows: sorted.length, cols: sorted[0]?.length ?? 1 },
      };
    }

    // Sort rows by column sortIndex
    const idx = Math.min(sortIndex - 1, Math.max(0, cols - 1));
    clone2D.sort((r1, r2) => cmp(r1[idx], r2[idx]));
    return {
      type: "2d-array",
      value: clone2D,
      dimensions: { rows: clone2D.length, cols: clone2D[0]?.length ?? 1 },
    };
  },
};

// UNIQUE(array, [by_col], [exactly_once])
const UNIQUE: FunctionDefinition = {
  name: "UNIQUE",
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (argEvaluatedValues.length < 1) throw new Error("#VALUE!");
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const firstArg = argEvaluatedValues[0];
    if (!firstArg) {
      throw new Error("#VALUE!");
    }

    const source = getArrayFromEvalResult(firstArg);
    const byColRaw = safeGetScalarValue(argEvaluatedValues, 1, false);
    const exactlyOnceRaw = safeGetScalarValue(argEvaluatedValues, 2, false);
    const byCol =
      typeof byColRaw === "boolean"
        ? byColRaw
        : typeof byColRaw === "number"
          ? byColRaw !== 0
          : false;
    const exactlyOnce =
      typeof exactlyOnceRaw === "boolean"
        ? exactlyOnceRaw
        : typeof exactlyOnceRaw === "number"
          ? exactlyOnceRaw !== 0
          : false;

    const rows = source.length;
    const cols = rows > 0 ? (source[0] as CellValue[]).length : 0;

    if (byCol) {
      // Unique columns
      const seen = new Map<string, number>();
      const columns: CellValue[][] = [];
      for (let c = 0; c < cols; c++) {
        const col: CellValue[] = [];
        for (let r = 0; r < rows; r++) col.push(source[r]?.[c]);
        const key = JSON.stringify(col);
        seen.set(key, (seen.get(key) ?? 0) + 1);
        columns.push(col);
      }
      const uniques: CellValue[][] = [];
      for (let c = 0; c < columns.length; c++) {
        const key = JSON.stringify(columns[c]);
        const count = seen.get(key) ?? 0;
        if (
          (exactlyOnce && count === 1) ||
          (!exactlyOnce &&
            count >= 1 &&
            uniques.findIndex((u) => JSON.stringify(u) === key) === -1)
        ) {
          uniques.push(columns[c] ?? []);
        }
      }
      // Reconstruct rows
      const result: CellValue[][] = Array.from({ length: rows }, () => []);
      for (const col of uniques) {
        for (let r = 0; r < rows; r++) result[r]?.push(col[r]);
      }
      return {
        type: "2d-array",
        value: result,
        dimensions: { rows: result.length, cols: result[0]?.length ?? 1 },
      };
    } else {
      // Unique rows
      const seen = new Map<string, number>();
      for (const row of source) {
        const key = JSON.stringify(row);
        seen.set(key, (seen.get(key) ?? 0) + 1);
      }
      const result: CellValue[][] = [];
      for (const row of source) {
        const key = JSON.stringify(row);
        const count = seen.get(key) ?? 0;
        if (
          (exactlyOnce && count === 1) ||
          (!exactlyOnce &&
            count >= 1 &&
            result.findIndex((r) => JSON.stringify(r) === key) === -1)
        ) {
          result.push(row);
        }
      }
      return {
        type: "2d-array",
        value: result,
        dimensions: { rows: result.length, cols: result[0]?.length ?? 1 },
      };
    }
  },
};

// SEQUENCE(rows, [columns], [start], [step])
const SEQUENCE: FunctionDefinition = {
  name: "SEQUENCE",
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (argEvaluatedValues.length < 1) throw new Error("#VALUE!");
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const rowsValue = safeGetScalarValue(argEvaluatedValues, 0, 0);
    const rows =
      typeof rowsValue === "number" ? Math.max(0, Math.floor(rowsValue)) : 0;

    const colsValue = safeGetScalarValue(argEvaluatedValues, 1, 1);
    const cols =
      typeof colsValue === "number" ? Math.max(0, Math.floor(colsValue)) : 1;

    const startValue = safeGetScalarValue(argEvaluatedValues, 2, 1);
    const start = typeof startValue === "number" ? startValue : 1;

    const stepValue = safeGetScalarValue(argEvaluatedValues, 3, 1);
    const step = typeof stepValue === "number" ? stepValue : 1;

    const result: CellValue[][] = [];
    let current = start;
    if (rows === 0 || cols === 0)
      return {
        type: "2d-array",
        value: [[]],
        dimensions: { rows: 1, cols: 1 },
      };
    for (let r = 0; r < rows; r++) {
      const row: CellValue[] = [];
      for (let c = 0; c < cols; c++) {
        row.push(current);
        current += step;
      }
      result.push(row);
    }
    return {
      type: "2d-array",
      value: result,
      dimensions: { rows: result.length, cols: result[0]?.length ?? 1 },
    };
  },
};

// ARRAY_CONSTRAIN(array, height, width)
const ARRAY_CONSTRAIN: FunctionDefinition = {
  name: "ARRAY_CONSTRAIN",
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (argEvaluatedValues.length !== 3) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const firstArg = argEvaluatedValues[0];
    if (!firstArg) {
      throw new Error("#VALUE!");
    }

    const array = getArrayFromEvalResult(firstArg);
    const height = safeGetScalarValue(argEvaluatedValues, 1, 1);
    const width = safeGetScalarValue(argEvaluatedValues, 2, 1);

    // Validate height and width are positive integers
    if (typeof height !== "number" || height < 1 || !Number.isInteger(height)) {
      throw new Error("#VALUE!");
    }
    if (typeof width !== "number" || width < 1 || !Number.isInteger(width)) {
      throw new Error("#VALUE!");
    }

    // Handle non-array input
    if (firstArg.type === "value") {
      return {
        type: "2d-array",
        value: [[firstArg.value]],
        dimensions: { rows: 1, cols: 1 },
      }; // Return as 1x1 array
    }

    // Get dimensions
    const is2D = Array.isArray(array[0]);
    const result: CellValue[][] = [];

    if (is2D) {
      // Constrain 2D array
      for (let r = 0; r < Math.min(height, array.length); r++) {
        const row: CellValue[] = [];
        const sourceRow = array[r];

        if (Array.isArray(sourceRow)) {
          for (let c = 0; c < Math.min(width, sourceRow.length); c++) {
            row.push(sourceRow[c]);
          }
        } else {
          row.push(sourceRow);
        }

        // Pad with undefined if needed
        while (row.length < width) {
          row.push(undefined);
        }

        result.push(row);
      }
    }

    // Pad with empty rows if needed
    while (result.length < height) {
      const emptyRow: CellValue[] = [];
      for (let c = 0; c < width; c++) {
        emptyRow.push(undefined);
      }
      result.push(emptyRow);
    }

    return {
      type: "2d-array",
      value: result,
      dimensions: { rows: result.length, cols: result[0]?.length ?? 1 },
    };
  },
};

// Export all array functions
export const arrayFunctions: FunctionDefinition[] = [
  FILTER,
  SORT,
  UNIQUE,
  SEQUENCE,
  ARRAY_CONSTRAIN,
];
