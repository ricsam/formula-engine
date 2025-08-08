import { FormulaError, type CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import { coerceToNumber, isFormulaError, propagateError, propagateErrorFromEvalResults } from "../utils";

/**
 * Helper function to flatten nested arrays (including 2D arrays from ranges)
 */
function flattenValues(
  values: CellValue | CellValue[] | CellValue[][]
): CellValue[] {
  const result: CellValue[] = [];

  function flatten(val: CellValue | CellValue[] | CellValue[][]): void {
    if (Array.isArray(val)) {
      // Handle both 1D and 2D arrays
      for (const item of val) {
        flatten(item);
      }
    } else {
      result.push(val);
    }
  }
  if (!values) {
    return [values];
  }
  flatten(values);
  return result;
}

/**
 * SUM function - Adds all numbers in the arguments
 * Ignores text, logical values, and empty cells
 */
const SUM: FunctionDefinition = {
  name: "SUM",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    let sum = 0;
    let hasNumbers = false;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      // Skip empty, text, and boolean values (Excel behavior)
      if (
        value === undefined ||
        value === null ||
        typeof value === "string" ||
        typeof value === "boolean"
      ) {
        continue;
      }

      if (typeof value === "number") {
        sum += value;
        hasNumbers = true;
      }
    }

    return { type: "value", value: sum };
  },
};

/**
 * PRODUCT function - Multiplies all numbers in the arguments
 * Ignores text, logical values, and empty cells
 */
const PRODUCT: FunctionDefinition = {
  name: "PRODUCT",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    let product = 1;
    let hasNumbers = false;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      // Skip empty, text, and boolean values (Excel behavior)
      if (
        value === undefined ||
        value === null ||
        typeof value === "string" ||
        typeof value === "boolean"
      ) {
        continue;
      }

      if (typeof value === "number") {
        product *= value;
        hasNumbers = true;
      }
    }

    return { type: "value", value: hasNumbers ? product : 0 };
  },
};

/**
 * COUNT function - Counts cells containing numbers
 * Ignores text, logical values, errors, and empty cells
 */
const COUNT: FunctionDefinition = {
  name: "COUNT",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    // Note: COUNT doesn't propagate errors, it just ignores them
    const flattened = flattenValues(args);
    let count = 0;

    for (const value of flattened) {
      // COUNT ignores errors and counts only numbers (Excel behavior)
      if (typeof value === "number") {
        count++;
      }
    }

    return { type: "value", value: count };
  },
};

/**
 * COUNTBLANK function - Counts empty cells in a range
 */
const COUNTBLANK: FunctionDefinition = {
  name: "COUNTBLANK",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    let count = 0;

    for (const value of flattened) {
      if (value === undefined || value === null || value === "") {
        count++;
      }
    }

    return { type: "value", value: count };
  },
};

/**
 * COUNTIF function - Counts cells that meet a criteria
 */
const COUNTIF: FunctionDefinition = {
  name: "COUNTIF",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    // Only propagate errors from the argument structure, not from data values
    for (const arg of argEvaluatedValues) {
      if (arg.type === "value" && typeof arg.value === "string" && isFormulaError(arg.value)) {
        return { type: "value", value: arg.value };
      }
    }

    const range = argEvaluatedValues[0];
    const criteria = argEvaluatedValues[1];

    if (!range || !criteria) {
      return { type: "value", value: 0 };
    }

    // Flatten the range values
    const flattened = flattenValues(range.value);
    let count = 0;

    for (const value of flattened) {
      if (matchesCriteria(value, flattenValues(criteria.value)[0])) {
        count++;
      }
    }

    return { type: "value", value: count };
  },
};

/**
 * SUMIF function - Sums cells that meet a criteria
 */
const SUMIF: FunctionDefinition = {
  name: "SUMIF",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const range = argEvaluatedValues[0];
    const criteria = argEvaluatedValues[1];
    const sumRange = argEvaluatedValues.length === 3 ? argEvaluatedValues[2] : range;

    if (!sumRange) {
      throw new Error("#VALUE!");
    }

    if (!range || !criteria) {
      return { type: "value", value: 0 };
    }

    // Flatten the range and sum range values
    const rangeFlattened = flattenValues(range.value);
    const sumFlattened = flattenValues(sumRange.value);

    // Ensure both arrays are the same length
    if (rangeFlattened.length !== sumFlattened.length) {
      return { type: "value", value: FormulaError.VALUE };
    }

    let sum = 0;
    let hasNumbers = false;

    const criteriaValue = flattenValues(criteria.value)[0];

    for (let i = 0; i < rangeFlattened.length; i++) {
      if (matchesCriteria(rangeFlattened[i], criteriaValue)) {
        const value = sumFlattened[i];
        if (typeof value === "number") {
          sum += value;
          hasNumbers = true;
        } else if (typeof value === "boolean") {
          sum += value ? 1 : 0;
          hasNumbers = true;
        }
      }
    }

    return { type: "value", value: hasNumbers ? sum : 0 };
  },
};

const AVERAGEIF: FunctionDefinition = {
  name: "AVERAGEIF",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const range = argEvaluatedValues[0];
    const criteria = argEvaluatedValues[1];
    const averageRange = argEvaluatedValues.length === 3 ? argEvaluatedValues[2] : range;

    if (!averageRange) {
      throw new Error("#VALUE!");
    }

    if (!range || !criteria) {
      return { type: "value", value: FormulaError.DIV0 };
    }

    // Flatten the range and average range values
    const rangeFlattened = flattenValues(range.value);
    const averageFlattened = flattenValues(averageRange.value);

    // Ensure both arrays are the same length
    if (rangeFlattened.length !== averageFlattened.length) {
      return { type: "value", value: FormulaError.VALUE };
    }

    let sum = 0;
    let count = 0;

    const criteriaValue = flattenValues(criteria.value)[0];

    for (let i = 0; i < rangeFlattened.length; i++) {
      if (matchesCriteria(rangeFlattened[i], criteriaValue)) {
        const value = averageFlattened[i];
        if (typeof value === "number") {
          sum += value;
          count++;
        } else if (typeof value === "boolean") {
          sum += value ? 1 : 0;
          count++;
        }
      }
    }

    // Return #DIV/0! if no matching values found
    if (count === 0) {
      return { type: "value", value: FormulaError.DIV0 };
    }

    return { type: "value", value: sum / count };
  },
};

/**
 * Helper function to check if a value matches criteria
 * Supports exact matches, comparison operators, and wildcards
 */
function matchesCriteria(value: CellValue, criteria: CellValue): boolean {
  // Handle error values
  if (isFormulaError(value) || isFormulaError(criteria)) {
    return false;
  }

  // Convert criteria to string for parsing
  const criteriaStr = String(criteria);

  // Check for comparison operators
  const comparisonMatch = criteriaStr.match(/^(>=|<=|<>|>|<|=)(.*)$/);

  if (comparisonMatch) {
    const operator = comparisonMatch[1];
    const targetStr = comparisonMatch[2];

    // Try to convert to number for numeric comparisons
    const numValue =
      typeof value === "number"
        ? value
        : typeof value === "boolean"
          ? value
            ? 1
            : 0
          : typeof value === "string" && !isNaN(Number(value))
            ? Number(value)
            : null;

    const numTarget = !isNaN(Number(targetStr)) ? Number(targetStr) : null;

    // If both can be numbers, do numeric comparison
    if (numValue !== null && numTarget !== null) {
      switch (operator) {
        case ">=":
          return numValue >= numTarget;
        case "<=":
          return numValue <= numTarget;
        case "<>":
          return numValue !== numTarget;
        case ">":
          return numValue > numTarget;
        case "<":
          return numValue < numTarget;
        case "=":
          return numValue === numTarget;
      }
    } else {
      // String comparison
      const strValue = String(value);
      const target = targetStr || "";
      switch (operator) {
        case ">=":
          return strValue >= target;
        case "<=":
          return strValue <= target;
        case "<>":
          return strValue !== target;
        case ">":
          return strValue > target;
        case "<":
          return strValue < target;
        case "=":
          return strValue === target;
      }
    }
  }

  // Check for wildcards (* and ?)
  if (
    typeof criteria === "string" &&
    (criteria.includes("*") || criteria.includes("?"))
  ) {
    const pattern = criteria
      .replace(/[.*+?^${}()|[\]\\]/g, "\\$&") // Escape regex special chars
      .replace(/\\\*/g, ".*") // Replace \* with .*
      .replace(/\\\?/g, "."); // Replace \? with .

    const regex = new RegExp(`^${pattern}$`, "i");
    return regex.test(String(value));
  }

  // Handle empty string matching
  if (criteria === "" || criteria === '""') {
    return value === "" || value === undefined || value === null;
  }

  // Exact match (case-insensitive for strings)
  if (typeof value === "string" && typeof criteria === "string") {
    return value.toLowerCase() === criteria.toLowerCase();
  }

  // For other types, use strict equality
  return value === criteria;
}

/**
 * AVERAGE function - Returns the average of numbers
 * Ignores text, logical values, and empty cells
 */
const AVERAGE: FunctionDefinition = {
  name: "AVERAGE",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    let sum = 0;
    let count = 0;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      if (typeof value === "number") {
        sum += value;
        count++;
      }
    }

    if (count === 0) {
      return { type: "value", value: FormulaError.DIV0 };
    }

    return { type: "value", value: sum / count };
  },
};

/**
 * MAX function - Returns the maximum value
 * Ignores text, logical values, and empty cells
 */
const MAX: FunctionDefinition = {
  name: "MAX",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    let max: number | null = null;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      if (typeof value === "number") {
        if (max === null || value > max) {
          max = value;
        }
      }
    }

    return { type: "value", value: max === null ? 0 : max };
  },
};

/**
 * MIN function - Returns the minimum value
 * Ignores text, logical values, and empty cells
 */
const MIN: FunctionDefinition = {
  name: "MIN",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    let min: number | null = null;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      if (typeof value === "number") {
        if (min === null || value < min) {
          min = value;
        }
      }
    }

    return { type: "value", value: min === null ? 0 : min };
  },
};

/**
 * MEDIAN function - Returns the median value
 * Ignores text, logical values, and empty cells
 */
const MEDIAN: FunctionDefinition = {
  name: "MEDIAN",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    const numbers: number[] = [];

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      if (typeof value === "number") {
        numbers.push(value);
      }
    }

    if (numbers.length === 0) {
      return { type: "value", value: FormulaError.NUM };
    }

    // Sort numbers
    numbers.sort((a, b) => a - b);

    const mid = Math.floor(numbers.length / 2);
    if (numbers.length % 2 === 0) {
      return { type: "value", value: (numbers[mid - 1]! + numbers[mid]!) / 2 };
    } else {
      return { type: "value", value: numbers[mid]! };
    }
  },
};

/**
 * STDEV function - Sample standard deviation
 * Ignores text, logical values, and empty cells
 */
const STDEV: FunctionDefinition = {
  name: "STDEV",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    const numbers: number[] = [];

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      if (typeof value === "number") {
        numbers.push(value);
      }
    }

    if (numbers.length < 2) {
      return { type: "value", value: FormulaError.DIV0 };
    }

    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    const squaredDiffs = numbers.map((n) => Math.pow(n - mean, 2));
    const variance =
      squaredDiffs.reduce((sum, d) => sum + d, 0) / (numbers.length - 1);

    return { type: "value", value: Math.sqrt(variance) };
  },
};

/**
 * VAR function - Sample variance
 * Ignores text, logical values, and empty cells
 */
const VAR: FunctionDefinition = {
  name: "VAR",
  minArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const flattened = flattenValues(args);
    const numbers: number[] = [];

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return { type: "value", value: value };
      }

      if (typeof value === "number") {
        numbers.push(value);
      }
    }

    if (numbers.length < 2) {
        return { type: "value", value: FormulaError.DIV0 };
    }

    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    const squaredDiffs = numbers.map((n) => Math.pow(n - mean, 2));
    const variance =
      squaredDiffs.reduce((sum, d) => sum + d, 0) / (numbers.length - 1);

    return { type: "value", value: variance };
  },
};

/**
 * SUMPRODUCT function - Returns the sum of the products of corresponding values
 * SUMPRODUCT(array1, [array2], ...)
 */
const SUMPRODUCT: FunctionDefinition = {
  name: "SUMPRODUCT",
  minArgs: 1,
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    const error = propagateErrorFromEvalResults(argEvaluatedValues);
    if (error) return { type: "value", value: error };

    const arrays: CellValue[][][] = argEvaluatedValues.map((evalResult) => {
      if (evalResult.type === '2d-array') {
        return evalResult.value as CellValue[][];
      }
      // Convert scalar to 1x1 array
      return [[evalResult.value]];
    });

    const rows = arrays[0]?.length ?? 0;
    const cols = rows > 0 ? (arrays[0]![0]?.length ?? 0) : 0;
    for (const arr of arrays) {
      if (arr.length !== rows) return { type: "value", value: FormulaError.VALUE };
      for (let r = 0; r < rows; r++) {
        if ((arr[r]?.length ?? 0) !== cols) return { type: "value", value: FormulaError.VALUE };
      }
    }

    let sum = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let product = 1;
        for (const arr of arrays) {
          const v = arr[r]![c]!;
          if (isFormulaError(v)) return { type: "value", value: v };
          let num = 0;
          if (typeof v === "number") num = v;
          else if (typeof v === "boolean") num = v ? 1 : 0;
          else if (typeof v === "string") {
            const parsed = parseFloat(v);
            num = isNaN(parsed) ? 0 : parsed;
          } else {
            num = 0;
          }
          product *= num;
        }
        sum += product;
      }
    }
    return { type: "value", value: sum };
  },
};

// Export all statistical functions
export const statisticalFunctions: FunctionDefinition[] = [
  SUM,
  PRODUCT,
  COUNT,
  COUNTBLANK,
  COUNTIF,
  SUMIF,
  AVERAGEIF,
  AVERAGE,
  MAX,
  MIN,
  MEDIAN,
  STDEV,
  VAR,
  SUMPRODUCT,
];
