import type { CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
} from "../../evaluator/evaluator";
import { coerceToNumber, isFormulaError, propagateError } from "../index";

/**
 * Helper function to flatten nested arrays (including 2D arrays from ranges)
 */
function flattenValues(values: CellValue[]): CellValue[] {
  const result: CellValue[] = [];

  function flatten(val: CellValue): void {
    if (Array.isArray(val)) {
      // Handle both 1D and 2D arrays
      for (const item of val) {
        flatten(item);
      }
    } else {
      result.push(val);
    }
  }

  values.forEach(flatten);
  return result;
}

/**
 * SUM function - Adds all numbers in the arguments
 * Ignores text, logical values, and empty cells
 */
export const SUM: FunctionDefinition = {
  name: "SUM",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    let sum = 0;
    let hasNumbers = false;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
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

    return sum;
  },
};

/**
 * PRODUCT function - Multiplies all numbers in the arguments
 * Ignores text, logical values, and empty cells
 */
export const PRODUCT: FunctionDefinition = {
  name: "PRODUCT",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    let product = 1;
    let hasNumbers = false;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
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

    return hasNumbers ? product : 0;
  },
};

/**
 * COUNT function - Counts cells containing numbers
 * Ignores text, logical values, errors, and empty cells
 */
export const COUNT: FunctionDefinition = {
  name: "COUNT",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    // Note: COUNT doesn't propagate errors, it just ignores them
    const flattened = flattenValues(args);
    let count = 0;

    for (const value of flattened) {
      // COUNT ignores errors and counts only numbers (Excel behavior)
      if (typeof value === "number") {
        count++;
      }
    }

    return count;
  },
};

/**
 * COUNTBLANK function - Counts empty cells in a range
 */
export const COUNTBLANK: FunctionDefinition = {
  name: "COUNTBLANK",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    let count = 0;

    for (const value of flattened) {
      if (value === undefined || value === null || value === "") {
        count++;
      }
    }

    return count;
  },
};

/**
 * COUNTIF function - Counts cells that meet a criteria
 */
export const COUNTIF: FunctionDefinition = {
  name: "COUNTIF",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const range = args[0];
    const criteria = args[1];

    // Flatten the range values
    const flattened = flattenValues([range]);
    let count = 0;

    for (const value of flattened) {
      if (matchesCriteria(value, criteria)) {
        count++;
      }
    }

    return count;
  },
};

/**
 * SUMIF function - Sums cells that meet a criteria
 */
export const SUMIF: FunctionDefinition = {
  name: "SUMIF",
  minArgs: 2,
  maxArgs: 3,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const range = args[0];
    const criteria = args[1];
    const sumRange = args.length === 3 ? args[2] : range;

    // Flatten the range and sum range values
    const rangeFlattened = flattenValues([range]);
    const sumFlattened = flattenValues([sumRange]);

    // Ensure both arrays are the same length
    if (rangeFlattened.length !== sumFlattened.length) {
      return "#VALUE!";
    }

    let sum = 0;
    let hasNumbers = false;

    for (let i = 0; i < rangeFlattened.length; i++) {
      if (matchesCriteria(rangeFlattened[i], criteria)) {
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

    return hasNumbers ? sum : 0;
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
export const AVERAGE: FunctionDefinition = {
  name: "AVERAGE",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    let sum = 0;
    let count = 0;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }

      if (typeof value === "number") {
        sum += value;
        count++;
      }
    }

    if (count === 0) {
      return "#DIV/0!";
    }

    return sum / count;
  },
};

/**
 * MAX function - Returns the maximum value
 * Ignores text, logical values, and empty cells
 */
export const MAX: FunctionDefinition = {
  name: "MAX",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    let max: number | null = null;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }

      if (typeof value === "number") {
        if (max === null || value > max) {
          max = value;
        }
      }
    }

    return max === null ? 0 : max;
  },
};

/**
 * MIN function - Returns the minimum value
 * Ignores text, logical values, and empty cells
 */
export const MIN: FunctionDefinition = {
  name: "MIN",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    let min: number | null = null;

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }

      if (typeof value === "number") {
        if (min === null || value < min) {
          min = value;
        }
      }
    }

    return min === null ? 0 : min;
  },
};

/**
 * MEDIAN function - Returns the median value
 * Ignores text, logical values, and empty cells
 */
export const MEDIAN: FunctionDefinition = {
  name: "MEDIAN",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    const numbers: number[] = [];

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }

      if (typeof value === "number") {
        numbers.push(value);
      }
    }

    if (numbers.length === 0) {
      return "#NUM!";
    }

    // Sort numbers
    numbers.sort((a, b) => a - b);

    const mid = Math.floor(numbers.length / 2);
    if (numbers.length % 2 === 0) {
      return (numbers[mid - 1]! + numbers[mid]!) / 2;
    } else {
      return numbers[mid]!;
    }
  },
};

/**
 * STDEV function - Sample standard deviation
 * Ignores text, logical values, and empty cells
 */
export const STDEV: FunctionDefinition = {
  name: "STDEV",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    const numbers: number[] = [];

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }

      if (typeof value === "number") {
        numbers.push(value);
      }
    }

    if (numbers.length < 2) {
      return "#DIV/0!";
    }

    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    const squaredDiffs = numbers.map((n) => Math.pow(n - mean, 2));
    const variance =
      squaredDiffs.reduce((sum, d) => sum + d, 0) / (numbers.length - 1);

    return Math.sqrt(variance);
  },
};

/**
 * VAR function - Sample variance
 * Ignores text, logical values, and empty cells
 */
export const VAR: FunctionDefinition = {
  name: "VAR",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const flattened = flattenValues(args);
    const numbers: number[] = [];

    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }

      if (typeof value === "number") {
        numbers.push(value);
      }
    }

    if (numbers.length < 2) {
      return "#DIV/0!";
    }

    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    const squaredDiffs = numbers.map((n) => Math.pow(n - mean, 2));
    const variance =
      squaredDiffs.reduce((sum, d) => sum + d, 0) / (numbers.length - 1);

    return variance;
  },
};

/**
 * SUMPRODUCT function - Returns the sum of the products of corresponding values
 * SUMPRODUCT(array1, [array2], ...)
 */
export const SUMPRODUCT: FunctionDefinition = {
  name: "SUMPRODUCT",
  minArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    const error = propagateError(args);
    if (error) return error;

    const arrays: CellValue[][][] = args.map((arg) => {
      if (Array.isArray(arg)) {
        if (arg.length > 0 && Array.isArray(arg[0])) {
          return arg as CellValue[][];
        }
        return [arg as CellValue[]];
      }
      return [[arg]];
    });

    const rows = arrays[0]?.length ?? 0;
    const cols = rows > 0 ? (arrays[0]![0]?.length ?? 0) : 0;
    for (const arr of arrays) {
      if (arr.length !== rows) return "#VALUE!";
      for (let r = 0; r < rows; r++) {
        if ((arr[r]?.length ?? 0) !== cols) return "#VALUE!";
      }
    }

    let sum = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let product = 1;
        for (const arr of arrays) {
          const v = arr[r]![c]!;
          if (isFormulaError(v)) return v;
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
    return sum;
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
  AVERAGE,
  MAX,
  MIN,
  MEDIAN,
  STDEV,
  VAR,
  SUMPRODUCT,
];
