import type { CellValue } from "../../core/types";
import type { FunctionDefinition } from "../../evaluator/evaluator";
import {
  coerceToString,
  coerceToNumber,
  isFormulaError,
  propagateError,
  validateArgCount,
} from "../index";

// Helper function to compare two values
function compareValues(a: CellValue, b: CellValue): number {
  // Handle errors first
  if (isFormulaError(a)) return NaN;
  if (isFormulaError(b)) return NaN;

  // Handle undefined/null values
  if (a === undefined || a === null) a = 0;
  if (b === undefined || b === null) b = 0;

  // If both are numbers, compare numerically
  if (typeof a === "number" && typeof b === "number") {
    return a - b;
  }

  // If both are strings, compare lexicographically
  if (typeof a === "string" && typeof b === "string") {
    return a.localeCompare(b);
  }

  // If both are booleans, compare as numbers
  if (typeof a === "boolean" && typeof b === "boolean") {
    return (a ? 1 : 0) - (b ? 1 : 0);
  }

  // Mixed types: try to coerce to numbers first
  try {
    const numA = typeof a === "number" ? a : coerceToNumber(a);
    const numB = typeof b === "number" ? b : coerceToNumber(b);
    return numA - numB;
  } catch {
    // If numeric coercion fails, coerce to strings
    const strA = coerceToString(a);
    const strB = coerceToString(b);
    return strA.localeCompare(strB);
  }
}

// FE.EQ(a, b) - Equality comparison
export const EQ: FunctionDefinition = {
  name: "FE.EQ",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.EQ", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      return compareValues(args[0], args[1]) === 0;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// FE.NE(a, b) - Inequality comparison
export const NE: FunctionDefinition = {
  name: "FE.NE",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.NE", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      return compareValues(args[0], args[1]) !== 0;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// FE.LT(a, b) - Less than comparison
export const LT: FunctionDefinition = {
  name: "FE.LT",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.LT", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      return compareValues(args[0], args[1]) < 0;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// FE.LTE(a, b) - Less than or equal comparison
export const LTE: FunctionDefinition = {
  name: "FE.LTE",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.LTE", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      return compareValues(args[0], args[1]) <= 0;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// FE.GT(a, b) - Greater than comparison
export const GT: FunctionDefinition = {
  name: "FE.GT",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.GT", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      return compareValues(args[0], args[1]) > 0;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// FE.GTE(a, b) - Greater than or equal comparison
export const GTE: FunctionDefinition = {
  name: "FE.GTE",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.GTE", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      return compareValues(args[0], args[1]) >= 0;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// Export all comparison functions
export const logicalComparisonFunctions: FunctionDefinition[] = [
  EQ,
  NE,
  LT,
  LTE,
  GT,
  GTE,
];
