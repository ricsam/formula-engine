import type { CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
} from "../../evaluator/evaluator";
import { isFormulaError, coerceToNumber, propagateError } from "../index";

// ISEVEN(number)
export const ISEVEN: FunctionDefinition = {
  name: "ISEVEN",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    try {
      const num = coerceToNumber(args[0]);
      // Round to nearest integer before checking
      const intNum = Math.round(num);
      return intNum % 2 === 0;
    } catch {
      return false;
    }
  },
};

// ISODD(number)
export const ISODD: FunctionDefinition = {
  name: "ISODD",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    try {
      const num = coerceToNumber(args[0]);
      // Round to nearest integer before checking
      const intNum = Math.round(num);
      return intNum % 2 !== 0;
    } catch {
      return false;
    }
  },
};

// ISBLANK(value)
export const ISBLANK: FunctionDefinition = {
  name: "ISBLANK",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];

    // If the value is an array (range), check if all cells are blank
    if (Array.isArray(value)) {
      // Flatten the array and check if all values are blank
      const checkBlank = (v: CellValue): boolean => {
        if (Array.isArray(v)) {
          return v.every(checkBlank);
        }
        return v === undefined || v === null || v === "";
      };
      return checkBlank(value);
    }

    // Single value check
    return value === undefined || value === null || value === "";
  },
};

// ISERROR(value)
export const ISERROR: FunctionDefinition = {
  name: "ISERROR",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return isFormulaError(value);
  },
};

// ISNA(value)
export const ISNA: FunctionDefinition = {
  name: "ISNA",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return value === "#N/A";
  },
};

// ISNUMBER(value)
export const ISNUMBER: FunctionDefinition = {
  name: "ISNUMBER",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return typeof value === "number";
  },
};

// ISTEXT(value)
export const ISTEXT: FunctionDefinition = {
  name: "ISTEXT",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return typeof value === "string" && !isFormulaError(value);
  },
};

// ISLOGICAL(value)
export const ISLOGICAL: FunctionDefinition = {
  name: "ISLOGICAL",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return typeof value === "boolean";
  },
};

// NA() - returns #N/A error
export const NA: FunctionDefinition = {
  name: "NA",
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 0) {
      throw new Error("#VALUE!");
    }
    return "#N/A";
  },
};

// Export all info functions
export const infoFunctions: FunctionDefinition[] = [
  ISEVEN,
  ISODD,
  ISBLANK,
  ISERROR,
  ISNA,
  ISNUMBER,
  ISTEXT,
  ISLOGICAL,
  NA,
];
