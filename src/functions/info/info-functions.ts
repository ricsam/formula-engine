import { FormulaError, type CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import { isFormulaError, coerceToNumber, propagateError } from "../utils";

// ISEVEN(number)
const ISEVEN: FunctionDefinition = {
  name: "ISEVEN",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const num = coerceToNumber(args[0]);
      // Round to nearest integer before checking
      const intNum = Math.round(num);
      return { type: "value", value: intNum % 2 === 0 };
    } catch {
      return { type: "value", value: false };
    }
  },
};

// ISODD(number)
const ISODD: FunctionDefinition = {
  name: "ISODD",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const num = coerceToNumber(args[0]);
      // Round to nearest integer before checking
      const intNum = Math.round(num);
      return { type: "value", value: intNum % 2 !== 0 };
    } catch {
      return { type: "value", value: false };
    }
  },
};

// ISBLANK(value)
const ISBLANK: FunctionDefinition = {
  name: "ISBLANK",
  evaluate: ({ argEvaluatedValues }): FunctionEvaluationResult => {
    if (argEvaluatedValues.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Don't propagate errors for IS* functions - they test the value
    const arg = argEvaluatedValues[0];
    if (!arg) {
      return { type: "value", value: true };
    }

    if (arg.type === "value") {
      // Single value check
      return {
        type: "value",
        value: arg.value === undefined || arg.value === null || arg.value === "",
      };
    } else if (arg.type === "2d-array") {
      // Range check - check if all cells are blank
      const checkBlank = (v: CellValue): boolean => {
        return v === undefined || v === null || v === "";
      };

      // Flatten the 2D array and check if all values are blank
      const flatValues: CellValue[] = [];
      for (const row of arg.value) {
        for (const cell of row) {
          flatValues.push(cell);
        }
      }

      // If the array is empty (no cells), consider it blank
      if (flatValues.length === 0) {
        return { type: "value", value: true };
      }

      // Check if all cells are blank
      return { type: "value", value: flatValues.every(checkBlank) };
    }

    // Fallback
    return { type: "value", value: false };
  },
};

// ISERROR(value)
const ISERROR: FunctionDefinition = {
  name: "ISERROR",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return { type: "value", value: isFormulaError(value) };
  },
};

// ISNA(value)
const ISNA: FunctionDefinition = {
  name: "ISNA",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return { type: "value", value: value === FormulaError.NA };
  },
};

// ISNUMBER(value)
const ISNUMBER: FunctionDefinition = {
  name: "ISNUMBER",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return { type: "value", value: typeof value === "number" };
  },
};

// ISTEXT(value)
const ISTEXT: FunctionDefinition = {
  name: "ISTEXT",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return {
      type: "value",
      value: typeof value === "string" && !isFormulaError(value),
    };
  },
};

// ISLOGICAL(value)
const ISLOGICAL: FunctionDefinition = {
  name: "ISLOGICAL",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error(FormulaError.VALUE);
    }

    // Don't propagate errors for IS* functions - they test the value
    const value = args[0];
    return { type: "value", value: typeof value === "boolean" };
  },
};

// NA() - returns #N/A error
const NA: FunctionDefinition = {
  name: "NA",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 0) {
      throw new Error(FormulaError.VALUE);
    }
    return { type: "value", value: FormulaError.NA };
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
