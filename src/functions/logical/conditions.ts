import type { CellValue } from "../../core/types";
import type {
  FunctionDefinition,
  EvaluationContext,
  FunctionEvaluationResult,
} from "../../evaluator/evaluator";
import { isFormulaError, propagateError } from "../utils";

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

// IF(condition, value_if_true, value_if_false)
const IF: FunctionDefinition = {
  name: "IF",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length < 2 || args.length > 3) {
      throw new Error("#VALUE!");
    }

    // Only check for errors in the condition argument
    if (isFormulaError(args[0])) {
      return { type: "value", value: args[0] };
    }

    const condition = coerceToBoolean(args[0]);
    const valueIfTrue = args[1];
    const valueIfFalse = args.length === 3 ? args[2] : false;

    return { type: "value", value: condition ? valueIfTrue : valueIfFalse };
  },
};

// NOT(logical)
const NOT: FunctionDefinition = {
  name: "NOT",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 1) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    const logical = coerceToBoolean(args[0]);
    return { type: "value", value: !logical };
  },
};

// OR(logical1, [logical2], ...)
const OR: FunctionDefinition = {
  name: "OR",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length === 0) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    // OR returns TRUE if any argument is TRUE
    for (const arg of args) {
      if (coerceToBoolean(arg)) {
        return { type: "value", value: true };
      }
    }
    return { type: "value", value: false };
  },
};

// AND(logical1, [logical2], ...)
const AND: FunctionDefinition = {
  name: "AND",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length === 0) {
      throw new Error("#VALUE!");
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    // AND returns FALSE if any argument is FALSE
    for (const arg of args) {
      if (!coerceToBoolean(arg)) {
        return { type: "value", value: false };
      }
    }
    return { type: "value", value: true };
  },
};

// TRUE()
const TRUE_FUNC: FunctionDefinition = {
  name: "TRUE",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 0) {
      throw new Error("#VALUE!");
    }
    return { type: "value", value: true };
  },
};

// FALSE()
const FALSE_FUNC: FunctionDefinition = {
  name: "FALSE",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 0) {
      throw new Error("#VALUE!");
    }
    return { type: "value", value: false };
  },
};

// IFERROR(value, value_if_error)
const IFERROR: FunctionDefinition = {
  name: "IFERROR",
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    if (args.length !== 2) {
      throw new Error("#VALUE!");
    }
    const value = args[0];
    if (isFormulaError(value)) {
      return { type: "value", value: args[1] };
    }
    return { type: "value", value: value };
  },
};

// Export all logical condition functions
export const logicalConditionFunctions: FunctionDefinition[] = [
  IF,
  NOT,
  OR,
  AND,
  TRUE_FUNC,
  FALSE_FUNC,
  IFERROR,
];
