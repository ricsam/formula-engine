import type { CellValue } from "../../core/types";
import type { FunctionDefinition } from "../../evaluator/evaluator";
import {
  coerceToNumber,
  isFormulaError,
  propagateError,
  validateArgCount,
} from "../index";

// Basic arithmetic operators

export const ADD: FunctionDefinition = {
  name: "FE.ADD",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.ADD", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);
      return a + b;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const MINUS: FunctionDefinition = {
  name: "FE.MINUS",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.MINUS", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);
      return a - b;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const MULTIPLY: FunctionDefinition = {
  name: "FE.MULTIPLY",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.MULTIPLY", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);
      return a * b;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const DIVIDE: FunctionDefinition = {
  name: "FE.DIVIDE",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.DIVIDE", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);

      if (b === 0) {
        return "#DIV/0!";
      }

      return a / b;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const POW: FunctionDefinition = {
  name: "FE.POW",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.POW", args, 2, 2);

    const error = propagateError(args);
    if (error) return error;

    try {
      const base = coerceToNumber(args[0]);
      const exponent = coerceToNumber(args[1]);

      const result = Math.pow(base, exponent);

      if (!isFinite(result)) {
        return "#NUM!";
      }

      return result;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const UMINUS: FunctionDefinition = {
  name: "FE.UMINUS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.UMINUS", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      const result = -value;
      // Handle JavaScript's -0 case
      return Object.is(result, -0) ? 0 : result;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const UPLUS: FunctionDefinition = {
  name: "FE.UPLUS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.UPLUS", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return +value;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

export const UNARY_PERCENT: FunctionDefinition = {
  name: "FE.UNARY_PERCENT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ argValues: args }): CellValue => {
    validateArgCount("FE.UNARY_PERCENT", args, 1, 1);

    const error = propagateError(args);
    if (error) return error;

    try {
      const value = coerceToNumber(args[0]);
      return value / 100;
    } catch (e) {
      return "#VALUE!";
    }
  },
};

// Export all basic math functions
export const basicMathFunctions: FunctionDefinition[] = [
  ADD,
  MINUS,
  MULTIPLY,
  DIVIDE,
  POW,
  UMINUS,
  UPLUS,
  UNARY_PERCENT,
];
