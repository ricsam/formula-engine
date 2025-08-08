import { FormulaError } from "../../core/types";
import type { FunctionDefinition, FunctionEvaluationResult } from "../../evaluator/evaluator";
import {
  coerceToNumber,
  isFormulaError,
  propagateError,
  validateArgCount,
} from "../utils";

// Basic arithmetic operators

const ADD: FunctionDefinition = {
  name: "FE.ADD",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.ADD", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);
      return { type: "value", value: a + b };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const MINUS: FunctionDefinition = {
  name: "FE.MINUS",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.MINUS", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);
      return { type: "value", value: a - b };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const MULTIPLY: FunctionDefinition = {
  name: "FE.MULTIPLY",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.MULTIPLY", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);
      return { type: "value", value: a * b };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const DIVIDE: FunctionDefinition = {
  name: "FE.DIVIDE",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.DIVIDE", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const a = coerceToNumber(args[0]);
      const b = coerceToNumber(args[1]);

      if (b === 0) {
        return { type: "value", value: FormulaError.DIV0 };
      }

      return { type: "value", value: a / b };
    } catch (e) {
        return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const POW: FunctionDefinition = {
  name: "FE.POW",
  minArgs: 2,
  maxArgs: 2,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.POW", args, 2, 2);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const base = coerceToNumber(args[0]);
      const exponent = coerceToNumber(args[1]);

      const result = Math.pow(base, exponent);

      if (!isFinite(result)) {
        return { type: "value", value: FormulaError.NUM };
      }

      return { type: "value", value: result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const UMINUS: FunctionDefinition = {
  name: "FE.UMINUS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.UMINUS", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      const result = -value;
      // Handle JavaScript's -0 case
      return { type: "value", value: Object.is(result, -0) ? 0 : result };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const UPLUS: FunctionDefinition = {
  name: "FE.UPLUS",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.UPLUS", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: +value };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
    }
  },
};

const UNARY_PERCENT: FunctionDefinition = {
  name: "FE.UNARY_PERCENT",
  minArgs: 1,
  maxArgs: 1,
  evaluate: ({ flatArgValues: args }): FunctionEvaluationResult => {
    validateArgCount("FE.UNARY_PERCENT", args, 1, 1);

    const error = propagateError(args);
    if (error) return { type: "value", value: error };

    try {
      const value = coerceToNumber(args[0]);
      return { type: "value", value: value / 100 };
    } catch (e) {
      return { type: "value", value: FormulaError.VALUE };
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
