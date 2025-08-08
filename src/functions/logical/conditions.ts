import type { CellValue } from '../../core/types';
import type { FunctionDefinition, EvaluationContext } from '../../evaluator/evaluator';
import { isFormulaError, propagateError } from '../index';

// Helper function to coerce values to boolean
function coerceToBoolean(value: CellValue): boolean {
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value !== 0;
  }
  if (typeof value === 'string') {
    const upper = value.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    return value.length > 0;
  }
  if (value === undefined || value === null) {
    return false;
  }
  return false;
}

// IF(condition, value_if_true, value_if_false)
export const IF: FunctionDefinition = {
  name: 'IF',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length < 2 || args.length > 3) {
      throw new Error('#VALUE!');
    }

    // Only check for errors in the condition argument
    if (isFormulaError(args[0])) {
      return args[0];
    }

    const condition = coerceToBoolean(args[0]);
    const valueIfTrue = args[1];
    const valueIfFalse = args.length === 3 ? args[2] : false;

    return condition ? valueIfTrue : valueIfFalse;
  }
};

// NOT(logical)
export const NOT: FunctionDefinition = {
  name: 'NOT',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 1) {
      throw new Error('#VALUE!');
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    const logical = coerceToBoolean(args[0]);
    return !logical;
  }
};

// OR(logical1, [logical2], ...)
export const OR: FunctionDefinition = {
  name: 'OR',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length === 0) {
      throw new Error('#VALUE!');
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    // OR returns TRUE if any argument is TRUE
    for (const arg of args) {
      if (coerceToBoolean(arg)) {
        return true;
      }
    }
    return false;
  }
};

// AND(logical1, [logical2], ...)
export const AND: FunctionDefinition = {
  name: 'AND',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length === 0) {
      throw new Error('#VALUE!');
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    // AND returns FALSE if any argument is FALSE
    for (const arg of args) {
      if (!coerceToBoolean(arg)) {
        return false;
      }
    }
    return true;
  }
};

// TRUE()
export const TRUE_FUNC: FunctionDefinition = {
  name: 'TRUE',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 0) {
      throw new Error('#VALUE!');
    }
    return true;
  }
};

// FALSE()
export const FALSE_FUNC: FunctionDefinition = {
  name: 'FALSE',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 0) {
      throw new Error('#VALUE!');
    }
    return false;
  }
};

// IFERROR(value, value_if_error)
export const IFERROR: FunctionDefinition = {
  name: 'IFERROR',
  evaluate: ({ argValues: args }): CellValue => {
    if (args.length !== 2) {
      throw new Error('#VALUE!');
    }
    const value = args[0];
    if (isFormulaError(value)) {
      return args[1];
    }
    return value;
  }
};

// Export all logical condition functions
export const logicalConditionFunctions: FunctionDefinition[] = [
  IF,
  NOT,
  OR,
  AND,
  TRUE_FUNC,
  FALSE_FUNC,
  IFERROR
];