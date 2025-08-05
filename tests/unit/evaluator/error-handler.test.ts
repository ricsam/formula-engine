import { test, expect, describe } from "bun:test";
import {
  isFormulaError,
  mapJSErrorToFormulaError,
  createFormulaError,
  propagateError,
  propagateError2D,
  validateNumericArgument,
  validateTextArgument,
  validateBooleanArgument,
  formatError,
  getDetailedErrorMessage,
  createStandardError,
  ErrorHandler,
  defaultErrorRecovery,
  suppressErrorsRecovery,
  ERROR_MESSAGES
} from '../../../src/evaluator/error-handler';
import type { CellValue, FormulaError } from '../../../src/core/types';

describe('Error Handler', () => {
  describe('isFormulaError', () => {
    test('should identify formula errors', () => {
      expect(isFormulaError('#DIV/0!')).toBe(true);
      expect(isFormulaError('#N/A')).toBe(true);
      expect(isFormulaError('#NAME?')).toBe(true);
      expect(isFormulaError('#NUM!')).toBe(true);
      expect(isFormulaError('#REF!')).toBe(true);
      expect(isFormulaError('#VALUE!')).toBe(true);
      expect(isFormulaError('#CYCLE!')).toBe(true);
      expect(isFormulaError('#ERROR!')).toBe(true);
    });

    test('should not identify non-errors as formula errors', () => {
      expect(isFormulaError('DIV/0')).toBe(false);
      expect(isFormulaError('#DIV')).toBe(false);
      expect(isFormulaError('ERROR!')).toBe(false);
      expect(isFormulaError(42)).toBe(false);
      expect(isFormulaError(true)).toBe(false);
      expect(isFormulaError(undefined)).toBe(false);
    });
  });

  describe('mapJSErrorToFormulaError', () => {
    test('should map division by zero errors', () => {
      expect(mapJSErrorToFormulaError(new Error('division by zero'))).toBe('#DIV/0!');
      expect(mapJSErrorToFormulaError(new Error('Cannot divide by zero'))).toBe('#DIV/0!');
    });

    test('should map reference errors', () => {
      expect(mapJSErrorToFormulaError(new Error('Invalid reference'))).toBe('#REF!');
      expect(mapJSErrorToFormulaError(new Error('Cell reference not found'))).toBe('#REF!');
    });

    test('should map name errors', () => {
      expect(mapJSErrorToFormulaError(new Error('Invalid name'))).toBe('#NAME?');
      expect(mapJSErrorToFormulaError(new Error('Unknown function: FOO'))).toBe('#NAME?');
    });

    test('should map numeric errors', () => {
      expect(mapJSErrorToFormulaError(new Error('Invalid number'))).toBe('#NUM!');
      expect(mapJSErrorToFormulaError(new Error('Result is NaN'))).toBe('#NUM!');
      expect(mapJSErrorToFormulaError(new Error('Infinity'))).toBe('#NUM!');
    });

    test('should map type errors', () => {
      expect(mapJSErrorToFormulaError(new Error('Invalid argument type'))).toBe('#VALUE!');
      expect(mapJSErrorToFormulaError(new Error('Type mismatch'))).toBe('#VALUE!');
    });

    test('should map circular reference errors', () => {
      expect(mapJSErrorToFormulaError(new Error('Circular reference detected'))).toBe('#CYCLE!');
      expect(mapJSErrorToFormulaError(new Error('Formula has cycle'))).toBe('#CYCLE!');
    });

    test('should map N/A errors', () => {
      expect(mapJSErrorToFormulaError(new Error('Value not available'))).toBe('#N/A');
      expect(mapJSErrorToFormulaError(new Error('N/A'))).toBe('#N/A');
    });

    test('should default to general error', () => {
      expect(mapJSErrorToFormulaError(new Error('Something went wrong'))).toBe('#ERROR!');
      expect(mapJSErrorToFormulaError(new Error(''))).toBe('#ERROR!');
    });
  });

  describe('createFormulaError', () => {
    test('should create error with context', () => {
      const error = createFormulaError('#VALUE!', {
        message: 'Type mismatch',
        functionName: 'SUM',
        argumentIndex: 1
      });

      expect(error.type).toBe('#VALUE!');
      expect(error.context.message).toBe('Type mismatch');
      expect(error.context.functionName).toBe('SUM');
      expect(error.context.argumentIndex).toBe(1);
    });

    test('should create error without context', () => {
      const error = createFormulaError('#DIV/0!');
      expect(error.type).toBe('#DIV/0!');
      expect(error.context).toEqual({});
    });
  });

  describe('propagateError', () => {
    test('should return first error found', () => {
      const values: CellValue[] = [1, 2, '#VALUE!', 3, '#REF!'];
      expect(propagateError(values)).toBe('#VALUE!');
    });

    test('should return null if no errors', () => {
      const values: CellValue[] = [1, 2, 3, 'text', true, undefined];
      expect(propagateError(values)).toBeNull();
    });

    test('should handle empty array', () => {
      expect(propagateError([])).toBeNull();
    });
  });

  describe('propagateError2D', () => {
    test('should return first error found in 2D array', () => {
      const values: CellValue[][] = [
        [1, 2, 3],
        [4, '#NUM!', 6],
        [7, 8, '#REF!']
      ];
      expect(propagateError2D(values)).toBe('#NUM!');
    });

    test('should return null if no errors in 2D array', () => {
      const values: CellValue[][] = [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
      ];
      expect(propagateError2D(values)).toBeNull();
    });

    test('should handle empty 2D array', () => {
      expect(propagateError2D([])).toBeNull();
      expect(propagateError2D([[]])).toBeNull();
    });
  });

  describe('validateNumericArgument', () => {
    test('should accept valid numbers', () => {
      expect(validateNumericArgument(42)).toBeNull();
      expect(validateNumericArgument(-3.14)).toBeNull();
      expect(validateNumericArgument(0)).toBeNull();
    });

    test('should accept numeric strings', () => {
      expect(validateNumericArgument('42')).toBeNull();
      expect(validateNumericArgument('-3.14')).toBeNull();
      expect(validateNumericArgument('0')).toBeNull();
    });

    test('should accept booleans', () => {
      expect(validateNumericArgument(true)).toBeNull();
      expect(validateNumericArgument(false)).toBeNull();
    });

    test('should accept undefined as 0', () => {
      expect(validateNumericArgument(undefined)).toBeNull();
    });

    test('should return error for non-finite numbers', () => {
      expect(validateNumericArgument(Infinity)).toBe('#NUM!');
      expect(validateNumericArgument(-Infinity)).toBe('#NUM!');
      expect(validateNumericArgument(NaN)).toBe('#NUM!');
    });

    test('should return error for non-numeric strings', () => {
      expect(validateNumericArgument('abc')).toBe('#VALUE!');
      expect(validateNumericArgument('12.34.56')).toBe('#VALUE!');
    });

    test('should propagate errors', () => {
      expect(validateNumericArgument('#DIV/0!' as FormulaError)).toBe('#DIV/0!');
      expect(validateNumericArgument('#REF!' as FormulaError)).toBe('#REF!');
    });
  });

  describe('validateTextArgument', () => {
    test('should accept any non-error value', () => {
      expect(validateTextArgument('text')).toBeNull();
      expect(validateTextArgument(42)).toBeNull();
      expect(validateTextArgument(true)).toBeNull();
      expect(validateTextArgument(undefined)).toBeNull();
    });

    test('should propagate errors', () => {
      expect(validateTextArgument('#VALUE!' as FormulaError)).toBe('#VALUE!');
    });
  });

  describe('validateBooleanArgument', () => {
    test('should accept booleans', () => {
      expect(validateBooleanArgument(true)).toBeNull();
      expect(validateBooleanArgument(false)).toBeNull();
    });

    test('should accept values coercible to boolean', () => {
      expect(validateBooleanArgument(1)).toBeNull();
      expect(validateBooleanArgument(0)).toBeNull();
      expect(validateBooleanArgument('TRUE')).toBeNull();
      expect(validateBooleanArgument('FALSE')).toBeNull();
      expect(validateBooleanArgument(undefined)).toBeNull();
    });

    test('should propagate errors', () => {
      expect(validateBooleanArgument('#NAME?' as FormulaError)).toBe('#NAME?');
    });
  });

  describe('formatError', () => {
    test('should format simple error', () => {
      expect(formatError('#DIV/0!')).toBe('#DIV/0!');
    });

    test('should format extended error', () => {
      const error = createFormulaError('#VALUE!', { message: 'Type error' });
      expect(formatError(error)).toBe('#VALUE!');
    });
  });

  describe('getDetailedErrorMessage', () => {
    test('should create detailed error message', () => {
      const error = createFormulaError('#VALUE!', {
        message: 'Type mismatch',
        cellAddress: { sheet: 0, col: 1, row: 2 },
        functionName: 'SUM',
        argumentIndex: 0,
        formula: '=SUM("text")'
      });

      const message = getDetailedErrorMessage(error);
      expect(message).toContain('#VALUE!');
      expect(message).toContain('Type mismatch');
      expect(message).toContain('at cell 0:1:2');
      expect(message).toContain('in function SUM');
      expect(message).toContain('(argument 1)');
      expect(message).toContain('in formula: =SUM("text")');
    });

    test('should handle partial context', () => {
      const error = createFormulaError('#DIV/0!', {
        cellAddress: { sheet: 0, col: 0, row: 0 }
      });

      const message = getDetailedErrorMessage(error);
      expect(message).toBe('#DIV/0! at cell 0:0:0');
    });
  });

  describe('createStandardError', () => {
    test('should create error with standard message', () => {
      const error = createStandardError('#DIV/0!');
      expect(error.type).toBe('#DIV/0!');
      expect(error.context.message).toBe(ERROR_MESSAGES.DIVISION_BY_ZERO);
    });

    test('should merge additional context', () => {
      const error = createStandardError('#VALUE!', {
        functionName: 'VLOOKUP',
        argumentIndex: 2
      });

      expect(error.type).toBe('#VALUE!');
      expect(error.context.message).toBe(ERROR_MESSAGES.TYPE_MISMATCH);
      expect(error.context.functionName).toBe('VLOOKUP');
      expect(error.context.argumentIndex).toBe(2);
    });
  });

  describe('ErrorHandler', () => {
    test('should record and retrieve errors', () => {
      const handler = new ErrorHandler();
      
      handler.recordError('A1', '#DIV/0!', { formula: '=1/0' });
      
      const error = handler.getError('A1');
      expect(error?.type).toBe('#DIV/0!');
      expect(error?.context.formula).toBe('=1/0');
    });

    test('should clear errors', () => {
      const handler = new ErrorHandler();
      
      handler.recordError('A1', '#VALUE!');
      handler.recordError('B1', '#REF!');
      
      handler.clearError('A1');
      expect(handler.getError('A1')).toBeUndefined();
      expect(handler.getError('B1')).toBeDefined();
    });

    test('should clear all errors', () => {
      const handler = new ErrorHandler();
      
      handler.recordError('A1', '#VALUE!');
      handler.recordError('B1', '#REF!');
      
      handler.clearAllErrors();
      expect(handler.getError('A1')).toBeUndefined();
      expect(handler.getError('B1')).toBeUndefined();
    });

    test('should get all error cells', () => {
      const handler = new ErrorHandler();
      
      handler.recordError('A1', '#VALUE!');
      handler.recordError('B2', '#REF!');
      handler.recordError('C3', '#NAME?');
      
      const errorCells = handler.getErrorCells();
      expect(errorCells).toHaveLength(3);
      expect(errorCells).toContain('A1');
      expect(errorCells).toContain('B2');
      expect(errorCells).toContain('C3');
    });

    test('should handle errors with default strategy', () => {
      const handler = new ErrorHandler();
      
      expect(handler.handleError('#DIV/0!')).toBe('#DIV/0!');
      
      const extendedError = createFormulaError('#VALUE!');
      expect(handler.handleError(extendedError)).toBe('#VALUE!');
    });

    test('should handle errors with suppress strategy', () => {
      const handler = new ErrorHandler(suppressErrorsRecovery);
      
      expect(handler.handleError('#DIV/0!')).toBe(0);
      expect(handler.handleError('#VALUE!')).toBe(0);
      expect(handler.handleError('#N/A')).toBeUndefined();
    });

    test('should change error recovery strategy', () => {
      const handler = new ErrorHandler();
      
      expect(handler.handleError('#VALUE!')).toBe('#VALUE!');
      
      handler.setStrategy(suppressErrorsRecovery);
      expect(handler.handleError('#VALUE!')).toBe(0);
      
      handler.setStrategy(defaultErrorRecovery);
      expect(handler.handleError('#VALUE!')).toBe('#VALUE!');
    });
  });
});