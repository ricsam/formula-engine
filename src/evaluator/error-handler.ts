/**
 * Error management system for FormulaEngine
 * Handles error propagation, recovery, and formatting
 */

import type { FormulaError, CellValue, SimpleCellAddress } from '../core/types';

/**
 * Error context information for debugging
 */
export interface ErrorContext {
  formula?: string;
  cellAddress?: SimpleCellAddress;
  functionName?: string;
  argumentIndex?: number;
  message?: string;
}

/**
 * Extended error information
 */
export interface ExtendedError {
  type: FormulaError;
  context: ErrorContext;
  cause?: Error | ExtendedError;
}

/**
 * Checks if a value is a formula error
 */
export function isFormulaError(value: CellValue): value is FormulaError {
  if (typeof value !== 'string') return false;
  
  // Check for all known formula errors
  const errors: FormulaError[] = [
    '#DIV/0!',
    '#N/A',
    '#NAME?',
    '#NUM!',
    '#REF!',
    '#VALUE!',
    '#CYCLE!',
    '#ERROR!'
  ];
  
  return errors.includes(value as FormulaError);
}

/**
 * Maps JavaScript errors to formula errors
 */
export function mapJSErrorToFormulaError(error: Error): FormulaError {
  const message = error.message.toLowerCase();
  
  if (message.includes('division by zero') || message.includes('divide by zero')) {
    return '#DIV/0!';
  }
  if (message.includes('circular') || message.includes('cycle')) {
    return '#CYCLE!';
  }
  if (message.includes('invalid reference') || (message.includes('reference') && !message.includes('circular'))) {
    return '#REF!';
  }
  if (message.includes('invalid name') || message.includes('unknown function')) {
    return '#NAME?';
  }
  if (message.includes('invalid number') || message.includes('nan') || message.includes('infinity')) {
    return '#NUM!';
  }
  if (message.includes('type') || message.includes('invalid argument')) {
    return '#VALUE!';
  }
  if (message.includes('not available') || message.includes('n/a')) {
    return '#N/A';
  }
  
  return '#ERROR!';
}

/**
 * Creates a formula error with context
 */
export function createFormulaError(
  type: FormulaError,
  context?: ErrorContext
): ExtendedError {
  return {
    type,
    context: context || {}
  };
}

/**
 * Error propagation handler for array operations
 */
export function propagateError(values: CellValue[]): FormulaError | null {
  for (const value of values) {
    if (isFormulaError(value)) {
      return value;
    }
  }
  return null;
}

/**
 * Error propagation for 2D arrays
 */
export function propagateError2D(values: CellValue[][]): FormulaError | null {
  for (const row of values) {
    const error = propagateError(row);
    if (error) return error;
  }
  return null;
}

/**
 * Validates numeric arguments and returns error if invalid
 */
export function validateNumericArgument(
  value: CellValue,
  argumentName?: string
): FormulaError | null {
  if (isFormulaError(value)) {
    return value;
  }
  
  if (typeof value === 'number') {
    if (!isFinite(value)) {
      return '#NUM!';
    }
    return null;
  }
  
  if (typeof value === 'string') {
    // Trim whitespace
    const trimmed = value.trim();
    if (trimmed === '') return '#VALUE!';
    
    const num = parseFloat(trimmed);
    if (!isNaN(num) && isFinite(num)) {
      // Check for valid number format (reject things like "12.34.56")
      if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
        return null;
      }
    }
    return '#VALUE!';
  }
  
  if (typeof value === 'boolean') {
    return null; // Booleans can be coerced to numbers
  }
  
  if (value === undefined) {
    return null; // Empty cells treated as 0
  }
  
  return '#VALUE!';
}

/**
 * Validates that a value can be converted to text
 */
export function validateTextArgument(
  value: CellValue,
  argumentName?: string
): FormulaError | null {
  if (isFormulaError(value)) {
    return value;
  }
  
  // All non-error values can be converted to text
  return null;
}

/**
 * Validates boolean arguments
 */
export function validateBooleanArgument(
  value: CellValue,
  argumentName?: string
): FormulaError | null {
  if (isFormulaError(value)) {
    return value;
  }
  
  if (typeof value === 'boolean') {
    return null;
  }
  
  if (typeof value === 'number' || typeof value === 'string') {
    // Numbers and strings can be coerced to booleans
    return null;
  }
  
  if (value === undefined) {
    return null; // Empty cells treated as FALSE
  }
  
  return '#VALUE!';
}

/**
 * Formats error for display
 */
export function formatError(error: FormulaError | ExtendedError): string {
  if (typeof error === 'string') {
    return error;
  }
  
  return error.type;
}

/**
 * Creates detailed error message for debugging
 */
export function getDetailedErrorMessage(error: ExtendedError): string {
  const parts: string[] = [error.type];
  
  if (error.context.message) {
    parts.push(`: ${error.context.message}`);
  }
  
  if (error.context.cellAddress) {
    const addr = error.context.cellAddress;
    parts.push(` at cell ${addr.sheet}:${addr.col}:${addr.row}`);
  }
  
  if (error.context.functionName) {
    parts.push(` in function ${error.context.functionName}`);
    if (error.context.argumentIndex !== undefined) {
      parts.push(` (argument ${error.context.argumentIndex + 1})`);
    }
  }
  
  if (error.context.formula) {
    parts.push(` in formula: ${error.context.formula}`);
  }
  
  return parts.join('');
}

/**
 * Error recovery strategies
 */
export interface ErrorRecoveryStrategy {
  onError: (error: ExtendedError) => CellValue;
}

/**
 * Default error recovery - just returns the error
 */
export const defaultErrorRecovery: ErrorRecoveryStrategy = {
  onError: (error) => error.type
};

/**
 * Suppress errors recovery - returns 0 for numeric contexts
 */
export const suppressErrorsRecovery: ErrorRecoveryStrategy = {
  onError: (error) => {
    if (error.type === '#N/A') {
      return undefined; // Keep N/A as undefined
    }
    return 0; // Convert other errors to 0
  }
};

/**
 * Error handler class for managing error state
 */
export class ErrorHandler {
  private errors: Map<string, ExtendedError> = new Map();
  private strategy: ErrorRecoveryStrategy;
  
  constructor(strategy: ErrorRecoveryStrategy = defaultErrorRecovery) {
    this.strategy = strategy;
  }
  
  /**
   * Records an error for a cell
   */
  recordError(
    cellKey: string,
    error: FormulaError | ExtendedError,
    context?: ErrorContext
  ): void {
    const extendedError = typeof error === 'string' 
      ? createFormulaError(error, context)
      : { ...error, context: { ...error.context, ...context } };
      
    this.errors.set(cellKey, extendedError);
  }
  
  /**
   * Gets error for a cell
   */
  getError(cellKey: string): ExtendedError | undefined {
    return this.errors.get(cellKey);
  }
  
  /**
   * Clears error for a cell
   */
  clearError(cellKey: string): void {
    this.errors.delete(cellKey);
  }
  
  /**
   * Clears all errors
   */
  clearAllErrors(): void {
    this.errors.clear();
  }
  
  /**
   * Gets all cells with errors
   */
  getErrorCells(): string[] {
    return Array.from(this.errors.keys());
  }
  
  /**
   * Handles error with recovery strategy
   */
  handleError(error: FormulaError | ExtendedError): CellValue {
    const extendedError = typeof error === 'string'
      ? createFormulaError(error)
      : error;
      
    return this.strategy.onError(extendedError);
  }
  
  /**
   * Sets error recovery strategy
   */
  setStrategy(strategy: ErrorRecoveryStrategy): void {
    this.strategy = strategy;
  }
}

/**
 * Global error messages for common errors
 */
export const ERROR_MESSAGES = {
  DIVISION_BY_ZERO: 'Division by zero',
  INVALID_REFERENCE: 'Invalid cell reference',
  UNKNOWN_NAME: 'Unknown name or function',
  INVALID_NUMBER: 'Invalid numeric value',
  TYPE_MISMATCH: 'Type mismatch in operation',
  CIRCULAR_REFERENCE: 'Circular reference detected',
  VALUE_NOT_AVAILABLE: 'Value not available',
  GENERAL_ERROR: 'Formula evaluation error',
  SPILL_BLOCKED: 'Spill range isn\'t blank'
};

/**
 * Creates error with standard message
 */
export function createStandardError(
  type: FormulaError,
  additionalContext?: Partial<ErrorContext>
): ExtendedError {
  const messageMap: Record<FormulaError, string> = {
    '#DIV/0!': ERROR_MESSAGES.DIVISION_BY_ZERO,
    '#REF!': ERROR_MESSAGES.INVALID_REFERENCE,
    '#NAME?': ERROR_MESSAGES.UNKNOWN_NAME,
    '#NUM!': ERROR_MESSAGES.INVALID_NUMBER,
    '#VALUE!': ERROR_MESSAGES.TYPE_MISMATCH,
    '#CYCLE!': ERROR_MESSAGES.CIRCULAR_REFERENCE,
    '#N/A': ERROR_MESSAGES.VALUE_NOT_AVAILABLE,
    '#ERROR!': ERROR_MESSAGES.GENERAL_ERROR,
    '#SPILL!': ERROR_MESSAGES.SPILL_BLOCKED
  };
  
  return createFormulaError(type, {
    message: messageMap[type],
    ...additionalContext
  });
}
