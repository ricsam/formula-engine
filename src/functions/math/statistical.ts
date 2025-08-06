import type { CellValue } from '../../core/types';
import type { FunctionDefinition, EvaluationContext } from '../../evaluator/evaluator';
import { coerceToNumber, isFormulaError, propagateError } from '../index';

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
  name: 'SUM',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
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
      if (value === undefined || value === null || 
          typeof value === 'string' || typeof value === 'boolean') {
        continue;
      }
      
      if (typeof value === 'number') {
        sum += value;
        hasNumbers = true;
      }
    }
    
    return sum;
  }
};

/**
 * PRODUCT function - Multiplies all numbers in the arguments
 * Ignores text, logical values, and empty cells
 */
export const PRODUCT: FunctionDefinition = {
  name: 'PRODUCT',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
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
      if (value === undefined || value === null || 
          typeof value === 'string' || typeof value === 'boolean') {
        continue;
      }
      
      if (typeof value === 'number') {
        product *= value;
        hasNumbers = true;
      }
    }
    
    return hasNumbers ? product : 0;
  }
};

/**
 * COUNT function - Counts cells containing numbers
 * Ignores text, logical values, errors, and empty cells
 */
export const COUNT: FunctionDefinition = {
  name: 'COUNT',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    // Note: COUNT doesn't propagate errors, it just ignores them
    const flattened = flattenValues(args);
    let count = 0;
    
    for (const value of flattened) {
      // COUNT ignores errors and counts only numbers (Excel behavior)
      if (typeof value === 'number') {
        count++;
      }
    }
    
    return count;
  }
};

/**
 * COUNTBLANK function - Counts empty cells in a range
 */
export const COUNTBLANK: FunctionDefinition = {
  name: 'COUNTBLANK',
  minArgs: 1,
  maxArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    let count = 0;
    
    for (const value of flattened) {
      if (value === undefined || value === null || value === '') {
        count++;
      }
    }
    
    return count;
  }
};

/**
 * AVERAGE function - Returns the average of numbers
 * Ignores text, logical values, and empty cells
 */
export const AVERAGE: FunctionDefinition = {
  name: 'AVERAGE',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    let sum = 0;
    let count = 0;
    
    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }
      
      if (typeof value === 'number') {
        sum += value;
        count++;
      }
    }
    
    if (count === 0) {
      return '#DIV/0!';
    }
    
    return sum / count;
  }
};

/**
 * MAX function - Returns the maximum value
 * Ignores text, logical values, and empty cells
 */
export const MAX: FunctionDefinition = {
  name: 'MAX',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    let max: number | null = null;
    
    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }
      
      if (typeof value === 'number') {
        if (max === null || value > max) {
          max = value;
        }
      }
    }
    
    return max === null ? 0 : max;
  }
};

/**
 * MIN function - Returns the minimum value
 * Ignores text, logical values, and empty cells
 */
export const MIN: FunctionDefinition = {
  name: 'MIN',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    let min: number | null = null;
    
    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }
      
      if (typeof value === 'number') {
        if (min === null || value < min) {
          min = value;
        }
      }
    }
    
    return min === null ? 0 : min;
  }
};

/**
 * MEDIAN function - Returns the median value
 * Ignores text, logical values, and empty cells
 */
export const MEDIAN: FunctionDefinition = {
  name: 'MEDIAN',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    const numbers: number[] = [];
    
    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }
      
      if (typeof value === 'number') {
        numbers.push(value);
      }
    }
    
    if (numbers.length === 0) {
      return '#NUM!';
    }
    
    // Sort numbers
    numbers.sort((a, b) => a - b);
    
    const mid = Math.floor(numbers.length / 2);
    if (numbers.length % 2 === 0) {
      return (numbers[mid - 1]! + numbers[mid]!) / 2;
    } else {
      return numbers[mid]!;
    }
  }
};

/**
 * STDEV function - Sample standard deviation
 * Ignores text, logical values, and empty cells
 */
export const STDEV: FunctionDefinition = {
  name: 'STDEV',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    const numbers: number[] = [];
    
    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }
      
      if (typeof value === 'number') {
        numbers.push(value);
      }
    }
    
    if (numbers.length < 2) {
      return '#DIV/0!';
    }
    
    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    const squaredDiffs = numbers.map(n => Math.pow(n - mean, 2));
    const variance = squaredDiffs.reduce((sum, d) => sum + d, 0) / (numbers.length - 1);
    
    return Math.sqrt(variance);
  }
};

/**
 * VAR function - Sample variance
 * Ignores text, logical values, and empty cells
 */
export const VAR: FunctionDefinition = {
  name: 'VAR',
  minArgs: 1,
  acceptsArrays: true,
  evaluate: (args: CellValue[]): CellValue => {
    const error = propagateError(args);
    if (error) return error;
    
    const flattened = flattenValues(args);
    const numbers: number[] = [];
    
    for (const value of flattened) {
      if (isFormulaError(value)) {
        return value;
      }
      
      if (typeof value === 'number') {
        numbers.push(value);
      }
    }
    
    if (numbers.length < 2) {
      return '#DIV/0!';
    }
    
    const mean = numbers.reduce((sum, n) => sum + n, 0) / numbers.length;
    const squaredDiffs = numbers.map(n => Math.pow(n - mean, 2));
    const variance = squaredDiffs.reduce((sum, d) => sum + d, 0) / (numbers.length - 1);
    
    return variance;
  }
};

// Export all statistical functions
export const statisticalFunctions: FunctionDefinition[] = [
  SUM,
  PRODUCT,
  COUNT,
  COUNTBLANK,
  AVERAGE,
  MAX,
  MIN,
  MEDIAN,
  STDEV,
  VAR
];
