/**
 * Array formula evaluation with NumPy-style broadcasting
 * Handles spilling behavior and array operation optimization
 */

import type { CellValue, FormulaError } from '../core/types';
import { isFormulaError, propagateError, propagateError2D } from './error-handler';

/**
 * Array dimensions
 */
export interface ArrayDimensions {
  rows: number;
  cols: number;
}

/**
 * Broadcasting result
 */
export interface BroadcastResult {
  values: CellValue[][];
  dimensions: ArrayDimensions;
}

/**
 * Gets dimensions of an array or scalar
 */
export function getArrayDimensions(value: CellValue | CellValue[] | CellValue[][]): ArrayDimensions {
  if (!Array.isArray(value)) {
    return { rows: 1, cols: 1 };
  }
  
  if (!Array.isArray(value[0])) {
    // 1D array - treat as column vector
    return { rows: (value as CellValue[]).length, cols: 1 };
  }
  
  // 2D array
  const array2d = value as CellValue[][];
  return {
    rows: array2d.length,
    cols: array2d[0]?.length || 0
  };
}

/**
 * Converts a value to a 2D array
 */
export function to2DArray(value: CellValue | CellValue[] | CellValue[][]): CellValue[][] {
  if (!Array.isArray(value)) {
    return [[value]];
  }
  
  if (!Array.isArray(value[0])) {
    // 1D array - convert to column vector
    return (value as CellValue[]).map(v => [v]);
  }
  
  return value as CellValue[][];
}

/**
 * Flattens a 2D array to 1D
 */
export function flatten(array: CellValue[][]): CellValue[] {
  const result: CellValue[] = [];
  for (const row of array) {
    result.push(...row);
  }
  return result;
}

/**
 * Checks if dimensions are compatible for broadcasting
 */
export function areDimensionsCompatible(dim1: ArrayDimensions, dim2: ArrayDimensions): boolean {
  // Scalars are always compatible
  if (dim1.rows === 1 && dim1.cols === 1) return true;
  if (dim2.rows === 1 && dim2.cols === 1) return true;
  
  // Check NumPy-style broadcasting rules
  const rowsCompatible = dim1.rows === dim2.rows || dim1.rows === 1 || dim2.rows === 1;
  const colsCompatible = dim1.cols === dim2.cols || dim1.cols === 1 || dim2.cols === 1;
  
  return rowsCompatible && colsCompatible;
}

/**
 * Result of broadcast operation
 */
export interface BroadcastArrayResult {
  array1: CellValue[][];
  array2: CellValue[][];
  dimensions: ArrayDimensions;
}

/**
 * Broadcasts two arrays to compatible dimensions
 */
export function broadcast(
  array1: CellValue[][],
  array2: CellValue[][]
): BroadcastArrayResult | FormulaError {
  const dim1 = getArrayDimensions(array1);
  const dim2 = getArrayDimensions(array2);
  
  if (!areDimensionsCompatible(dim1, dim2)) {
    return '#VALUE!';
  }
  
  // Calculate broadcast dimensions
  const resultDims: ArrayDimensions = {
    rows: Math.max(dim1.rows, dim2.rows),
    cols: Math.max(dim1.cols, dim2.cols)
  };
  
  // Broadcast array1
  const broadcast1 = broadcastToSize(array1, resultDims);
  if (typeof broadcast1 === 'string' && isFormulaError(broadcast1)) return broadcast1;
  
  // Broadcast array2
  const broadcast2 = broadcastToSize(array2, resultDims);
  if (typeof broadcast2 === 'string' && isFormulaError(broadcast2)) return broadcast2;
  
  return {
    array1: broadcast1 as CellValue[][],
    array2: broadcast2 as CellValue[][],
    dimensions: resultDims
  };
}

/**
 * Broadcasts an array to specific dimensions
 */
export function broadcastToSize(
  array: CellValue[][],
  targetDims: ArrayDimensions
): CellValue[][] | FormulaError {
  const sourceDims = getArrayDimensions(array);
  
  // Handle empty arrays
  if (sourceDims.rows === 0 || sourceDims.cols === 0) {
    // Create array filled with undefined
    const result: CellValue[][] = [];
    for (let row = 0; row < targetDims.rows; row++) {
      const resultRow: CellValue[] = [];
      for (let col = 0; col < targetDims.cols; col++) {
        resultRow.push(undefined);
      }
      result.push(resultRow);
    }
    return result;
  }
  
  // Check if broadcasting is valid
  if (sourceDims.rows !== 1 && sourceDims.rows !== targetDims.rows) {
    return '#VALUE!';
  }
  if (sourceDims.cols !== 1 && sourceDims.cols !== targetDims.cols) {
    return '#VALUE!';
  }
  
  const result: CellValue[][] = [];
  
  for (let row = 0; row < targetDims.rows; row++) {
    const resultRow: CellValue[] = [];
    const sourceRow = sourceDims.rows === 1 ? 0 : row;
    
    for (let col = 0; col < targetDims.cols; col++) {
      const sourceCol = sourceDims.cols === 1 ? 0 : col;
      resultRow.push(array[sourceRow]?.[sourceCol] ?? undefined);
    }
    
    result.push(resultRow);
  }
  
  return result;
}

/**
 * Element-wise binary operation on arrays
 */
export function elementWiseBinaryOp(
  array1: CellValue[][],
  array2: CellValue[][],
  operation: (a: CellValue, b: CellValue) => CellValue
): CellValue[][] | FormulaError {
  // Broadcast arrays to compatible dimensions
  const broadcastResult = broadcast(array1, array2);
  if (typeof broadcastResult === 'string' && isFormulaError(broadcastResult)) return broadcastResult;
  
  const { array1: broadcast1, array2: broadcast2, dimensions } = broadcastResult;
  const result: CellValue[][] = [];
  
  for (let row = 0; row < dimensions.rows; row++) {
    const resultRow: CellValue[] = [];
    
    for (let col = 0; col < dimensions.cols; col++) {
      const value1 = broadcast1[row]?.[col];
      const value2 = broadcast2[row]?.[col];
      
      // Propagate errors
      if (isFormulaError(value1)) {
        resultRow.push(value1);
      } else if (isFormulaError(value2)) {
        resultRow.push(value2);
      } else {
        resultRow.push(operation(value1, value2));
      }
    }
    
    result.push(resultRow);
  }
  
  return result;
}

/**
 * Element-wise unary operation on array
 */
export function elementWiseUnaryOp(
  array: CellValue[][],
  operation: (value: CellValue) => CellValue
): CellValue[][] {
  return array.map(row =>
    row.map(value => {
      if (isFormulaError(value)) {
        return value;
      }
      return operation(value);
    })
  );
}

/**
 * Reduces an array along an axis
 */
export function reduceArray(
  array: CellValue[][],
  operation: (acc: CellValue, value: CellValue) => CellValue,
  initialValue: CellValue,
  axis?: 0 | 1 | null
): CellValue | CellValue[] | CellValue[][] {
  // Check for errors first
  const error = propagateError2D(array);
  if (error) return error;
  
  if (axis === null || axis === undefined) {
    // Reduce to scalar
    let result = initialValue;
    for (const row of array) {
      for (const value of row) {
        if (!isFormulaError(value)) {
          result = operation(result, value);
        }
      }
    }
    return result;
  }
  
  if (axis === 0) {
    // Reduce along rows (result is a row)
    const cols = array[0]?.length || 0;
    const result: CellValue[] = new Array(cols).fill(initialValue);
    
    for (let col = 0; col < cols; col++) {
      for (let row = 0; row < array.length; row++) {
        const value = array[row]?.[col];
        if (value !== undefined && !isFormulaError(value)) {
          result[col] = operation(result[col], value);
        }
      }
    }
    
    return [result]; // Return as 2D array with single row
  }
  
  if (axis === 1) {
    // Reduce along columns (result is a column)
    return array.map(row => {
      let result = initialValue;
      for (const value of row) {
        if (!isFormulaError(value)) {
          result = operation(result, value);
        }
      }
      return [result]; // Return as column
    });
  }
  
  return '#VALUE!';
}

/**
 * Filters an array based on a condition array
 */
export function filterArray(
  dataArray: CellValue[][],
  conditionArray: CellValue[][]
): CellValue[][] | FormulaError {
  // Broadcast arrays to same size
  const broadcastResult = broadcast(dataArray, conditionArray);
  if (typeof broadcastResult === 'string' && isFormulaError(broadcastResult)) return broadcastResult;
  
  const { array1: data, array2: conditions } = broadcastResult;
  const result: CellValue[] = [];
  
  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < (data[row]?.length || 0); col++) {
      const condition = conditions[row]?.[col];
      
      // Check for errors in condition
      if (isFormulaError(condition)) {
        return condition;
      }
      
      // Evaluate condition as boolean
      if (coerceToBoolean(condition)) {
        result.push(data[row]?.[col]);
      }
    }
  }
  
  // Return as column vector
  return result.length > 0 ? result.map(v => [v]) : [[undefined]];
}

/**
 * Constrains an array to specified dimensions
 */
export function constrainArray(
  array: CellValue[][],
  maxRows: number,
  maxCols: number
): CellValue[][] {
  const dims = getArrayDimensions(array);
  const resultRows = Math.min(dims.rows, maxRows);
  const resultCols = Math.min(dims.cols, maxCols);
  
  const result: CellValue[][] = [];
  
  for (let row = 0; row < resultRows; row++) {
    const resultRow: CellValue[] = [];
    for (let col = 0; col < resultCols; col++) {
      resultRow.push(array[row]?.[col] ?? undefined);
    }
    result.push(resultRow);
  }
  
  return result;
}

/**
 * Transposes a 2D array
 */
export function transpose(array: CellValue[][]): CellValue[][] {
  if (array.length === 0) return [];
  
  const rows = array.length;
  const cols = array[0]?.length || 0;
  const result: CellValue[][] = [];
  
  for (let col = 0; col < cols; col++) {
    const newRow: CellValue[] = [];
    for (let row = 0; row < rows; row++) {
      newRow.push(array[row]?.[col] ?? undefined);
    }
    result.push(newRow);
  }
  
  return result;
}

/**
 * Coerces a value to boolean for array conditions
 */
export function coerceToBoolean(value: CellValue): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    // Check if it's an error first
    if (isFormulaError(value)) return false;
    return value.length > 0;
  }
  if (value === undefined) return false;
  return false;
}

/**
 * Checks if an array operation will spill
 */
export function willSpill(
  targetRow: number,
  targetCol: number,
  dimensions: ArrayDimensions,
  maxRow: number,
  maxCol: number
): boolean {
  if (dimensions.rows === 1 && dimensions.cols === 1) {
    return false; // Single values don't spill
  }
  
  return (targetRow + dimensions.rows - 1 > maxRow) ||
         (targetCol + dimensions.cols - 1 > maxCol);
}

/**
 * Calculates spill range for an array formula
 */
export interface SpillRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function calculateSpillRange(
  targetRow: number,
  targetCol: number,
  array: CellValue[][]
): SpillRange {
  const dims = getArrayDimensions(array);
  
  return {
    startRow: targetRow,
    startCol: targetCol,
    endRow: targetRow + dims.rows - 1,
    endCol: targetCol + dims.cols - 1
  };
}

/**
 * Array operation cache for performance
 */
export class ArrayOperationCache {
  private cache: Map<string, CellValue[][]> = new Map();
  private maxSize: number;
  
  constructor(maxSize: number = 1000) {
    this.maxSize = maxSize;
  }
  
  /**
   * Creates a cache key from operation parameters
   */
  private createKey(operation: string, ...args: any[]): string {
    return `${operation}:${JSON.stringify(args)}`;
  }
  
  /**
   * Gets cached result
   */
  get(operation: string, ...args: any[]): CellValue[][] | undefined {
    const key = this.createKey(operation, ...args);
    return this.cache.get(key);
  }
  
  /**
   * Sets cached result
   */
  set(operation: string, result: CellValue[][], ...args: any[]): void {
    if (this.cache.size >= this.maxSize) {
      // Remove oldest entry (first in map)
      const firstKey = this.cache.keys().next().value;
      if (firstKey) {
        this.cache.delete(firstKey);
      }
    }
    
    const key = this.createKey(operation, ...args);
    this.cache.set(key, result);
  }
  
  /**
   * Clears the cache
   */
  clear(): void {
    this.cache.clear();
  }
}

/**
 * Optimized array sum operation
 */
export function arraySum(array: CellValue[][]): number {
  let sum = 0;
  let hasNumbers = false;
  
  for (const row of array) {
    for (const value of row) {
      if (typeof value === 'number') {
        sum += value;
        hasNumbers = true;
      } else if (typeof value === 'boolean') {
        sum += value ? 1 : 0;
        hasNumbers = true;
      }
      // Skip strings, undefined, and errors
    }
  }
  
  return hasNumbers ? sum : 0;
}

/**
 * Optimized array product operation
 */
export function arrayProduct(array: CellValue[][]): number {
  let product = 1;
  let hasNumbers = false;
  
  for (const row of array) {
    for (const value of row) {
      if (typeof value === 'number') {
        product *= value;
        hasNumbers = true;
      } else if (typeof value === 'boolean') {
        product *= value ? 1 : 0;
        hasNumbers = true;
      }
      // Skip strings, undefined, and errors
    }
  }
  
  return hasNumbers ? product : 1;
}

/**
 * Optimized array count operation
 */
export function arrayCount(array: CellValue[][]): number {
  let count = 0;
  
  for (const row of array) {
    for (const value of row) {
      if (typeof value === 'number' || typeof value === 'boolean') {
        count++;
      }
    }
  }
  
  return count;
}
