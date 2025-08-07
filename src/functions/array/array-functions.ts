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

// Helper to flatten a value into a 1D array
function flattenToArray(value: CellValue): CellValue[] {
  if (!Array.isArray(value)) {
    return [value];
  }
  
  // If it's a 2D array, flatten it
  if (Array.isArray(value[0])) {
    const result: CellValue[] = [];
    for (const row of value) {
      if (Array.isArray(row)) {
        result.push(...row);
      } else {
        result.push(row);
      }
    }
    return result;
  }
  
  return value;
}

// Helper to get dimensions of an array
function getArrayDimensions(value: CellValue): { rows: number; cols: number; isRowRange: boolean; isColumnRange: boolean } {
  if (!Array.isArray(value)) {
    return { rows: 1, cols: 1, isRowRange: false, isColumnRange: false };
  }
  
  if (Array.isArray(value[0])) {
    // 2D array
    const rows = value.length;
    const cols = Math.max(...value.map(row => Array.isArray(row) ? row.length : 1));
    
    // Row range: 1 row, multiple columns
    const isRowRange = rows === 1 && cols > 1;
    
    // Column range: multiple rows, 1 column each
    const isColumnRange = rows > 1 && value.every(row => Array.isArray(row) && row.length === 1);
    
    return { rows, cols, isRowRange, isColumnRange };
  }
  
  // 1D array - treat as column range
  return { rows: value.length, cols: 1, isRowRange: false, isColumnRange: true };
}

// Helper function to filter row ranges (columns within a row)
function filterRowRange(sourceArray: CellValue[][], conditionArrays: CellValue[]): CellValue {
  const sourceRow = sourceArray[0]; // Single row with multiple columns
  if (!sourceRow) {
    throw new Error('#VALUE!');
  }
  
  const result: CellValue[] = [];
  
  // Validate all condition arrays are also row ranges with same column count
  for (const condArray of conditionArrays) {
    if (!Array.isArray(condArray) || condArray.length !== 1 || 
        !Array.isArray(condArray[0]) || condArray[0].length !== sourceRow.length) {
      throw new Error('#VALUE!');
    }
  }
  
  // Filter columns based on conditions
  for (let colIndex = 0; colIndex < sourceRow.length; colIndex++) {
    let includeColumn = true;
    
    // Check all conditions for this column
    for (const condArray of conditionArrays) {
      if (Array.isArray(condArray) && Array.isArray(condArray[0])) {
        const conditionRow = condArray[0] as CellValue[];
        if (!coerceToBoolean(conditionRow[colIndex])) {
          includeColumn = false;
          break;
        }
      }
    }
    
    if (includeColumn) {
      result.push(sourceRow[colIndex]);
    }
  }
  
  // If no results, return #N/A
  if (result.length === 0) {
    return '#N/A';
  }
  
  // Return as column vector (standard FILTER output format)
  return result.map(value => [value]) as unknown as CellValue;
}

// FILTER(sourceArray, ...boolArrays)
export const FILTER: FunctionDefinition = {
  name: 'FILTER',
  returnsArray: true,
  evaluate: (args: CellValue[], context: EvaluationContext): CellValue => {
    if (args.length < 2) {
      throw new Error('#VALUE!');
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    const sourceArray = args[0];
    
    // Handle non-array source
    if (!Array.isArray(sourceArray)) {
      // Single value - check if all conditions are true
      for (let i = 1; i < args.length; i++) {
        const condition = args[i];
        if (!coerceToBoolean(Array.isArray(condition) ? condition[0] : condition)) {
          return '#N/A'; // No matching values
        }
      }
      return [[sourceArray]] as unknown as CellValue; // Return as 2D array
    }

    // Get dimensions of source array
    const sourceDims = getArrayDimensions(sourceArray);
    const is2D = Array.isArray(sourceArray[0]);
    
    // Detect if this is a row range (1 row, multiple columns)
    if (sourceDims.isRowRange) {
      return filterRowRange(sourceArray as CellValue[][], args.slice(1));
    }
    
    // Validate all condition arrays have compatible dimensions
    for (let i = 1; i < args.length; i++) {
      const condArray = args[i];
      const condDims = getArrayDimensions(condArray);
      
      // Condition must have same number of rows
      if (condDims.rows !== sourceDims.rows) {
        throw new Error('#VALUE!');
      }
      
      // For 2D source arrays, condition can be either 1D (applied row-wise) or 2D (exact match)
      if (is2D && condDims.cols !== 1 && condDims.cols !== sourceDims.cols) {
        throw new Error('#VALUE!');
      }
    }

    // Filter the array
    const result: CellValue[][] = [];
    
    if (is2D) {
      // 2D array filtering
      for (let r = 0; r < sourceDims.rows; r++) {
        let includeRow = true;
        
        // Check all conditions for this row
        for (let i = 1; i < args.length; i++) {
          const condArray = args[i];
          const condDims = getArrayDimensions(condArray);
          
          if (condDims.cols === 1) {
            // 1D condition array - check the single value for this row
            const condValue = Array.isArray(condArray) ? 
              (Array.isArray(condArray[r]) ? condArray[r][0] : condArray[r]) : 
              condArray;
            if (!coerceToBoolean(condValue)) {
              includeRow = false;
              break;
            }
          } else {
            // 2D condition array - check if any value in this row is true
            const condRow = (condArray as unknown as CellValue[][])[r];
            let rowHasTrue = false;
            for (let c = 0; c < condDims.cols; c++) {
              const condValue = condRow?.[c];
              if (coerceToBoolean(condValue)) {
                rowHasTrue = true;
                break;
              }
            }
            if (!rowHasTrue) {
              includeRow = false;
              break;
            }
          }
        }
        
        if (includeRow) {
          result.push(sourceArray[r] as CellValue[]);
        }
      }
    } else {
      // 1D array filtering
      const sourceFlat = sourceArray as CellValue[];
      const tempResult: CellValue[] = [];
      
      for (let i = 0; i < sourceFlat.length; i++) {
        let includeValue = true;
        
        // Check all conditions for this value
        for (let j = 1; j < args.length; j++) {
          const condArray = flattenToArray(args[j]);
          if (!coerceToBoolean(condArray[i])) {
            includeValue = false;
            break;
          }
        }
        
        if (includeValue) {
          tempResult.push(sourceFlat[i]);
        }
      }
      
      // Convert to 2D array (single column)
      tempResult.forEach(val => result.push([val]));
    }

    // If no results, return #N/A
    if (result.length === 0) {
      return '#N/A';
    }

    return result as unknown as CellValue;
  }
};

// ARRAY_CONSTRAIN(array, height, width)
export const ARRAY_CONSTRAIN: FunctionDefinition = {
  name: 'ARRAY_CONSTRAIN',
  returnsArray: true,
  evaluate: (args: CellValue[], context: EvaluationContext): CellValue => {
    if (args.length !== 3) {
      throw new Error('#VALUE!');
    }

    // Check for errors
    const error = propagateError(args);
    if (error) return error;

    const array = args[0];
    const height = args[1];
    const width = args[2];

    // Validate height and width are positive integers
    if (typeof height !== 'number' || height < 1 || !Number.isInteger(height)) {
      throw new Error('#VALUE!');
    }
    if (typeof width !== 'number' || width < 1 || !Number.isInteger(width)) {
      throw new Error('#VALUE!');
    }

    // Handle non-array input
    if (!Array.isArray(array)) {
      return [[array]] as unknown as CellValue; // Return as 1x1 array
    }

    // Get dimensions
    const is2D = Array.isArray(array[0]);
    const result: CellValue[][] = [];

    if (is2D) {
      // Constrain 2D array
      const sourceArray = array as CellValue[][];
      for (let r = 0; r < Math.min(height, sourceArray.length); r++) {
        const row: CellValue[] = [];
        const sourceRow = sourceArray[r];
        
        if (Array.isArray(sourceRow)) {
          for (let c = 0; c < Math.min(width, sourceRow.length); c++) {
            row.push(sourceRow[c]);
          }
        } else {
          row.push(sourceRow);
        }
        
        // Pad with undefined if needed
        while (row.length < width) {
          row.push(undefined);
        }
        
        result.push(row);
      }
    } else {
      // Convert 1D to 2D and constrain
      const sourceArray = array as CellValue[];
      let idx = 0;
      
      for (let r = 0; r < height && idx < sourceArray.length; r++) {
        const row: CellValue[] = [];
        for (let c = 0; c < width && idx < sourceArray.length; c++) {
          row.push(sourceArray[idx++]);
        }
        
        // Pad with undefined if needed
        while (row.length < width) {
          row.push(undefined);
        }
        
        result.push(row);
      }
    }

    // Pad with empty rows if needed
    while (result.length < height) {
      const emptyRow: CellValue[] = [];
      for (let c = 0; c < width; c++) {
        emptyRow.push(undefined);
      }
      result.push(emptyRow);
    }

    return result as unknown as CellValue;
  }
};

// Export all array functions
export const arrayFunctions: FunctionDefinition[] = [
  FILTER,
  ARRAY_CONSTRAIN
];