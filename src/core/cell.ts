/**
 * Cell data structures and utilities
 * Manages cell values, formulas, and types
 */

import type {
  Cell,
  CellValue,
  CellType,
  CellValueType,
  SimpleCellAddress,
  ArrayFormula
} from './types';

import {
  isFormulaError,
  isCellEmpty,
  getCellValueType,
  isNumber,
  isString,
  isBoolean
} from './types';

/**
 * Create a new cell with a value
 */
export function createValueCell(value: CellValue): Cell {
  return {
    value,
    type: 'VALUE',
    formula: undefined,
    dependencies: undefined,
    dependents: undefined,
    arrayFormula: undefined
  };
}

/**
 * Create a new cell with a formula
 */
export function createFormulaCell(formula: string, value: CellValue = undefined): Cell {
  return {
    value,
    type: 'FORMULA',
    formula,
    dependencies: new Set<string>(),
    dependents: new Set<string>(),
    arrayFormula: undefined
  };
}

/**
 * Create a new array formula cell
 */
export function createArrayCell(
  formula: string,
  value: CellValue,
  arrayFormula: ArrayFormula
): Cell {
  return {
    value,
    type: 'ARRAY',
    formula,
    dependencies: new Set<string>(),
    dependents: new Set<string>(),
    arrayFormula
  };
}

/**
 * Create an empty cell
 */
export function createEmptyCell(): Cell {
  return {
    value: undefined,
    type: 'EMPTY',
    formula: undefined,
    dependencies: undefined,
    dependents: undefined,
    arrayFormula: undefined
  };
}

/**
 * Check if a cell has a formula
 */
export function hasFormula(cell: Cell | undefined): boolean {
  return cell !== undefined && cell.type === 'FORMULA' && cell.formula !== undefined;
}

/**
 * Check if a cell is part of an array formula
 */
export function isArrayCell(cell: Cell | undefined): boolean {
  return cell !== undefined && cell.type === 'ARRAY' && cell.arrayFormula !== undefined;
}

/**
 * Check if a cell is empty
 */
export function isEmptyCell(cell: Cell | undefined): boolean {
  return cell === undefined || (cell.type === 'EMPTY' && cell.value === undefined);
}

/**
 * Get the display value of a cell
 */
export function getCellDisplayValue(cell: Cell | undefined): string {
  if (!cell || cell.value === undefined) {
    return '';
  }

  if (isFormulaError(cell.value)) {
    return cell.value;
  }

  if (isBoolean(cell.value)) {
    return cell.value ? 'TRUE' : 'FALSE';
  }

  if (isNumber(cell.value)) {
    // Handle special number cases
    if (!isFinite(cell.value)) {
      return '#NUM!';
    }
    return cell.value.toString();
  }

  if (isString(cell.value)) {
    return cell.value;
  }

  return '';
}

/**
 * Parse a value from user input
 */
export function parseInputValue(input: any): CellValue {
  // Handle null/undefined
  if (input === null || input === undefined || input === '') {
    return undefined;
  }

  // Already a valid cell value
  if (typeof input === 'number' || typeof input === 'boolean') {
    return input;
  }

  // String input needs parsing
  if (typeof input === 'string') {
    const trimmed = input.trim();

    // Empty string
    if (trimmed === '') {
      return undefined;
    }

    // Boolean values
    if (trimmed.toUpperCase() === 'TRUE') {
      return true;
    }
    if (trimmed.toUpperCase() === 'FALSE') {
      return false;
    }

    // Formula error
    if (isFormulaError(trimmed)) {
      return trimmed as CellValue;
    }

    // Try to parse as number
    if (/^[+-]?\d*\.?\d+([eE][+-]?\d+)?$/.test(trimmed)) {
      const num = parseFloat(trimmed);
      if (!isNaN(num) && isFinite(num)) {
        return num;
      }
    }

    // Return as string
    return trimmed;
  }

  // Convert other types to string
  return String(input);
}

/**
 * Check if input is a formula (starts with =)
 */
export function isFormula(input: string): boolean {
  return typeof input === 'string' && input.trim().startsWith('=');
}

/**
 * Extract formula from input (remove leading =)
 */
export function extractFormula(input: string): string {
  if (!isFormula(input)) {
    return '';
  }
  return input.trim().substring(1);
}

/**
 * Clone a cell (deep copy)
 */
export function cloneCell(cell: Cell): Cell {
  return {
    value: cell.value,
    type: cell.type,
    formula: cell.formula,
    dependencies: cell.dependencies ? new Set(cell.dependencies) : undefined,
    dependents: cell.dependents ? new Set(cell.dependents) : undefined,
    arrayFormula: cell.arrayFormula ? { ...cell.arrayFormula } : undefined
  };
}

/**
 * Convert a cell to a serializable format
 */
export function serializeCell(cell: Cell | undefined, address?: SimpleCellAddress): CellValue | string {
  if (!cell) {
    return undefined;
  }

  // If it's an array cell but not the origin, return the value
  // This ensures spilled cells copy their values, not the formula
  if (cell.type === 'ARRAY' && cell.arrayFormula && address) {
    const isOrigin = cell.arrayFormula.originAddress.sheet === address.sheet &&
                     cell.arrayFormula.originAddress.row === address.row &&
                     cell.arrayFormula.originAddress.col === address.col;
    
    if (!isOrigin) {
      // This is a spilled cell, return the value
      return cell.value;
    }
  }

  // If it has a formula, return the formula with = prefix
  if (cell.formula) {
    return `=${cell.formula}`;
  }

  // Otherwise return the value
  return cell.value;
}

/**
 * Create a cell from serialized content
 */
export function deserializeCell(content: CellValue | string): Cell {
  // Handle formulas
  if (typeof content === 'string' && isFormula(content)) {
    return createFormulaCell(extractFormula(content));
  }

  // Handle regular values
  const value = parseInputValue(content);
  if (value === undefined) {
    return createEmptyCell();
  }

  return createValueCell(value);
}

/**
 * Compare two cell values for equality
 */
export function areCellValuesEqual(a: CellValue, b: CellValue): boolean {
  // Both undefined
  if (a === undefined && b === undefined) {
    return true;
  }

  // One undefined
  if (a === undefined || b === undefined) {
    return false;
  }

  // Direct equality (handles numbers, strings, booleans, errors)
  return a === b;
}

/**
 * Get the numeric value of a cell (for calculations)
 */
export function getNumericValue(value: CellValue): number | null {
  if (isNumber(value)) {
    return value;
  }

  if (isBoolean(value)) {
    return value ? 1 : 0;
  }

  if (isString(value)) {
    // Try to parse as number
    const trimmed = value.trim();
    if (trimmed === '') {
      return 0;
    }
    const num = parseFloat(trimmed);
    if (!isNaN(num) && isFinite(num)) {
      return num;
    }
  }

  return null;
}

/**
 * Coerce a value to boolean (for logical operations)
 */
export function coerceToBoolean(value: CellValue): boolean {
  if (isBoolean(value)) {
    return value;
  }

  if (isNumber(value)) {
    return value !== 0;
  }

  if (isString(value)) {
    const upper = value.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    
    // Non-empty strings are truthy
    return value.length > 0;
  }

  // undefined and errors are falsy
  return false;
}

/**
 * Coerce a value to string (for text operations)
 */
export function coerceToString(value: CellValue): string {
  if (value === undefined) {
    return '';
  }

  if (isString(value)) {
    return value;
  }

  if (isNumber(value)) {
    return value.toString();
  }

  if (isBoolean(value)) {
    return value ? 'TRUE' : 'FALSE';
  }

  if (isFormulaError(value)) {
    return value;
  }

  return '';
}

/**
 * Get cell metadata for debugging/inspection
 */
export interface CellMetadata {
  type: CellType;
  valueType: CellValueType;
  hasFormula: boolean;
  hasDependencies: boolean;
  hasDependents: boolean;
  isArrayFormula: boolean;
  dependencyCount: number;
  dependentCount: number;
}

export function getCellMetadata(cell: Cell | undefined): CellMetadata | null {
  if (!cell) {
    return null;
  }

  return {
    type: cell.type,
    valueType: getCellValueType(cell.value),
    hasFormula: !!cell.formula,
    hasDependencies: !!cell.dependencies && cell.dependencies.size > 0,
    hasDependents: !!cell.dependents && cell.dependents.size > 0,
    isArrayFormula: !!cell.arrayFormula,
    dependencyCount: cell.dependencies ? cell.dependencies.size : 0,
    dependentCount: cell.dependents ? cell.dependents.size : 0
  };
}
