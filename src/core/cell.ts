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
 * Convert a cell to a serializable format
 */
export function serializeCell(cell: Cell | undefined, address?: SimpleCellAddress): CellValue | string {
  if (!cell) {
    return undefined;
  }

  // If it's an array cell but not the origin, return undefined (empty)
  // This ensures spilled cells are empty in serialized format (Excel compatibility)
  if (cell.type === 'ARRAY' && cell.arrayFormula && address) {
    const isOrigin = cell.arrayFormula.originAddress.sheet === address.sheet &&
                     cell.arrayFormula.originAddress.row === address.row &&
                     cell.arrayFormula.originAddress.col === address.col;
    
    if (!isOrigin) {
      // This is a spilled cell, should be empty in serialized format
      return undefined;
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








