/**
 * Sheet data structure and operations
 * Manages cells within a sheet using sparse storage
 */

import type {
  Sheet,
  Cell,
  SimpleCellAddress,
  SimpleCellRange,
  BoundingRect,
  CellValue,
  SheetDimensions
} from './types';

import {
  addressToKey
} from './types';
import { addressToA1, isAddressInRange, iterateRange } from './address';
import { createEmptyCell, isEmptyCell } from './cell';

/**
 * Create a new empty sheet
 */
export function createSheet(id: number, name: string): Sheet {
  return {
    id,
    name,
    cells: new Map<string, Cell>(),
    dimensions: undefined
  };
}

/**
 * Get a cell from a sheet
 */
export function getCell(sheet: Sheet, address: SimpleCellAddress): Cell | undefined {
  const key = addressToA1(address, { absolute: false });
  return sheet.cells.get(key);
}

/**
 * Set a cell in a sheet
 */
export function setCell(sheet: Sheet, address: SimpleCellAddress, cell: Cell): void {
  const key = addressToA1(address, { absolute: false });
  
  if (isEmptyCell(cell)) {
    // Remove empty cells to maintain sparsity
    sheet.cells.delete(key);
  } else {
    sheet.cells.set(key, cell);
  }
  
  // Update dimensions
  updateSheetDimensions(sheet);
}

/**
 * Remove a cell from a sheet
 */
export function removeCell(sheet: Sheet, address: SimpleCellAddress): void {
  const key = addressToA1(address, { absolute: false });
  sheet.cells.delete(key);
  updateSheetDimensions(sheet);
}

/**
 * Clear all cells in a sheet
 */
export function clearSheet(sheet: Sheet): void {
  sheet.cells.clear();
  sheet.dimensions = undefined;
}

/**
 * Get all cells in a range
 */
export function getCellsInRange(sheet: Sheet, range: SimpleCellRange): Map<string, Cell> {
  const result = new Map<string, Cell>();
  
  // For small ranges, iterate through all possible cells
  const rangeSize = (range.end.col - range.start.col + 1) * (range.end.row - range.start.row + 1);
  const sheetSize = sheet.cells.size;
  
  if (rangeSize < sheetSize) {
    // Iterate through range
    for (const address of iterateRange(range)) {
      const cell = getCell(sheet, address);
      if (cell && !isEmptyCell(cell)) {
        const key = addressToA1(address, { absolute: false });
        result.set(key, cell);
      }
    }
  } else {
    // Iterate through sheet cells and filter
    for (const [key, cell] of sheet.cells) {
      const address = parseA1Key(key, sheet.id);
      if (address && isAddressInRange(address, range)) {
        result.set(key, cell);
      }
    }
  }
  
  return result;
}

/**
 * Get values from a range as a 2D array
 */
export function getRangeValues(sheet: Sheet, range: SimpleCellRange): CellValue[][] {
  const width = range.end.col - range.start.col + 1;
  const height = range.end.row - range.start.row + 1;
  const result: CellValue[][] = [];
  
  for (let row = 0; row < height; row++) {
    const rowData: CellValue[] = [];
    for (let col = 0; col < width; col++) {
      const address: SimpleCellAddress = {
        sheet: sheet.id,
        col: range.start.col + col,
        row: range.start.row + row
      };
      const cell = getCell(sheet, address);
      rowData.push(cell?.value);
    }
    result.push(rowData);
  }
  
  return result;
}

/**
 * Set values in a range from a 2D array
 */
export function setRangeValues(
  sheet: Sheet,
  topLeft: SimpleCellAddress,
  values: CellValue[][]
): void {
  for (let row = 0; row < values.length; row++) {
    const rowData = values[row];
    for (let col = 0; rowData && col < rowData.length; col++) {
      const address: SimpleCellAddress = {
        sheet: sheet.id,
        col: topLeft.col + col,
        row: topLeft.row + row
      };
      
      const value = rowData![col];
      if (value === undefined) {
        removeCell(sheet, address);
      } else {
        setCell(sheet, address, {
          value,
          type: 'VALUE',
          formula: undefined,
          dependencies: undefined,
          dependents: undefined,
          arrayFormula: undefined
        });
      }
    }
  }
}

/**
 * Get the bounding rectangle of populated cells
 */
export function getBoundingRect(sheet: Sheet): BoundingRect | undefined {
  if (sheet.cells.size === 0) {
    return undefined;
  }
  
  let minCol = Infinity;
  let maxCol = -Infinity;
  let minRow = Infinity;
  let maxRow = -Infinity;
  
  for (const key of sheet.cells.keys()) {
    const address = parseA1Key(key, sheet.id);
    if (address) {
      minCol = Math.min(minCol, address.col);
      maxCol = Math.max(maxCol, address.col);
      minRow = Math.min(minRow, address.row);
      maxRow = Math.max(maxRow, address.row);
    }
  }
  
  if (!isFinite(minCol) || !isFinite(maxCol) || !isFinite(minRow) || !isFinite(maxRow)) {
    return undefined;
  }
  
  return {
    minCol,
    maxCol,
    minRow,
    maxRow,
    width: maxCol - minCol + 1,
    height: maxRow - minRow + 1
  };
}

/**
 * Update sheet dimensions after cell changes
 */
function updateSheetDimensions(sheet: Sheet): void {
  const bounds = getBoundingRect(sheet);
  if (bounds) {
    sheet.dimensions = {
      minRow: bounds.minRow,
      maxRow: bounds.maxRow,
      minCol: bounds.minCol,
      maxCol: bounds.maxCol
    };
  } else {
    sheet.dimensions = undefined;
  }
}

/**
 * Parse an A1 key back to an address
 * This is a simplified version - assumes keys are in format "A1", "B2", etc.
 */
function parseA1Key(key: string, sheetId: number): SimpleCellAddress | null {
  const match = key.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return null;
  }
  
  const [, colLetters, rowDigits] = match;
  if (!colLetters || !rowDigits) {
    throw new Error(`Invalid cell address: ${key}`);
  }
  const col = letterToColNumber(colLetters);
  const row = parseInt(rowDigits, 10) - 1;
  
  return {
    sheet: sheetId,
    col,
    row
  };
}

/**
 * Convert column letters to number (A=0, B=1, Z=25, AA=26, etc.)
 */
function letterToColNumber(letters: string): number {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result - 1;
}

/**
 * Count non-empty cells in a sheet
 */
export function countNonEmptyCells(sheet: Sheet): number {
  return sheet.cells.size;
}

/**
 * Get all non-empty cells as a Map
 */
export function getAllCells(sheet: Sheet): Map<string, Cell> {
  return new Map(sheet.cells);
}

/**
 * Copy cells from one range to another within the same sheet
 */
export function copyCellsInSheet(
  sheet: Sheet,
  sourceRange: SimpleCellRange,
  targetTopLeft: SimpleCellAddress
): void {
  const cells = getCellsInRange(sheet, sourceRange);
  const offsetCol = targetTopLeft.col - sourceRange.start.col;
  const offsetRow = targetTopLeft.row - sourceRange.start.row;
  
  // First collect all the cells to copy (to avoid modifying while iterating)
  const cellsToCopy: Array<[SimpleCellAddress, Cell]> = [];
  
  for (const [key, cell] of cells) {
    const sourceAddr = parseA1Key(key, sheet.id);
    if (sourceAddr) {
      const targetAddr: SimpleCellAddress = {
        sheet: sheet.id,
        col: sourceAddr.col + offsetCol,
        row: sourceAddr.row + offsetRow
      };
      cellsToCopy.push([targetAddr, cell]);
    }
  }
  
  // Then apply all the copies
  for (const [addr, cell] of cellsToCopy) {
    setCell(sheet, addr, cell);
  }
}

/**
 * Remove cells in a range
 */
export function removeCellsInRange(sheet: Sheet, range: SimpleCellRange): void {
  const cellsToRemove = getCellsInRange(sheet, range);
  
  for (const key of cellsToRemove.keys()) {
    sheet.cells.delete(key);
  }
  
  updateSheetDimensions(sheet);
}

/**
 * Shift cells when inserting/deleting rows or columns
 */
export function shiftCells(
  sheet: Sheet,
  dimension: 'row' | 'col',
  startIndex: number,
  offset: number
): void {
  if (offset === 0) return;
  
  const cellsToMove: Array<[string, SimpleCellAddress, Cell]> = [];
  const cellsToRemove: string[] = [];
  
  // Collect cells that need to be moved
  for (const [key, cell] of sheet.cells) {
    const address = parseA1Key(key, sheet.id);
    if (!address) continue;
    
    const index = dimension === 'row' ? address.row : address.col;
    
    if (index >= startIndex) {
      const newAddress: SimpleCellAddress = {
        sheet: sheet.id,
        col: dimension === 'col' ? address.col + offset : address.col,
        row: dimension === 'row' ? address.row + offset : address.row
      };
      
      // Check if the new position is valid
      if (newAddress.col >= 0 && newAddress.row >= 0) {
        const newKey = addressToA1(newAddress, { absolute: false });
        cellsToMove.push([newKey, newAddress, cell]);
        cellsToRemove.push(key);
      } else if (offset < 0) {
        // Cell is being shifted out of bounds (deleted)
        cellsToRemove.push(key);
      }
    }
  }
  
  // Remove old cells
  for (const key of cellsToRemove) {
    sheet.cells.delete(key);
  }
  
  // Add cells at new positions
  for (const [key, , cell] of cellsToMove) {
    sheet.cells.set(key, cell);
  }
  
  updateSheetDimensions(sheet);
}

/**
 * Check if a sheet has any formulas
 */
export function hasFormulas(sheet: Sheet): boolean {
  for (const cell of sheet.cells.values()) {
    if (cell.formula) {
      return true;
    }
  }
  return false;
}

/**
 * Get all cells with formulas
 */
export function getFormulaCells(sheet: Sheet): Map<string, Cell> {
  const result = new Map<string, Cell>();
  
  for (const [key, cell] of sheet.cells) {
    if (cell.formula) {
      result.set(key, cell);
    }
  }
  
  return result;
}
