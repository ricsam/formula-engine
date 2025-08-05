/**
 * Cell addressing utilities
 * Handles conversion between different address formats and validation
 */

import type {
  SimpleCellAddress,
  SimpleCellRange
} from './types';

import {
  colNumberToLetter,
  letterToColNumber
} from './types';

// Regular expressions for parsing addresses
const A1_CELL_REGEX = /^(\$?)([A-Z]+)(\$?)(\d+)$/i;
const A1_RANGE_REGEX = /^(\$?)([A-Z]+)(\$?)(\d+):(\$?)([A-Z]+)(\$?)(\d+)$/i;
const SHEET_NAME_REGEX = /^'?([^'!]+)'?!(.+)$/;

/**
 * Parse an A1-style cell address (e.g., "A1", "$B$2", "Sheet1!C3")
 */
export function parseCellAddress(address: string, contextSheetId: number): SimpleCellAddress | null {
  if (!address || typeof address !== 'string') {
    return null;
  }

  let sheetId = contextSheetId;
  let cellPart = address;

  // Check for sheet reference
  const sheetMatch = address.match(SHEET_NAME_REGEX);
  if (sheetMatch && sheetMatch[2]) {
    // For now, we'll need to handle sheet name resolution elsewhere
    // This function will be updated when we have sheet management
    cellPart = sheetMatch[2];
  }

  // Parse the cell reference
  const match = cellPart.match(A1_CELL_REGEX);
  if (!match) {
    return null;
  }

  const [, colAbsolute, colLetters, rowAbsolute, rowDigits] = match;
  if (!colLetters || !rowDigits) {
    return null;
  }
  const col = letterToColNumber(colLetters.toUpperCase());
  const row = parseInt(rowDigits, 10) - 1; // Convert to 0-based

  if (col < 0 || row < 0) {
    return null;
  }

  return {
    sheet: sheetId,
    col,
    row
  };
}

/**
 * Parse an A1-style range (e.g., "A1:B2", "$A$1:$B$2")
 */
export function parseCellRange(range: string, contextSheetId: number): SimpleCellRange | null {
  if (!range || typeof range !== 'string') {
    return null;
  }

  let sheetId = contextSheetId;
  let rangePart = range;

  // Check for sheet reference
  const sheetMatch = range.match(SHEET_NAME_REGEX);
  if (sheetMatch && sheetMatch[2]) {
    rangePart = sheetMatch[2];
  }

  // Check if it's a single cell reference
  const singleCell = parseCellAddress(rangePart, sheetId);
  if (singleCell) {
    return {
      start: singleCell,
      end: singleCell
    };
  }

  // Parse as range
  const match = rangePart.match(A1_RANGE_REGEX);
  if (!match) {
    return null;
  }

  const [, startColAbs, startColLetters, startRowAbs, startRowDigits,
         endColAbs, endColLetters, endRowAbs, endRowDigits] = match;

  if (!startColLetters || !startRowDigits || !endColLetters || !endRowDigits) {
    return null;
  }

  const startCol = letterToColNumber(startColLetters.toUpperCase());
  const startRow = parseInt(startRowDigits, 10) - 1;
  const endCol = letterToColNumber(endColLetters.toUpperCase());
  const endRow = parseInt(endRowDigits, 10) - 1;

  if (startCol < 0 || startRow < 0 || endCol < 0 || endRow < 0) {
    return null;
  }

  // Normalize range (ensure start is before end)
  return {
    start: {
      sheet: sheetId,
      col: Math.min(startCol, endCol),
      row: Math.min(startRow, endRow)
    },
    end: {
      sheet: sheetId,
      col: Math.max(startCol, endCol),
      row: Math.max(startRow, endRow)
    }
  };
}

/**
 * Convert a SimpleCellAddress to A1 notation
 */
export function addressToA1(
  address: SimpleCellAddress,
  options?: {
    includeSheetName?: boolean;
    sheetName?: string;
    absolute?: boolean;
  }
): string {
  const colLetter = colNumberToLetter(address.col);
  const rowNumber = address.row + 1; // Convert to 1-based

  let result = '';
  
  if (options?.includeSheetName && options?.sheetName) {
    // Quote sheet name if it contains spaces or special characters
    const needsQuotes = /[^A-Za-z0-9_]/.test(options.sheetName);
    if (needsQuotes) {
      result = `'${options.sheetName}'!`;
    } else {
      result = `${options.sheetName}!`;
    }
  }

  if (options?.absolute) {
    result += `$${colLetter}$${rowNumber}`;
  } else {
    result += `${colLetter}${rowNumber}`;
  }

  return result;
}

/**
 * Convert a SimpleCellRange to A1 notation
 */
export function rangeToA1(
  range: SimpleCellRange,
  options?: {
    includeSheetName?: boolean;
    sheetName?: string;
    absolute?: boolean;
  }
): string {
  const start = addressToA1(range.start, { ...options, includeSheetName: false });
  const end = addressToA1(range.end, { ...options, includeSheetName: false });

  let result = '';

  if (options?.includeSheetName && options?.sheetName) {
    const needsQuotes = /[^A-Za-z0-9_]/.test(options.sheetName);
    if (needsQuotes) {
      result = `'${options.sheetName}'!`;
    } else {
      result = `${options.sheetName}!`;
    }
  }

  // If start and end are the same, return single cell
  if (start === end) {
    result += start;
  } else {
    result += `${start}:${end}`;
  }

  return result;
}

/**
 * Check if an address is within a range
 */
export function isAddressInRange(address: SimpleCellAddress, range: SimpleCellRange): boolean {
  return address.sheet === range.start.sheet &&
         address.col >= range.start.col &&
         address.col <= range.end.col &&
         address.row >= range.start.row &&
         address.row <= range.end.row;
}

/**
 * Check if two ranges overlap
 */
export function doRangesOverlap(range1: SimpleCellRange, range2: SimpleCellRange): boolean {
  if (range1.start.sheet !== range2.start.sheet) {
    return false;
  }

  return !(range1.end.col < range2.start.col ||
           range2.end.col < range1.start.col ||
           range1.end.row < range2.start.row ||
           range2.end.row < range1.start.row);
}

/**
 * Calculate the size of a range
 */
export function getRangeSize(range: SimpleCellRange): { width: number; height: number; cells: number } {
  const width = range.end.col - range.start.col + 1;
  const height = range.end.row - range.start.row + 1;
  return {
    width,
    height,
    cells: width * height
  };
}

/**
 * Offset an address by a given amount
 */
export function offsetAddress(
  address: SimpleCellAddress,
  colOffset: number,
  rowOffset: number
): SimpleCellAddress {
  return {
    sheet: address.sheet,
    col: Math.max(0, address.col + colOffset),
    row: Math.max(0, address.row + rowOffset)
  };
}

/**
 * Offset a range by a given amount
 */
export function offsetRange(
  range: SimpleCellRange,
  colOffset: number,
  rowOffset: number
): SimpleCellRange {
  return {
    start: offsetAddress(range.start, colOffset, rowOffset),
    end: offsetAddress(range.end, colOffset, rowOffset)
  };
}

/**
 * Expand a range to include an address
 */
export function expandRangeToInclude(
  range: SimpleCellRange,
  address: SimpleCellAddress
): SimpleCellRange {
  if (address.sheet !== range.start.sheet) {
    throw new Error('Cannot expand range across sheets');
  }

  return {
    start: {
      sheet: range.start.sheet,
      col: Math.min(range.start.col, address.col),
      row: Math.min(range.start.row, address.row)
    },
    end: {
      sheet: range.end.sheet,
      col: Math.max(range.end.col, address.col),
      row: Math.max(range.end.row, address.row)
    }
  };
}

/**
 * Get all addresses in a range (generator for memory efficiency)
 */
export function* iterateRange(range: SimpleCellRange): Generator<SimpleCellAddress> {
  for (let row = range.start.row; row <= range.end.row; row++) {
    for (let col = range.start.col; col <= range.end.col; col++) {
      yield {
        sheet: range.start.sheet,
        col,
        row
      };
    }
  }
}

/**
 * Convert relative references in a formula when copying
 */
export function adjustReferences(
  formula: string,
  sourceAddress: SimpleCellAddress,
  targetAddress: SimpleCellAddress
): string {
  const colOffset = targetAddress.col - sourceAddress.col;
  const rowOffset = targetAddress.row - sourceAddress.row;

  // This is a simplified version - full implementation would use the parser
  return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/gi, (match, colAbs, colLetters, rowAbs, rowDigits) => {
    if (colAbs && rowAbs) {
      // Absolute reference, don't change
      return match;
    }

    let newCol = letterToColNumber(colLetters);
    let newRow = parseInt(rowDigits, 10) - 1;

    if (!colAbs) {
      newCol += colOffset;
    }
    if (!rowAbs) {
      newRow += rowOffset;
    }

    if (newCol < 0 || newRow < 0) {
      return '#REF!';
    }

    const newColLetter = colNumberToLetter(newCol);
    const newRowNumber = newRow + 1;

    return `${colAbs}${newColLetter}${rowAbs}${newRowNumber}`;
  });
}

/**
 * Validate that an address is within reasonable bounds
 */
export function isValidAddress(address: SimpleCellAddress): boolean {
  // Excel limits: 16,384 columns (XFD) and 1,048,576 rows
  const MAX_COLS = 16384;
  const MAX_ROWS = 1048576;

  return address.col >= 0 &&
         address.col < MAX_COLS &&
         address.row >= 0 &&
         address.row < MAX_ROWS &&
         Number.isInteger(address.col) &&
         Number.isInteger(address.row) &&
         Number.isInteger(address.sheet);
}

/**
 * Create a range from two addresses
 */
export function createRange(addr1: SimpleCellAddress, addr2: SimpleCellAddress): SimpleCellRange {
  if (addr1.sheet !== addr2.sheet) {
    throw new Error('Cannot create range across sheets');
  }

  return {
    start: {
      sheet: addr1.sheet,
      col: Math.min(addr1.col, addr2.col),
      row: Math.min(addr1.row, addr2.row)
    },
    end: {
      sheet: addr1.sheet,
      col: Math.max(addr1.col, addr2.col),
      row: Math.max(addr1.row, addr2.row)
    }
  };
}
