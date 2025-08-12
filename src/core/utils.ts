import type {
  CellAddress,
  CellNumber,
  DependencyNode,
  SpreadsheetRange,
} from "./types";

// Column utilities
export const columnToIndex = (col: string): number => {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64); // A=1, B=2, etc.
  }
  return result - 1; // Convert to 0-based index
};

export const indexToColumn = (index: number): string => {
  let result = "";
  let num = index + 1; // Convert to 1-based

  while (num > 0) {
    const remainder = (num - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    num = Math.floor((num - 1) / 26);
  }

  return result;
};

export function getRowNumber(index: number): number {
  return index + 1;
}

export function getCellReference({
  rowIndex,
  colIndex,
}: {
  rowIndex: number;
  colIndex: number;
}) {
  return `${indexToColumn(colIndex)}${getRowNumber(rowIndex)}`;
}

// Convert row number to letter(s) for reversed headers (1 -> A, 2 -> B, etc.)
export function rowToLetter(row: number): string {
  let result = "";
  let num = row - 1; // Convert to 0-based

  do {
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26) - 1;
  } while (num >= 0);

  return result;
}

// Convert letter(s) to row number for reversed headers (A -> 1, B -> 2, etc.)
export function letterToRow(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

export const parseCellReference = (
  cellReference: string
): {
  colIndex: number;
  rowIndex: number;
} => {
  const match = cellReference.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${cellReference}`);
  }

  const column = match[1];
  if (match[2] === undefined || column === undefined) {
    throw new Error(`Invalid cell reference: ${cellReference}`);
  }
  const row = parseInt(match[2], 10);
  const colIndex = columnToIndex(column);
  const rowIndex = row - 1; // Convert to 0-based index

  return {
    colIndex,
    rowIndex,
  };
};

