import type {
  CellAddress,
  LocalCellAddress,
  RelativeRange,
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

/**
 * Returns true if the range is a single cell, false otherwise
 */
export function isRangeOneCell(range: SpreadsheetRange) {
  if (range.end.col.type === "infinity" || range.end.row.type === "infinity") {
    return false;
  }
  return (
    range.start.col === range.end.col.value &&
    range.start.row === range.end.row.value
  );
}

export function isCellInRange(
  cellAddress: LocalCellAddress,
  range: SpreadsheetRange
) {
  const endCol = range.end.col;
  const endRow = range.end.row;
  if (endCol.type === "number" && endRow.type === "number") {
    return (
      cellAddress.colIndex >= range.start.col &&
      cellAddress.colIndex <= endCol.value &&
      cellAddress.rowIndex >= range.start.row &&
      cellAddress.rowIndex <= endRow.value
    );
  } else if (endCol.type === "infinity" && endRow.type === "number") {
    return (
      cellAddress.colIndex >= range.start.col &&
      cellAddress.rowIndex >= range.start.row &&
      cellAddress.rowIndex <= endRow.value
    );
  } else if (endCol.type === "number" && endRow.type === "infinity") {
    return (
      cellAddress.rowIndex >= range.start.row &&
      cellAddress.colIndex >= range.start.col &&
      cellAddress.colIndex <= endCol.value
    );
  } else if (endCol.type === "infinity" && endRow.type === "infinity") {
    return (
      cellAddress.colIndex >= range.start.col &&
      cellAddress.rowIndex >= range.start.row
    );
  }
  return false;
}

export function getRangeKey(range: SpreadsheetRange) {
  let rangeKey = "";
  rangeKey += getCellReference({
    rowIndex: range.start.row,
    colIndex: range.start.col,
  });
  rangeKey += ":";
  if (range.end.col.type === "number" && range.end.row.type === "number") {
    rangeKey += getCellReference({
      rowIndex: range.end.row.value,
      colIndex: range.end.col.value,
    });
  } else if (
    range.end.col.type === "infinity" &&
    range.end.row.type === "number"
  ) {
    rangeKey += `${getRowNumber(range.end.row.value)}`;
  } else if (
    range.end.col.type === "number" &&
    range.end.row.type === "infinity"
  ) {
    rangeKey += indexToColumn(range.end.col.value);
  } else if (
    range.end.col.type === "infinity" &&
    range.end.row.type === "infinity"
  ) {
    rangeKey += "INFINITY";
  }
  return rangeKey;
}

export function getRelativeRangeKey(range: RelativeRange): string {
  let rangeKey = "";
  rangeKey += getCellReference({
    rowIndex: range.start.row,
    colIndex: range.start.col,
  });
  rangeKey += ":";

  if (range.width.type === "number") {
    rangeKey += range.width.value;
  } else {
    rangeKey += "INFINITY";
  }

  rangeKey += "x";

  if (range.height.type === "number") {
    rangeKey += range.height.value;
  } else {
    rangeKey += "INFINITY";
  }

  return rangeKey;
}

export function getRelativeRange(
  range: SpreadsheetRange,
  relativeTo: LocalCellAddress
): RelativeRange {
  const dx = range.start.col - relativeTo.colIndex;
  const dy = range.start.row - relativeTo.rowIndex;
  const relativeRange: RelativeRange = {
    start: {
      col: dx,
      row: dy,
    },
    width:
      range.end.col.type === "number"
        ? {
            type: "number",
            value: range.end.col.value - range.start.col + 1,
          }
        : range.end.col,
    height:
      range.end.row.type === "number"
        ? {
            type: "number",
            value: range.end.row.value - range.start.row + 1,
          }
        : range.end.row,
  };

  return relativeRange;
}

export function getAbsoluteRange(
  range: RelativeRange,
  relativeTo: LocalCellAddress
): SpreadsheetRange {
  return {
    start: {
      col: range.start.col + relativeTo.colIndex,
      row: range.start.row + relativeTo.rowIndex,
    },
    end: {
      col:
        range.width.type === "number"
          ? {
              type: "number",
              value: range.start.col + range.width.value - 1 + relativeTo.colIndex,
            }
          : range.width,
      row:
        range.height.type === "number"
          ? {
              type: "number",
              value: range.start.row + range.height.value - 1 + relativeTo.rowIndex,
            }
          : range.height,
    },
  };
}
