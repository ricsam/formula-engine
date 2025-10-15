import {
  AwaitingEvaluationError,
  EvaluationError,
} from "src/evaluator/evaluation-error";
import {
  FormulaError,
  type CellAddress,
  type ErrorEvaluationResult,
  type LocalCellAddress,
  type RangeAddress,
  type RelativeRange,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
} from "./types";

// Type narrowing utilities
export function isCellAddress(
  obj: RangeAddress | CellAddress
): obj is CellAddress {
  return (
    obj &&
    typeof obj === "object" &&
    "rowIndex" in obj &&
    "colIndex" in obj &&
    "sheetName" in obj &&
    "workbookName" in obj &&
    !("range" in obj)
  );
}

export function isRangeAddress(
  obj: RangeAddress | CellAddress
): obj is RangeAddress {
  return (
    obj &&
    typeof obj === "object" &&
    "range" in obj &&
    "sheetName" in obj &&
    "workbookName" in obj
  );
}

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
              value:
                range.start.col + range.width.value - 1 + relativeTo.colIndex,
            }
          : range.width,
      row:
        range.height.type === "number"
          ? {
              type: "number",
              value:
                range.start.row + range.height.value - 1 + relativeTo.rowIndex,
            }
          : range.height,
    },
  };
}

export function keyToCellAddress(key: string): CellAddress {
  const parts = key.split(":");

  if (parts.length < 2) {
    throw new Error(`Invalid dependency key format: ${key}`);
  }

  if (parts.length !== 4) {
    throw new Error(`Invalid cell key format: ${key}`);
  }
  const workbookName = parts[1];
  const sheetName = parts[2];
  const cellRef = parts[3];

  if (
    workbookName === undefined ||
    sheetName === undefined ||
    cellRef === undefined
  ) {
    throw new Error(`Invalid cell key format: ${key}`);
  }

  const { rowIndex, colIndex } = parseCellReference(cellRef);

  if (
    rowIndex === undefined ||
    colIndex === undefined ||
    Number.isNaN(rowIndex) ||
    Number.isNaN(colIndex)
  ) {
    throw new Error(`Invalid cell reference: ${cellRef}`);
  }

  return {
    workbookName,
    sheetName,
    colIndex,
    rowIndex,
  };
}

export function cellAddressToKey(cellAddress: CellAddress): string {
  if (
    cellAddress.rowIndex === undefined ||
    cellAddress.colIndex === undefined
  ) {
    throw new Error(
      `Invalid cell address: rowIndex and colIndex must be defined (got rowIndex=${cellAddress.rowIndex}, colIndex=${cellAddress.colIndex})`
    );
  }
  const cellRef = getCellReference(cellAddress);
  return `cell:${cellAddress.workbookName}:${cellAddress.sheetName}:${cellRef}`;
}

export function rangeAddressToKey(rangeAddress: RangeAddress): string {
  return `range:${rangeAddress.workbookName}:${rangeAddress.sheetName}:${getRangeKey(rangeAddress.range)}`;
}

/**
 *
 * @param key - the range key to parse e.g. range:workbook:sheet:A3:A5
 * @returns
 */
export function keyToRangeAddress(key: string): RangeAddress {
  const parts = key.split(":");
  if (parts.length !== 5) {
    throw new Error(`Invalid range key format: ${key}`);
  }
  const workbookName = parts[1];
  const sheetName = parts[2];
  const rangeKey = parts[3] + ":" + parts[4];
  return {
    workbookName: workbookName!,
    sheetName: sheetName!,
    range: parseRangeKey(rangeKey!),
  };
}

/**
 *
 * @param key - the range key to parse e.g. A3:A5 or A4:INFINITY
 * @returns
 */
function parseRangeKey(key: string): SpreadsheetRange {
  const parts = key.split(":");
  if (parts.length !== 2) {
    throw new Error(`Invalid range key format: ${key}`);
  }

  const start = parseCellReference(parts[0]!);
  const end = parts[1]!;

  let endRow: SpreadsheetRangeEnd;
  let endCol: SpreadsheetRangeEnd;

  if (end === "INFINITY") {
    endRow = {
      type: "infinity",
      sign: "positive",
    };
    endCol = {
      type: "infinity",
      sign: "positive",
    };
  } else if (!end?.match(/\d/)) {
    endRow = {
      type: "infinity",
      sign: "positive",
    };
    endCol = {
      type: "number",
      value: columnToIndex(end!),
    };
  } else if (!end?.match(/[A-Z]/)) {
    endRow = {
      type: "number",
      value: parseInt(end!, 10) - 1,
    };
    endCol = {
      type: "infinity",
      sign: "positive",
    };
  } else {
    const { rowIndex, colIndex } = parseCellReference(end!);
    endRow = {
      type: "number",
      value: rowIndex,
    };
    endCol = {
      type: "number",
      value: colIndex,
    };
  }

  return {
    start: {
      col: start.colIndex,
      row: start.rowIndex,
    },
    end: {
      row: endRow,
      col: endCol,
    },
  };
}

/**
 * Check if two ranges intersect
 */
export function checkRangeIntersection(
  range1: SpreadsheetRange,
  range2: SpreadsheetRange
): boolean {
  // Check if ranges don't intersect
  if (
    range1.end.col.type === "number" &&
    range2.start.col > range1.end.col.value
  )
    return false;
  if (
    range2.end.col.type === "number" &&
    range1.start.col > range2.end.col.value
  )
    return false;
  if (
    range1.end.row.type === "number" &&
    range2.start.row > range1.end.row.value
  )
    return false;
  if (
    range2.end.row.type === "number" &&
    range1.start.row > range2.end.row.value
  )
    return false;

  return true;
}

/**
 * Get the intersection of two ranges
 */
export function getRangeIntersection(
  range1: SpreadsheetRange,
  range2: SpreadsheetRange
): SpreadsheetRange | null {
  if (!checkRangeIntersection(range1, range2)) {
    return null;
  }

  const startRow = Math.max(range1.start.row, range2.start.row);
  const startCol = Math.max(range1.start.col, range2.start.col);

  let endRow, endCol;

  // Handle end row
  if (
    range1.end.row.type === "infinity" &&
    range2.end.row.type === "infinity"
  ) {
    endRow = { type: "infinity" as const, sign: "positive" as const };
  } else if (
    range1.end.row.type === "number" &&
    range2.end.row.type === "number"
  ) {
    endRow = {
      type: "number" as const,
      value: Math.min(range1.end.row.value, range2.end.row.value),
    };
  } else {
    // One is finite, one is infinite
    endRow = range1.end.row.type === "number" ? range1.end.row : range2.end.row;
  }

  // Handle end col
  if (
    range1.end.col.type === "infinity" &&
    range2.end.col.type === "infinity"
  ) {
    endCol = { type: "infinity" as const, sign: "positive" as const };
  } else if (
    range1.end.col.type === "number" &&
    range2.end.col.type === "number"
  ) {
    endCol = {
      type: "number" as const,
      value: Math.min(range1.end.col.value, range2.end.col.value),
    };
  } else {
    // One is finite, one is infinite
    endCol = range1.end.col.type === "number" ? range1.end.col : range2.end.col;
  }

  return {
    start: { row: startRow, col: startCol },
    end: { row: endRow, col: endCol },
  };
}

export function captureEvaluationErrors<T>(
  errAddress: CellAddress,
  fn: () => T
): T | ErrorEvaluationResult {
  try {
    return fn();
  } catch (error) {
    if (error instanceof EvaluationError) {
      return {
        type: "error",
        err: error.type,
        message: error.message,
        errAddress: error.errAddress ?? errAddress,
      };
    }
    if (error instanceof AwaitingEvaluationError) {
      return {
        type: "awaiting-evaluation",
        waitingFor: error.waitingFor,
        errAddress: error.errAddress,
      };
    }

    // Convert JavaScript errors to formula errors
    const formulaError =
      error instanceof Error
        ? mapJSErrorToFormulaError(error)
        : FormulaError.ERROR;

    return {
      type: "error",
      err: formulaError,
      message: (error as any)?.stack || "An error was thrown",
      errAddress,
    };
  }
}

function isFormulaError(value: string): value is FormulaError {
  if (typeof value !== "string") return false;

  // Check for all known formula errors
  const errors: FormulaError[] = Object.values(FormulaError);

  return errors.includes(value as FormulaError);
}

/**
 * Maps JavaScript errors to formula errors
 */
function mapJSErrorToFormulaError(error: Error): FormulaError {
  const message = error.message.toLowerCase();

  if (isFormulaError(error.message)) {
    return error.message;
  }

  if (
    message.includes("division by zero") ||
    message.includes("divide by zero")
  ) {
    return FormulaError.DIV0;
  }
  if (message.includes("circular") || message.includes("cycle")) {
    return FormulaError.CYCLE;
  }
  if (
    message.includes("invalid reference") ||
    (message.includes("reference") && !message.includes("circular"))
  ) {
    return FormulaError.REF;
  }
  if (
    message.includes("invalid name") ||
    message.includes("unknown function")
  ) {
    return FormulaError.NAME;
  }
  if (
    message.includes("invalid number") ||
    message.includes("nan") ||
    message.includes("infinity")
  ) {
    return FormulaError.NUM;
  }
  if (message.includes("type") || message.includes("invalid argument")) {
    return FormulaError.VALUE;
  }
  if (message.includes("not available") || message.includes("n/a")) {
    return FormulaError.NA;
  }

  return FormulaError.ERROR;
}
