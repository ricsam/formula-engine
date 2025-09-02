import type { SpreadsheetRange } from "../types";
import { indexToColumn, getRowNumber } from "../utils";

/**
 * Serializes a SpreadsheetRange into a human-readable range string following canonical format.
 * 
 * Examples:
 * - Finite range: A2:B10
 * - Column-infinite range: A2:B (row infinite, column bounded)
 * - Row-infinite range: A2:10 (column infinite, row bounded)
 * - Fully infinite range: A2:INFINITY (both infinite)
 * 
 * @param range The SpreadsheetRange to serialize
 * @returns A human-readable range string in canonical format
 */
export function serializeRange(range: SpreadsheetRange): string {
  // Format start cell
  const startCell = `${indexToColumn(range.start.col)}${getRowNumber(range.start.row)}`;
  
  // Determine end format based on infinity types
  const endRowIsInfinity = range.end.row.type === "infinity";
  const endColIsInfinity = range.end.col.type === "infinity";
  
  let rangeEnd: string;
  
  if (endRowIsInfinity && endColIsInfinity) {
    // Both infinite: A2:INFINITY
    rangeEnd = "INFINITY";
  } else if (endRowIsInfinity && range.end.col.type === "number") {
    // Row infinite, col finite: A2:B (column only)
    rangeEnd = indexToColumn(range.end.col.value);
  } else if (range.end.row.type === "number" && endColIsInfinity) {
    // Row finite, col infinite: A2:10 (row only)
    rangeEnd = getRowNumber(range.end.row.value).toString();
  } else if (range.end.row.type === "number" && range.end.col.type === "number") {
    // Both finite: A2:B10
    rangeEnd = `${indexToColumn(range.end.col.value)}${getRowNumber(range.end.row.value)}`;
  } else {
    throw new Error("Invalid range end configuration");
  }
  
  return `${startCell}:${rangeEnd}`;
}

/**
 * Serializes a SpreadsheetRange with an optional sheet name prefix.
 * 
 * @param range The SpreadsheetRange to serialize
 * @param sheetName Optional sheet name to prefix the range
 * @returns A human-readable range string, optionally prefixed with sheet name
 */
export function serializeRangeWithSheet(range: SpreadsheetRange, sheetName?: string): string {
  const rangeStr = serializeRange(range);
  
  if (sheetName) {
    // Handle sheet names that need quoting (contain spaces or special characters)
    const needsQuotes = /[^A-Za-z0-9_]/.test(sheetName);
    const quotedSheetName = needsQuotes ? `'${sheetName}'` : sheetName;
    return `${quotedSheetName}!${rangeStr}`;
  }
  
  return rangeStr;
}
