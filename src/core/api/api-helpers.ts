/**
 * API Helper utilities for row/object conversion and filtering
 */

import type { CellAddress, SerializedCellValue, TableDefinition } from "../types";
import type { FormulaEngine } from "../engine";

export type ParseFunction<TCellMetadata = unknown> = (
  value: unknown,
  metadata: TCellMetadata
) => unknown;

export type TableSchemaHeaders<TCellMetadata = unknown> = Record<
  string,
  {
    parse: ParseFunction<TCellMetadata>;
    index: number;
  }
>;

/**
 * Convert a spreadsheet row to a typed object using header parsers
 */
export function rowToObject<TItem extends Record<string, unknown>>(
  engine: FormulaEngine<any, any>,
  table: TableDefinition,
  rowIndex: number,
  headers: TableSchemaHeaders,
  getCellMetadata?: (cell: CellAddress) => unknown
): TItem {
  const result: Record<string, unknown> = {};

  for (const [columnName, header] of Object.entries(headers)) {
    const cellAddress: CellAddress = {
      workbookName: table.workbookName,
      sheetName: table.sheetName,
      colIndex: table.start.colIndex + header.index,
      rowIndex,
    };

    const value = engine.getCellValue(cellAddress);
    const metadata = getCellMetadata ? getCellMetadata(cellAddress) : undefined;

    // Parse the value using the column's parse function
    result[columnName] = header.parse(value, metadata);
  }

  return result as TItem;
}

/**
 * Convert a typed object to cell values based on headers
 */
export function objectToRowValues<TItem extends Record<string, unknown>>(
  obj: TItem,
  headers: TableSchemaHeaders
): Map<number, SerializedCellValue> {
  const values = new Map<number, SerializedCellValue>();

  for (const [columnName, header] of Object.entries(headers)) {
    if (columnName in obj) {
      const value = obj[columnName];
      // Convert to serializable value
      if (
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean"
      ) {
        values.set(header.index, value);
      } else if (value === null || value === undefined) {
        values.set(header.index, undefined);
      } else {
        // For complex types, convert to string
        values.set(header.index, String(value));
      }
    }
  }

  return values;
}

/**
 * Check if an object matches a filter predicate
 */
export function matchesFilter<TItem extends Record<string, unknown>>(
  item: TItem,
  filter: Partial<TItem>
): boolean {
  for (const [key, value] of Object.entries(filter)) {
    if (item[key] !== value) {
      return false;
    }
  }
  return true;
}

/**
 * Get the number of data rows in a table (excluding header)
 */
export function getTableRowCount(
  engine: FormulaEngine<any, any>,
  table: TableDefinition
): number {
  if (table.endRow.type === "infinity") {
    // For infinite tables, count non-empty rows
    // Start from data row (after header)
    let count = 0;
    let emptyRowStreak = 0;
    const maxEmptyRows = 10; // Stop after 10 consecutive empty rows

    for (let row = table.start.rowIndex + 1; emptyRowStreak < maxEmptyRows; row++) {
      let hasContent = false;

      for (let col = 0; col < table.headers.size; col++) {
        const cellAddress: CellAddress = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          colIndex: table.start.colIndex + col,
          rowIndex: row,
        };

        const value = engine.getCellValue(cellAddress);
        if (value !== "" && value !== undefined) {
          hasContent = true;
          break;
        }
      }

      if (hasContent) {
        count++;
        emptyRowStreak = 0;
      } else {
        emptyRowStreak++;
      }
    }

    return count;
  } else {
    // For finite tables, calculate from endRow
    return table.endRow.value - table.start.rowIndex;
  }
}

/**
 * Get the next empty row index for appending data
 */
export function getNextEmptyRow(
  engine: FormulaEngine<any, any>,
  table: TableDefinition
): number {
  const dataStartRow = table.start.rowIndex + 1;

  if (table.endRow.type === "infinity") {
    // Find first empty row
    let emptyRowStreak = 0;
    const maxEmptyRows = 10;

    for (let row = dataStartRow; ; row++) {
      let hasContent = false;

      for (let col = 0; col < table.headers.size; col++) {
        const cellAddress: CellAddress = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          colIndex: table.start.colIndex + col,
          rowIndex: row,
        };

        const value = engine.getCellValue(cellAddress);
        if (value !== "" && value !== undefined) {
          hasContent = true;
          break;
        }
      }

      if (!hasContent) {
        return row;
      }

      emptyRowStreak = hasContent ? 0 : emptyRowStreak + 1;
      if (emptyRowStreak >= maxEmptyRows) {
        return row - maxEmptyRows + 1;
      }
    }
  } else {
    // For finite tables, return the row after the last data row
    return table.endRow.value + 1;
  }
}

/**
 * Iterate over all data rows in a table
 */
export function* iterateTableRows(
  engine: FormulaEngine<any, any>,
  table: TableDefinition
): Generator<{ rowIndex: number; isEmpty: boolean }> {
  const dataStartRow = table.start.rowIndex + 1;

  if (table.endRow.type === "infinity") {
    let emptyRowStreak = 0;
    const maxEmptyRows = 10;

    for (let row = dataStartRow; emptyRowStreak < maxEmptyRows; row++) {
      let hasContent = false;

      for (let col = 0; col < table.headers.size; col++) {
        const cellAddress: CellAddress = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          colIndex: table.start.colIndex + col,
          rowIndex: row,
        };

        const value = engine.getCellValue(cellAddress);
        if (value !== "" && value !== undefined) {
          hasContent = true;
          break;
        }
      }

      if (hasContent) {
        emptyRowStreak = 0;
        yield { rowIndex: row, isEmpty: false };
      } else {
        emptyRowStreak++;
        // Don't yield empty rows at the end
      }
    }
  } else {
    for (let row = dataStartRow; row <= table.endRow.value; row++) {
      let hasContent = false;

      for (let col = 0; col < table.headers.size; col++) {
        const cellAddress: CellAddress = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          colIndex: table.start.colIndex + col,
          rowIndex: row,
        };

        const value = engine.getCellValue(cellAddress);
        if (value !== "" && value !== undefined) {
          hasContent = true;
          break;
        }
      }

      yield { rowIndex: row, isEmpty: !hasContent };
    }
  }
}

