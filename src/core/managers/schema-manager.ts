/**
 * SchemaManager - Manages schema definitions and validation
 *
 * Tracks registered schemas and provides validation for cell writes,
 * spill operations, and copy/paste operations.
 */

import type {
  CellAddress,
  FiniteSpreadsheetRange,
  SerializedCellValue,
  SpreadsheetRange,
  TableDefinition,
} from "../types";
import { cellAddressToKey, isCellInRange } from "../utils";
import type { TableManager } from "./table-manager";

export type ParseFunction<TCellMetadata = unknown> = (
  value: SerializedCellValue,
  metadata: TCellMetadata
) => unknown;

export type WriteFunction<TCellMetadata = unknown> = (value: unknown) => {
  value: SerializedCellValue;
  metadata?: TCellMetadata;
};

export type TableSchemaHeaders<TCellMetadata = unknown> = Record<
  string,
  {
    parse: ParseFunction<TCellMetadata>;
    write: WriteFunction<TCellMetadata>;
    index: number;
  }
>;

export interface RegisteredTableSchema {
  type: "table";
  namespace: string;
  workbookName: string;
  tableName: string;
  headers: TableSchemaHeaders;
}

export interface RegisteredCellSchema {
  type: "cell";
  namespace: string;
  cellAddress: CellAddress;
  parse: ParseFunction;
}

export interface RegisteredGridSchema {
  type: "grid";
  namespace: string;
  workbookName: string;
  sheetName: string;
  range: FiniteSpreadsheetRange;
  parse: ParseFunction;
}

export type RegisteredSchema = RegisteredTableSchema | RegisteredCellSchema | RegisteredGridSchema;

export interface ValidationResult {
  valid: boolean;
  error?: string;
  originalError?: Error;
}

export class SchemaValidationError extends Error {
  constructor(
    message: string,
    public cellAddress?: CellAddress,
    public originalError?: Error
  ) {
    super(message);
    this.name = "SchemaValidationError";
  }
}

export class SchemaManager {
  private schemas: Map<string, RegisteredSchema> = new Map();

  constructor(private tableManager: TableManager) {}

  /**
   * Register a table schema
   */
  registerTableSchema(
    namespace: string,
    workbookName: string,
    tableName: string,
    headers: TableSchemaHeaders
  ): void {
    this.schemas.set(namespace, {
      type: "table",
      namespace,
      workbookName,
      tableName,
      headers,
    });
  }

  /**
   * Register a cell schema
   */
  registerCellSchema(
    namespace: string,
    cellAddress: CellAddress,
    parse: ParseFunction
  ): void {
    this.schemas.set(namespace, {
      type: "cell",
      namespace,
      cellAddress,
      parse,
    });
  }

  /**
   * Register a grid schema
   */
  registerGridSchema(
    namespace: string,
    workbookName: string,
    sheetName: string,
    range: FiniteSpreadsheetRange,
    parse: ParseFunction
  ): void {
    this.schemas.set(namespace, {
      type: "grid",
      namespace,
      workbookName,
      sheetName,
      range,
      parse,
    });
  }

  /**
   * Get a schema by namespace
   */
  getSchema(namespace: string): RegisteredSchema | undefined {
    return this.schemas.get(namespace);
  }

  /**
   * Remove a schema
   */
  removeSchema(namespace: string): boolean {
    return this.schemas.delete(namespace);
  }

  /**
   * Check if any schemas are registered
   */
  hasSchemas(): boolean {
    return this.schemas.size > 0;
  }

  /**
   * Check if a cell is protected by a schema
   * Returns the schema info if protected, null otherwise
   */
  isCellProtected(
    cell: CellAddress
  ): { schema: RegisteredSchema; columnName?: string } | null {
    for (const schema of this.schemas.values()) {
      if (schema.type === "cell") {
        if (
          schema.cellAddress.workbookName === cell.workbookName &&
          schema.cellAddress.sheetName === cell.sheetName &&
          schema.cellAddress.colIndex === cell.colIndex &&
          schema.cellAddress.rowIndex === cell.rowIndex
        ) {
          return { schema };
        }
      } else if (schema.type === "table") {
        const table = this.tableManager.getTable({
          workbookName: schema.workbookName,
          name: schema.tableName,
        });

        if (!table) continue;

        // Check if cell is in table's data range (excluding header row)
        if (
          cell.workbookName === table.workbookName &&
          cell.sheetName === table.sheetName
        ) {
          const isInDataRange = this.isCellInTableDataRange(cell, table);
          if (isInDataRange) {
            // Find which column this cell is in
            const colOffset = cell.colIndex - table.start.colIndex;
            const columnName = this.getColumnNameByIndex(
              schema.headers,
              colOffset
            );
            return { schema, columnName };
          }
        }
      } else if (schema.type === "grid") {
        // Check if cell is within the grid range
        if (
          cell.workbookName === schema.workbookName &&
          cell.sheetName === schema.sheetName &&
          cell.colIndex >= schema.range.start.col &&
          cell.colIndex <= schema.range.end.col &&
          cell.rowIndex >= schema.range.start.row &&
          cell.rowIndex <= schema.range.end.row
        ) {
          return { schema };
        }
      }
    }

    return null;
  }

  /**
   * Check if a cell is within a table's data range (excluding header row)
   */
  private isCellInTableDataRange(
    cell: CellAddress,
    table: TableDefinition
  ): boolean {
    const { start, endRow, headers } = table;

    // Data starts one row below header
    const dataStartRow = start.rowIndex + 1;

    // Check row bounds (data rows only, not header)
    const isInRowRange =
      endRow.type === "infinity"
        ? cell.rowIndex >= dataStartRow
        : cell.rowIndex >= dataStartRow && cell.rowIndex <= endRow.value;

    // Check column bounds
    const endColIndex = start.colIndex + headers.size - 1;
    const isInColRange =
      cell.colIndex >= start.colIndex && cell.colIndex <= endColIndex;

    return isInRowRange && isInColRange;
  }

  /**
   * Get column name by index from headers
   */
  private getColumnNameByIndex(
    headers: TableSchemaHeaders,
    index: number
  ): string | undefined {
    for (const [name, header] of Object.entries(headers)) {
      if (header.index === index) {
        return name;
      }
    }
    return undefined;
  }

  /**
   * Validate a cell write against schemas
   */
  validateCellWrite(
    cell: CellAddress,
    value: SerializedCellValue,
    metadata?: unknown
  ): ValidationResult {
    // Allow clearing cells (undefined or empty string)
    if (value === undefined || value === "") {
      return { valid: true };
    }

    const protection = this.isCellProtected(cell);

    if (!protection) {
      return { valid: true };
    }

    const { schema, columnName } = protection;

    try {
      if (schema.type === "cell") {
        schema.parse(value, metadata);
      } else if (schema.type === "table" && columnName) {
        const header = schema.headers[columnName];
        if (header) {
          header.parse(value, metadata);
        }
      } else if (schema.type === "grid") {
        schema.parse(value, metadata);
      }
      return { valid: true };
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      return {
        valid: false,
        error: `Schema validation failed for ${schema.namespace}${
          columnName ? `.${columnName}` : ""
        }: ${error.message}`,
        originalError: error,
      };
    }
  }

  /**
   * Validate a spill area against schemas
   * Returns validation result - if invalid, spill should produce #SPILL! error
   */
  validateSpillArea(
    origin: CellAddress,
    spillArea: SpreadsheetRange,
    getSpillValue: (row: number, col: number) => SerializedCellValue,
    getMetadata?: (cell: CellAddress) => unknown
  ): ValidationResult {
    // Get the bounds of the spill area
    const startRow = spillArea.start.row;
    const startCol = spillArea.start.col;
    const endRow =
      spillArea.end.row.type === "number"
        ? spillArea.end.row.value
        : startRow + 100; // Limit for infinite ranges
    const endCol =
      spillArea.end.col.type === "number"
        ? spillArea.end.col.value
        : startCol + 100; // Limit for infinite ranges

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell: CellAddress = {
          workbookName: origin.workbookName,
          sheetName: origin.sheetName,
          rowIndex: row,
          colIndex: col,
        };

        const protection = this.isCellProtected(cell);
        if (protection) {
          const value = getSpillValue(row - startRow, col - startCol);
          const metadata = getMetadata ? getMetadata(cell) : undefined;
          const validation = this.validateCellWrite(cell, value, metadata);

          if (!validation.valid) {
            return {
              valid: false,
              error: `Spill blocked by schema at ${cellAddressToKey(cell)}: ${
                validation.error
              }`,
              originalError: validation.originalError,
            };
          }
        }
      }
    }

    return { valid: true };
  }

  /**
   * Update schemas when a table is renamed
   */
  updateForTableRename(
    workbookName: string,
    oldTableName: string,
    newTableName: string
  ): void {
    for (const schema of this.schemas.values()) {
      if (
        schema.type === "table" &&
        schema.workbookName === workbookName &&
        schema.tableName === oldTableName
      ) {
        schema.tableName = newTableName;
      }
    }
  }

  /**
   * Update schemas when a sheet is renamed
   */
  updateForSheetRename(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    for (const schema of this.schemas.values()) {
      if (schema.type === "cell") {
        if (
          schema.cellAddress.workbookName === workbookName &&
          schema.cellAddress.sheetName === oldSheetName
        ) {
          schema.cellAddress.sheetName = newSheetName;
        }
      } else if (schema.type === "grid") {
        if (
          schema.workbookName === workbookName &&
          schema.sheetName === oldSheetName
        ) {
          schema.sheetName = newSheetName;
        }
      }
      // Table schemas don't store sheetName directly - they get it from TableManager
    }
  }

  /**
   * Update schemas when a workbook is renamed
   */
  updateForWorkbookRename(
    oldWorkbookName: string,
    newWorkbookName: string
  ): void {
    for (const schema of this.schemas.values()) {
      if (schema.type === "cell") {
        if (schema.cellAddress.workbookName === oldWorkbookName) {
          schema.cellAddress.workbookName = newWorkbookName;
        }
      } else if (schema.type === "table") {
        if (schema.workbookName === oldWorkbookName) {
          schema.workbookName = newWorkbookName;
        }
      } else if (schema.type === "grid") {
        if (schema.workbookName === oldWorkbookName) {
          schema.workbookName = newWorkbookName;
        }
      }
    }
  }

  /**
   * Get all schemas (for debugging/testing)
   */
  getAllSchemas(): Map<string, RegisteredSchema> {
    return new Map(this.schemas);
  }

  /**
   * Validate all schema constraints using evaluated cell values.
   *
   * This is the key method for schema validation with rollback:
   * - It validates EVALUATED values, not raw content
   * - A cell with "=123+123" validates as number (246), not string
   * - Called after re-evaluation to check all schema-constrained cells
   *
   * @param getCellValue - Function to get the evaluated value of a cell
   * @param getCellMetadata - Optional function to get cell metadata
   * @returns Validation result with all errors found
   */
  validateAllSchemaConstraints(
    getCellValue: (cell: CellAddress) => SerializedCellValue,
    getCellMetadata?: (cell: CellAddress) => unknown,
    getTableDataCells?: (table: TableDefinition) => CellAddress[]
  ): {
    valid: boolean;
    errors: Array<{
      message: string;
      cellAddress?: CellAddress;
      schemaNamespace?: string;
      columnName?: string;
      originalError?: Error;
    }>;
  } {
    const errors: Array<{
      message: string;
      cellAddress?: CellAddress;
      schemaNamespace?: string;
      columnName?: string;
      originalError?: Error;
    }> = [];

    for (const [namespace, schema] of this.schemas) {
      if (schema.type === "cell") {
        // Validate single cell using evaluated value
        let value: SerializedCellValue;
        try {
          value = getCellValue(schema.cellAddress);
        } catch {
          // Sheet/cell doesn't exist yet, skip validation
          continue;
        }

        const metadata = getCellMetadata
          ? getCellMetadata(schema.cellAddress)
          : undefined;

        // Skip empty cells
        if (value === undefined || value === "") {
          continue;
        }

        try {
          schema.parse(value, metadata);
        } catch (err) {
          const error = err instanceof Error ? err : new Error(String(err));
          errors.push({
            message: `Schema validation failed for ${namespace}: ${error.message}`,
            cellAddress: schema.cellAddress,
            schemaNamespace: namespace,
            originalError: error,
          });
        }
      } else if (schema.type === "table") {
        // Get the table definition
        const table = this.tableManager.getTable({
          workbookName: schema.workbookName,
          name: schema.tableName,
        });

        if (!table) continue;

        const { start, headers } = table;

        // Get cells to validate - use callback if provided (handles infinite ranges efficiently)
        // Otherwise fall back to simple iteration for finite tables
        let cellsToValidate: CellAddress[];
        if (getTableDataCells) {
          cellsToValidate = getTableDataCells(table);
        } else {
          // Fallback for finite tables when no callback provided
          const { endRow } = table;
          if (endRow.type === "infinity") {
            // Skip validation for infinite tables without a cell iterator
            continue;
          }
          cellsToValidate = [];
          const dataStartRow = start.rowIndex + 1;
          const endColIndex = start.colIndex + headers.size - 1;
          for (let row = dataStartRow; row <= endRow.value; row++) {
            for (let col = start.colIndex; col <= endColIndex; col++) {
              cellsToValidate.push({
                workbookName: table.workbookName,
                sheetName: table.sheetName,
                rowIndex: row,
                colIndex: col,
              });
            }
          }
        }

        for (const cell of cellsToValidate) {
          // Skip header row
          if (cell.rowIndex <= start.rowIndex) continue;

          let value: SerializedCellValue;
          try {
            value = getCellValue(cell);
          } catch {
            // Sheet/cell doesn't exist yet, skip validation
            continue;
          }

          // Skip empty cells
          if (value === undefined || value === "") {
            continue;
          }

          // Find the column name
          const colOffset = cell.colIndex - start.colIndex;
          const columnName = this.getColumnNameByIndex(
            schema.headers,
            colOffset
          );

          if (columnName && schema.headers[columnName]) {
            const header = schema.headers[columnName];
            const metadata = getCellMetadata
              ? getCellMetadata(cell)
              : undefined;

            try {
              header.parse(value, metadata);
            } catch (err) {
              const error =
                err instanceof Error ? err : new Error(String(err));
              errors.push({
                message: `Schema validation failed for ${namespace}.${columnName}: ${error.message}`,
                cellAddress: cell,
                schemaNamespace: namespace,
                columnName,
                originalError: error,
              });
            }
          }
        }
      } else if (schema.type === "grid") {
        // Validate all cells in the grid range
        const { range } = schema;

        for (let row = range.start.row; row <= range.end.row; row++) {
          for (let col = range.start.col; col <= range.end.col; col++) {
            const cell: CellAddress = {
              workbookName: schema.workbookName,
              sheetName: schema.sheetName,
              rowIndex: row,
              colIndex: col,
            };

            let value: SerializedCellValue;
            try {
              value = getCellValue(cell);
            } catch {
              // Sheet/cell doesn't exist yet, skip validation
              continue;
            }

            // Skip empty cells
            if (value === undefined || value === "") {
              continue;
            }

            const metadata = getCellMetadata ? getCellMetadata(cell) : undefined;

            try {
              schema.parse(value, metadata);
            } catch (err) {
              const error = err instanceof Error ? err : new Error(String(err));
              errors.push({
                message: `Schema validation failed for ${namespace}: ${error.message}`,
                cellAddress: cell,
                schemaNamespace: namespace,
                originalError: error,
              });
            }
          }
        }
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }
}
