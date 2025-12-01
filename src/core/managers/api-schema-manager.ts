/**
 * ApiSchemaManager - Manages API schema definitions and validation
 *
 * Tracks registered schemas and provides validation for cell writes,
 * spill operations, and copy/paste operations.
 */

import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRange,
  TableDefinition,
} from "../types";
import { cellAddressToKey, isCellInRange } from "../utils";
import type { TableManager } from "./table-manager";

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

export interface RegisteredTableSchema {
  type: "table";
  namespace: string;
  workbookName: string;
  tableName: string;
  headers: TableSchemaHeaders;
  isValid: boolean;
}

export interface RegisteredCellSchema {
  type: "cell";
  namespace: string;
  cellAddress: CellAddress;
  parse: ParseFunction;
  isValid: boolean;
}

export type RegisteredSchema = RegisteredTableSchema | RegisteredCellSchema;

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

export class ApiSchemaManager {
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
      isValid: true,
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
      isValid: true,
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
   * Check if a cell is protected by a schema
   * Returns the schema info if protected, null otherwise
   */
  isCellProtected(
    cell: CellAddress
  ): { schema: RegisteredSchema; columnName?: string } | null {
    for (const schema of this.schemas.values()) {
      if (!schema.isValid) continue;

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
      }
      return { valid: true };
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      return {
        valid: false,
        error: `Schema validation failed for ${schema.namespace}${columnName ? `.${columnName}` : ""}: ${error.message}`,
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
              error: `Spill blocked by schema at ${cellAddressToKey(cell)}: ${validation.error}`,
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
      }
    }
  }

  /**
   * Invalidate schemas when a sheet is deleted
   */
  invalidateForSheetDeletion(workbookName: string, sheetName: string): void {
    for (const schema of this.schemas.values()) {
      if (schema.type === "cell") {
        if (
          schema.cellAddress.workbookName === workbookName &&
          schema.cellAddress.sheetName === sheetName
        ) {
          schema.isValid = false;
        }
      }
      // Table schemas will be invalidated when the table itself is deleted
    }
  }

  /**
   * Invalidate schemas when a table is deleted
   */
  invalidateForTableDeletion(workbookName: string, tableName: string): void {
    for (const schema of this.schemas.values()) {
      if (
        schema.type === "table" &&
        schema.workbookName === workbookName &&
        schema.tableName === tableName
      ) {
        schema.isValid = false;
      }
    }
  }

  /**
   * Invalidate schemas when a workbook is deleted
   */
  invalidateForWorkbookDeletion(workbookName: string): void {
    for (const schema of this.schemas.values()) {
      if (schema.type === "cell") {
        if (schema.cellAddress.workbookName === workbookName) {
          schema.isValid = false;
        }
      } else if (schema.type === "table") {
        if (schema.workbookName === workbookName) {
          schema.isValid = false;
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
}

