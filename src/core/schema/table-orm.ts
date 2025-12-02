/**
 * TableOrm - Object-Relational Mapping for table data
 *
 * Provides CRUD operations for table data with type-safe access.
 * Used as the `this` context for custom table API methods.
 */

import type { CellAddress, SerializedCellValue, TableDefinition } from "../types";
import type { FormulaEngine } from "../engine";
import {
  rowToObject,
  objectToRowValues,
  matchesFilter,
  getTableRowCount,
  getNextEmptyRow,
  iterateTableRows,
  type TableSchemaHeaders,
} from "./schema-helpers";

export class TableOrm<TItem extends Record<string, unknown>> {
  constructor(
    private engine: FormulaEngine<any, any>,
    private workbookName: string,
    private tableName: string,
    private headers: TableSchemaHeaders,
    private namespace: string
  ) {}

  /**
   * Get the table definition from the engine
   */
  private getTable(): TableDefinition {
    const table = this.engine._tableManager.getTable({
      workbookName: this.workbookName,
      name: this.tableName,
    });

    if (!table) {
      throw new Error(
        `Table "${this.tableName}" not found in workbook "${this.workbookName}"`
      );
    }

    return table;
  }

  /**
   * Find the first row matching the filter
   */
  findWhere(filter: Partial<TItem>): TItem | undefined {
    const table = this.getTable();

    for (const { rowIndex, isEmpty } of iterateTableRows(this.engine, table)) {
      if (isEmpty) continue;

      try {
        const item = rowToObject<TItem>(
          this.engine,
          table,
          rowIndex,
          this.headers,
          (cell) => this.engine.getCellMetadata(cell)
        );

        if (matchesFilter(item, filter)) {
          return item;
        }
      } catch {
        // Skip rows that fail parsing
        continue;
      }
    }

    return undefined;
  }

  /**
   * Find all rows matching the filter
   */
  findAllWhere(filter: Partial<TItem>): TItem[] {
    const table = this.getTable();
    const results: TItem[] = [];

    for (const { rowIndex, isEmpty } of iterateTableRows(this.engine, table)) {
      if (isEmpty) continue;

      try {
        const item = rowToObject<TItem>(
          this.engine,
          table,
          rowIndex,
          this.headers,
          (cell) => this.engine.getCellMetadata(cell)
        );

        if (matchesFilter(item, filter)) {
          results.push(item);
        }
      } catch {
        // Skip rows that fail parsing
        continue;
      }
    }

    return results;
  }

  /**
   * Append a new row to the table
   */
  append(item: TItem): TItem {
    const table = this.getTable();
    const nextRow = getNextEmptyRow(this.engine, table);
    const values = objectToRowValues(item, this.headers);

    // Write each cell value
    for (const [colOffset, value] of values) {
      const cellAddress: CellAddress = {
        workbookName: table.workbookName,
        sheetName: table.sheetName,
        colIndex: table.start.colIndex + colOffset,
        rowIndex: nextRow,
      };

      this.engine.setCellContent(cellAddress, value);
    }

    return item;
  }

  /**
   * Update all rows matching the filter
   * Returns the number of rows updated
   */
  updateWhere(filter: Partial<TItem>, update: Partial<TItem>): number {
    const table = this.getTable();
    let updatedCount = 0;

    for (const { rowIndex, isEmpty } of iterateTableRows(this.engine, table)) {
      if (isEmpty) continue;

      try {
        const item = rowToObject<TItem>(
          this.engine,
          table,
          rowIndex,
          this.headers,
          (cell) => this.engine.getCellMetadata(cell)
        );

        if (matchesFilter(item, filter)) {
          // Apply updates
          const values = objectToRowValues(update as TItem, this.headers);

          for (const [colOffset, value] of values) {
            const cellAddress: CellAddress = {
              workbookName: table.workbookName,
              sheetName: table.sheetName,
              colIndex: table.start.colIndex + colOffset,
              rowIndex,
            };

            this.engine.setCellContent(cellAddress, value);
          }

          updatedCount++;
        }
      } catch {
        // Skip rows that fail parsing
        continue;
      }
    }

    return updatedCount;
  }

  /**
   * Remove all rows matching the filter
   * Returns the number of rows removed
   */
  removeWhere(filter: Partial<TItem>): number {
    const table = this.getTable();
    let removedCount = 0;

    // Collect rows to remove (in reverse order to avoid index shifting issues)
    const rowsToRemove: number[] = [];

    for (const { rowIndex, isEmpty } of iterateTableRows(this.engine, table)) {
      if (isEmpty) continue;

      try {
        const item = rowToObject<TItem>(
          this.engine,
          table,
          rowIndex,
          this.headers,
          (cell) => this.engine.getCellMetadata(cell)
        );

        if (matchesFilter(item, filter)) {
          rowsToRemove.push(rowIndex);
        }
      } catch {
        // Skip rows that fail parsing
        continue;
      }
    }

    // Remove rows in reverse order
    for (const rowIndex of rowsToRemove.reverse()) {
      // Clear all cells in the row
      for (const header of Object.values(this.headers)) {
        const cellAddress: CellAddress = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          colIndex: table.start.colIndex + header.index,
          rowIndex,
        };

        this.engine.setCellContent(cellAddress, undefined);
      }

      removedCount++;
    }

    return removedCount;
  }

  /**
   * Get the count of data rows in the table
   */
  count(): number {
    const table = this.getTable();
    return getTableRowCount(this.engine, table);
  }
}
