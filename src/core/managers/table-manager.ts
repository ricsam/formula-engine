import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRange,
  SpreadsheetRangeEnd,
  TableDefinition,
} from "../types";
import type { TableManagerSnapshot } from "../engine-snapshot";
import {
  checkRangeIntersection,
  getCellReference,
  parseCellReference,
} from "../utils";
import type { WorkbookManager } from "./workbook-manager";

export type TableHeaderUpdate = {
  table: TableDefinition;
  index: number;
  oldName: string;
  newName: string;
};

export type GeneratedTableHeader = {
  address: CellAddress;
  name: string;
};

export class TableManager {
  tables: Map<
    /**
     * workbook name -> table name -> table definition
     */
    string,
    Map<string, TableDefinition>
  > = new Map();
  private workbookManager: WorkbookManager;

  constructor(workbookManager: WorkbookManager) {
    this.workbookManager = workbookManager;
  }

  getTables(workbookName: string): Map<string, TableDefinition> {
    return this.tables.get(workbookName) ?? new Map();
  }

  getTable(opts: {
    workbookName: string;
    name: string;
  }): TableDefinition | undefined {
    return this.tables.get(opts.workbookName)?.get(opts.name);
  }

  private getDefaultHeaderName(index: number, usedNames: Set<string>): string {
    let number = index + 1;
    let name = `Column ${number}`;
    while (usedNames.has(name)) {
      number++;
      name = `Column ${number}`;
    }
    return name;
  }

  private getHeaderName(
    value: SerializedCellValue,
    index: number,
    usedNames: Set<string>
  ): string {
    if (value === undefined || value === "") {
      return this.getDefaultHeaderName(index, usedNames);
    }
    return String(value);
  }

  private replaceHeader(
    table: TableDefinition,
    index: number,
    newName: string
  ): void {
    const headers = new Map<string, { name: string; index: number }>();
    for (const header of Array.from(table.headers.values()).sort(
      (a, b) => a.index - b.index
    )) {
      const name = header.index === index ? newName : header.name;
      headers.set(name, { name, index: header.index });
    }
    table.headers = headers;
  }

  prepareHeaderUpdate(
    address: CellAddress,
    value: SerializedCellValue
  ): { content: SerializedCellValue; updates: TableHeaderUpdate[] } {
    const updates: TableHeaderUpdate[] = [];

    for (const table of this.getTables(address.workbookName).values()) {
      if (
        table.sheetName !== address.sheetName ||
        address.rowIndex !== table.start.rowIndex
      ) {
        continue;
      }

      const index = address.colIndex - table.start.colIndex;
      if (index < 0 || index >= table.headers.size) {
        continue;
      }

      const oldHeader = Array.from(table.headers.values()).find(
        (header) => header.index === index
      );
      if (!oldHeader) {
        continue;
      }

      const usedNames = new Set(
        Array.from(table.headers.values())
          .filter((header) => header.index !== index)
          .map((header) => header.name)
      );
      const newName = this.getHeaderName(value, index, usedNames);
      if (usedNames.has(newName)) {
        throw new Error(`Duplicate table header "${newName}"`);
      }

      updates.push({
        table,
        index,
        oldName: oldHeader.name,
        newName,
      });
    }

    const generatedNames = new Set(
      updates
        .filter(() => value === undefined || value === "")
        .map((update) => update.newName)
    );
    if (generatedNames.size > 1) {
      throw new Error("Overlapping tables require the same header name");
    }

    return {
      content:
        value === undefined || value === ""
          ? updates[0]?.newName ?? value
          : value,
      updates,
    };
  }

  prepareHeaderUpdatesForSheet(options: {
    workbookName: string;
    sheetName: string;
    getCellContent: (address: CellAddress) => SerializedCellValue;
  }): {
    updates: TableHeaderUpdate[];
    generatedHeaders: GeneratedTableHeader[];
  } {
    const updates: TableHeaderUpdate[] = [];
    const generatedHeaders = new Map<string, GeneratedTableHeader>();

    for (const table of this.getTables(options.workbookName).values()) {
      if (table.sheetName !== options.sheetName) {
        continue;
      }

      const oldHeadersByIndex = new Map(
        Array.from(table.headers.values()).map((header) => [
          header.index,
          header,
        ])
      );
      const usedNames = new Set<string>();
      for (let index = 0; index < table.headers.size; index++) {
        const address = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          rowIndex: table.start.rowIndex,
          colIndex: table.start.colIndex + index,
        };
        const value = options.getCellContent(address);
        const newName = this.getHeaderName(value, index, usedNames);
        if (usedNames.has(newName)) {
          throw new Error(`Duplicate table header "${newName}"`);
        }
        usedNames.add(newName);

        const oldHeader = oldHeadersByIndex.get(index);
        if (!oldHeader) {
          continue;
        }
        updates.push({
          table,
          index,
          oldName: oldHeader.name,
          newName,
        });

        if (value === undefined || value === "") {
          const key = `${address.workbookName}:${address.sheetName}:${address.rowIndex}:${address.colIndex}`;
          const existing = generatedHeaders.get(key);
          if (existing && existing.name !== newName) {
            throw new Error("Overlapping tables require the same header name");
          }
          generatedHeaders.set(key, { address, name: newName });
        }
      }
    }

    return {
      updates,
      generatedHeaders: Array.from(generatedHeaders.values()),
    };
  }

  applyHeaderUpdates(updates: TableHeaderUpdate[]): void {
    for (const update of updates) {
      this.replaceHeader(update.table, update.index, update.newName);
    }
  }

  makeTable({
    tableName,
    sheetName,
    workbookName,
    start,
    numRows,
    numCols,
    getCellValue,
  }: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    workbookName: string;
    getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  }): TableDefinition {
    const { rowIndex, colIndex } = parseCellReference(start);

    const headers = new Map<string, { name: string; index: number }>();
    const usedNames = new Set<string>();
    for (let i = 0; i < numCols; i++) {
      const header = getCellValue({
        rowIndex,
        colIndex: colIndex + i,
        sheetName,
        workbookName,
      });

      const name = this.getHeaderName(header, i, usedNames);
      if (usedNames.has(name)) {
        throw new Error(`Duplicate table header "${name}"`);
      }
      usedNames.add(name);
      headers.set(name, { name, index: i });
    }

    const endRow: SpreadsheetRangeEnd =
      numRows.type === "number"
        ? { type: "number", value: rowIndex + numRows.value }
        : numRows;

    const table: TableDefinition = {
      name: tableName,
      sheetName,
      workbookName,
      start: {
        rowIndex,
        colIndex,
      },
      headers,
      endRow,
    };

    return table;
  }

  copyTable(
    from: {
      workbookName: string;
      tableName: string;
    },
    to: {
      workbookName: string;
      tableName: string;
    },
  ): void {
    const fromTable = this.getTable({
      workbookName: from.workbookName,
      name: from.tableName,
    });
    if (!fromTable) {
      throw new Error("Table not found");
    }
    const wb = this.tables.get(to.workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }
    const newTable: TableDefinition = {
      ...fromTable,
      workbookName: to.workbookName,
    };
    wb.set(to.tableName, newTable);
  }

  addTable(props: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    workbookName: string;
    getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  }): TableDefinition {
    const tableName = props.tableName;
    const table = this.makeTable(props);

    let wb = this.tables.get(props.workbookName);
    if (!wb) {
      wb = new Map();
      this.tables.set(props.workbookName, wb);
    }

    wb.set(tableName, table);

    return table;
  }

  renameTable(
    workbookName: string,
    names: { oldName: string; newName: string },
  ): void {
    const wb = this.tables.get(workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }
    const table = wb.get(names.oldName);
    if (!table) {
      throw new Error("Table not found");
    }
    table.name = names.newName;
    wb.set(names.newName, table);
    wb.delete(names.oldName);
  }

  updateTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
    workbookName,
    getCellValue,
  }: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    workbookName: string;
    numCols?: number;
    getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  }): void {
    const wb = this.tables.get(workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }

    const table = wb.get(tableName);
    if (!table) {
      throw new Error("Table not found");
    }

    const newStart = start ? parseCellReference(start) : table.start;

    let newNumRows: SpreadsheetRangeEnd;
    if (numRows) {
      newNumRows = numRows;
    } else {
      if (table.endRow.type === "infinity") {
        newNumRows = table.endRow;
      } else {
        newNumRows = {
          type: "number",
          value: table.endRow.value - newStart.rowIndex,
        };
      }
    }

    const newTable = this.makeTable({
      tableName,
      sheetName: sheetName ?? table.sheetName,
      workbookName: workbookName ?? table.workbookName,
      start: getCellReference(newStart),
      numRows: newNumRows,
      numCols: numCols ?? table.headers.size,
      getCellValue,
    });

    wb.set(tableName, newTable);
  }

  removeTable({
    tableName,
    workbookName,
  }: {
    tableName: string;
    workbookName: string;
  }): boolean {
    const wb = this.tables.get(workbookName);
    if (!wb) {
      return false;
    }
    const found = wb.delete(tableName);

    return found;
  }

  updateTablesForSheetRename(options: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    // Update tables that belong to the renamed sheet
    const wb = this.tables.get(options.workbookName);
    if (!wb) {
      // No tables exist for this workbook yet — nothing to update
      return;
    }

    wb.forEach((table) => {
      if (table.sheetName === options.sheetName) {
        table.sheetName = options.newSheetName;
      }
    });
  }

  updateTablesForWorkbookRename(options: {
    workbookName: string;
    newWorkbookName: string;
  }): void {
    const wb = this.tables.get(options.workbookName);
    if (!wb) {
      // No tables exist for this workbook yet — nothing to update
      return;
    }
    this.tables.set(options.newWorkbookName, wb);
    this.tables.delete(options.workbookName);
    // Update tables that belong to the renamed sheet
    wb.forEach((table, tableName) => {
      if (table.workbookName === options.workbookName) {
        table.workbookName = options.newWorkbookName;
      }
    });
  }

  resetTables(newTables: Map<string, Map<string, TableDefinition>>): void {
    // Clear existing tables without breaking the Map reference
    this.tables.clear();

    // Repopulate with new tables
    newTables.forEach((table, workbookName) => {
      table.forEach((table, tableName) => {
        let wb = this.tables.get(workbookName);
        if (!wb) {
          wb = new Map();
          this.tables.set(workbookName, wb);
        }
        wb.set(tableName, table);
      });
    });
  }

  toSnapshot(): TableManagerSnapshot {
    return this.tables;
  }

  restoreFromSnapshot(snapshot: TableManagerSnapshot): void {
    this.resetTables(snapshot);
  }

  /**
   * When adding a workbook, we need to initialize the new maps
   */
  addWorkbook(workbookName: string) {
    this.tables.set(workbookName, new Map());
  }

  /**
   * When removing a workbook, we need to remove the maps
   */
  removeWorkbook(workbookName: string) {
    this.tables.delete(workbookName);
  }

  /**
   * When removing a sheet, we need to remove the tables that belong to the sheet
   */
  removeSheet(opts: { sheetName: string; workbookName: string }): void {
    // Remove tables that belong to the removed sheet
    const wb = this.tables.get(opts.workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }
    wb.forEach((table, tableName) => {
      if (table.sheetName === opts.sheetName) {
        wb.delete(tableName);
      }
    });
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    const { rowIndex, colIndex } = cellAddress;

    // Get all tables for this sheet

    for (const table of this.getTables(cellAddress.workbookName).values()) {
      // Check each table to see if the cell is within its bounds
      if (table.sheetName !== cellAddress.sheetName) {
        continue;
      }

      const { start, endRow, headers } = table;

      // Check row bounds
      const isInRowRange =
        endRow.type === "infinity"
          ? rowIndex >= start.rowIndex
          : rowIndex >= start.rowIndex && rowIndex <= endRow.value;

      // Check column bounds
      const endColIndex = start.colIndex + headers.size - 1;
      const isInColRange =
        colIndex >= start.colIndex && colIndex <= endColIndex;

      if (isInRowRange && isInColRange) {
        return table;
      }
    }

    return undefined;
  }

  /**
   * Check if a range intersects with any table in the given workbook/sheet.
   * Used to prevent spilling into tables (Excel behavior).
   */
  doesRangeIntersectTable(
    workbookName: string,
    sheetName: string,
    range: SpreadsheetRange,
  ): boolean {
    for (const table of this.getTables(workbookName).values()) {
      if (table.sheetName !== sheetName) {
        continue;
      }

      // Build the table's range
      const { start, endRow, headers } = table;
      const endColIndex = start.colIndex + headers.size - 1;

      const tableRange: SpreadsheetRange = {
        start: { col: start.colIndex, row: start.rowIndex },
        end: {
          col: { type: "number", value: endColIndex },
          row: endRow,
        },
      };

      if (checkRangeIntersection(range, tableRange)) {
        return true;
      }
    }

    return false;
  }
}
