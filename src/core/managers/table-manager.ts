import { renameTableInFormula } from "../table-renamer";
import type { SpreadsheetRangeEnd, TableDefinition } from "../types";
import { getCellReference, parseCellReference } from "../utils";
import type { EventManager } from "./event-manager";
import type { WorkbookManager } from "./workbook-manager";

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

  makeTable({
    tableName,
    sheetName,
    workbookName,
    start,
    numRows,
    numCols,
  }: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    workbookName: string;
  }): TableDefinition {
    const { rowIndex, colIndex } = parseCellReference(start);

    const sheet = this.workbookManager.getSheet({
      workbookName,
      sheetName,
    });
    if (!sheet) {
      throw new Error("Sheet not found");
    }
    const sheetContent = sheet.content;

    const headers = new Map<string, { name: string; index: number }>();
    for (let i = 0; i < numCols; i++) {
      const header = sheetContent.get(
        getCellReference({ rowIndex, colIndex: colIndex + i })
      );

      if (header) {
        headers.set(String(header), { name: String(header), index: i });
      } else {
        headers.set(`Column ${i + 1}`, { name: `Column ${i + 1}`, index: i });
      }
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

  addTable(props: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    workbookName: string;
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
    names: { oldName: string; newName: string }
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
  }: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    workbookName: string;
    numCols?: number;
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
    this.tables.forEach((wb, workbookName) => {
      wb.forEach((table, tableName) => {
        if (
          table.sheetName === options.sheetName &&
          table.workbookName === options.workbookName
        ) {
          table.sheetName = options.newSheetName;
        }
      });
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
}
