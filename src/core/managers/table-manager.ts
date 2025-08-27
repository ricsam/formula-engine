import type {
  FormulaEngineEvents,
  SpreadsheetRangeEnd,
  TableDefinition,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";
import { renameTableInFormula } from "../table-renamer";

export class TableManager {
  private tables: Map<string, TableDefinition> = new Map();
  private eventEmitter?: {
    emit<K extends keyof FormulaEngineEvents>(
      event: K,
      data: FormulaEngineEvents[K]
    ): void;
  };

  constructor(eventEmitter?: {
    emit<K extends keyof FormulaEngineEvents>(
      event: K,
      data: FormulaEngineEvents[K]
    ): void;
  }) {
    this.eventEmitter = eventEmitter;
  }

  getTables(): Map<string, TableDefinition> {
    return this.tables;
  }

  getTable(name: string): TableDefinition | undefined {
    return this.tables.get(name);
  }

  makeTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
    getSheetContent,
  }: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    getSheetContent: (sheetName: string) => Map<string, any> | undefined;
  }): TableDefinition {
    const { rowIndex, colIndex } = parseCellReference(start);
    const sheetContent = getSheetContent(sheetName);
    if (!sheetContent) {
      throw new Error("Sheet not found");
    }

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
    getSheetContent: (sheetName: string) => Map<string, any> | undefined;
  }): TableDefinition {
    const tableName = props.tableName;
    const table = this.makeTable(props);

    this.tables.set(tableName, table);
    this.eventEmitter?.emit("tables-updated", this.tables);

    return table;
  }

  renameTable(names: { oldName: string; newName: string }): void {
    const table = this.tables.get(names.oldName);
    if (!table) {
      throw new Error("Table not found");
    }
    table.name = names.newName;
    this.tables.set(names.newName, table);
    this.tables.delete(names.oldName);

    this.eventEmitter?.emit("tables-updated", this.tables);
  }

  updateFormulasForTableRename(
    oldName: string,
    newName: string,
    updateCallback: (formula: string) => string = (formula) =>
      renameTableInFormula(formula, oldName, newName)
  ): void {
    // This method will be called by the engine to update formulas
    // The actual formula updating logic will be handled by the engine
  }

  updateTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
    getSheetContent,
  }: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
    getSheetContent: (sheetName: string) => Map<string, any> | undefined;
  }): void {
    const table = this.tables.get(tableName);
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
      start: getCellReference(newStart),
      numRows: newNumRows,
      numCols: numCols ?? table.headers.size,
      getSheetContent,
    });

    this.tables.set(tableName, newTable);
    this.eventEmitter?.emit("tables-updated", this.tables);
  }

  removeTable({ tableName }: { tableName: string }): boolean {
    const found = this.tables.delete(tableName);

    if (found) {
      this.eventEmitter?.emit("tables-updated", this.tables);
    }

    return found;
  }

  getTablesSerialized(): Map<string, TableDefinition> {
    return this.tables;
  }



  updateTablesForSheetRename(oldName: string, newName: string): void {
    // Update tables that belong to the renamed sheet
    this.tables.forEach((table, tableName) => {
      if (table.sheetName === oldName) {
        table.sheetName = newName;
      }
    });
  }

  removeTablesForSheet(sheetName: string): void {
    // Remove tables that belong to the removed sheet
    const tablesToRemove: string[] = [];
    this.tables.forEach((table, tableName) => {
      if (table.sheetName === sheetName) {
        tablesToRemove.push(tableName);
      }
    });
    
    tablesToRemove.forEach(tableName => {
      this.tables.delete(tableName);
    });
    
    if (tablesToRemove.length > 0) {
      this.eventEmitter?.emit("tables-updated", this.tables);
    }
  }

  /**
   * Replace all tables with new ones (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setTables(newTables: Map<string, TableDefinition>): void {
    // Clear existing tables without breaking the Map reference
    this.tables.clear();
    
    // Repopulate with new tables
    newTables.forEach((table, tableName) => {
      this.tables.set(tableName, table);
    });
    
    this.eventEmitter?.emit("tables-updated", this.tables);
  }
}
