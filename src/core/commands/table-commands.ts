/**
 * Table Commands - Commands that modify table definitions
 *
 * These commands all require re-evaluation after execution.
 */

import type { TableManager } from "../managers/table-manager";
import type { NamedExpressionManager } from "../managers/named-expression-manager";
import type { WorkbookManager } from "../managers/workbook-manager";
import type { ApiSchemaManager } from "../managers/api-schema-manager";
import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRangeEnd,
  TableDefinition,
} from "../types";
import type { EngineCommand, EngineAction } from "./types";
import { ActionTypes } from "./types";

/**
 * Dependencies needed for table commands.
 */
export interface TableCommandDeps {
  tableManager: TableManager;
  namedExpressionManager: NamedExpressionManager;
  workbookManager: WorkbookManager;
  apiSchemaManager: ApiSchemaManager;
  getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  renameTableInFormula: (
    formula: string,
    oldName: string,
    newName: string
  ) => string;
}

/**
 * Command to add a table.
 */
export class AddTableCommand implements EngineCommand {
  readonly requiresReevaluation = true;

  constructor(
    private deps: TableCommandDeps,
    private props: {
      tableName: string;
      sheetName: string;
      workbookName: string;
      start: string;
      numRows: SpreadsheetRangeEnd;
      numCols: number;
    }
  ) {}

  execute(): void {
    this.deps.tableManager.addTable({
      ...this.props,
      getCellValue: this.deps.getCellValue,
    });
  }

  undo(): void {
    this.deps.tableManager.removeTable({
      workbookName: this.props.workbookName,
      tableName: this.props.tableName,
    });
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.ADD_TABLE,
      payload: this.props,
    };
  }
}

/**
 * Command to remove a table.
 */
export class RemoveTableCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private removedTable: TableDefinition | undefined;

  constructor(
    private deps: TableCommandDeps,
    private opts: { tableName: string; workbookName: string }
  ) {}

  execute(): void {
    // Capture table before removal
    this.removedTable = this.deps.tableManager.getTable({
      workbookName: this.opts.workbookName,
      name: this.opts.tableName,
    });

    this.deps.tableManager.removeTable(this.opts);
  }

  undo(): void {
    if (!this.removedTable) return;

    // Recreate the table
    const { start, endRow, headers, sheetName } = this.removedTable;
    const startRef = `${String.fromCharCode(65 + start.colIndex)}${start.rowIndex + 1}`;

    this.deps.tableManager.addTable({
      workbookName: this.opts.workbookName,
      tableName: this.opts.tableName,
      sheetName,
      start: startRef,
      numRows: endRow,
      numCols: headers.size,
      getCellValue: this.deps.getCellValue,
    });
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.REMOVE_TABLE,
      payload: this.opts,
    };
  }
}

/**
 * Command to rename a table.
 */
export class RenameTableCommand implements EngineCommand {
  readonly requiresReevaluation = true;

  constructor(
    private deps: TableCommandDeps,
    private workbookName: string,
    private oldName: string,
    private newName: string
  ) {}

  execute(): void {
    this.deps.tableManager.renameTable(this.workbookName, {
      oldName: this.oldName,
      newName: this.newName,
    });

    // Update formulas in sheet cells
    this.deps.workbookManager.updateAllFormulas((formula) =>
      this.deps.renameTableInFormula(formula, this.oldName, this.newName)
    );

    // Update named expressions
    this.deps.namedExpressionManager.updateAllNamedExpressions((formula) =>
      this.deps.renameTableInFormula(formula, this.oldName, this.newName)
    );

    // Update API schemas
    this.deps.apiSchemaManager.updateForTableRename(
      this.workbookName,
      this.oldName,
      this.newName
    );
  }

  undo(): void {
    // Rename back
    this.deps.tableManager.renameTable(this.workbookName, {
      oldName: this.newName,
      newName: this.oldName,
    });

    // Update formulas back
    this.deps.workbookManager.updateAllFormulas((formula) =>
      this.deps.renameTableInFormula(formula, this.newName, this.oldName)
    );

    // Update named expressions back
    this.deps.namedExpressionManager.updateAllNamedExpressions((formula) =>
      this.deps.renameTableInFormula(formula, this.newName, this.oldName)
    );

    // Update API schemas back
    this.deps.apiSchemaManager.updateForTableRename(
      this.workbookName,
      this.newName,
      this.oldName
    );
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.RENAME_TABLE,
      payload: {
        workbookName: this.workbookName,
        oldName: this.oldName,
        newName: this.newName,
      },
    };
  }
}

/**
 * Command to update a table.
 */
export class UpdateTableCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousTable: TableDefinition | undefined;

  constructor(
    private deps: TableCommandDeps,
    private opts: {
      tableName: string;
      sheetName?: string;
      start?: string;
      numRows?: SpreadsheetRangeEnd;
      numCols?: number;
      workbookName: string;
    }
  ) {}

  execute(): void {
    // Capture previous table state
    this.previousTable = this.deps.tableManager.getTable({
      workbookName: this.opts.workbookName,
      name: this.opts.tableName,
    });

    this.deps.tableManager.updateTable({
      ...this.opts,
      getCellValue: this.deps.getCellValue,
    });
  }

  undo(): void {
    if (!this.previousTable) return;

    // Restore previous table state
    const { start, endRow, headers, sheetName } = this.previousTable;
    const startRef = `${String.fromCharCode(65 + start.colIndex)}${start.rowIndex + 1}`;

    this.deps.tableManager.updateTable({
      workbookName: this.opts.workbookName,
      tableName: this.opts.tableName,
      sheetName,
      start: startRef,
      numRows: endRow,
      numCols: headers.size,
      getCellValue: this.deps.getCellValue,
    });
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.UPDATE_TABLE,
      payload: this.opts,
    };
  }
}

/**
 * Command to reset all tables.
 */
export class ResetTablesCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private previousTables: Map<string, Map<string, TableDefinition>> | undefined;

  constructor(
    private deps: TableCommandDeps,
    private newTables: Map<string, Map<string, TableDefinition>>
  ) {}

  execute(): void {
    // Capture previous tables
    this.previousTables = new Map();
    for (const [workbookName, tables] of this.deps.tableManager.tables) {
      this.previousTables.set(workbookName, new Map(tables));
    }

    this.deps.tableManager.resetTables(this.newTables);
  }

  undo(): void {
    if (!this.previousTables) return;
    this.deps.tableManager.resetTables(this.previousTables);
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.RESET_TABLES,
      payload: {
        tables: Array.from(this.newTables.entries()).map(([wb, tables]) => [
          wb,
          Array.from(tables.entries()),
        ]),
      },
    };
  }
}

