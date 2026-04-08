/**
 * Table Commands - Commands that modify table definitions
 *
 * These commands all require re-evaluation after execution.
 */

import type { TableManager } from "../managers/table-manager";
import type { NamedExpressionManager } from "../managers/named-expression-manager";
import type { WorkbookManager } from "../managers/workbook-manager";
import type { SchemaManager } from "../managers/schema-manager";
import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRangeEnd,
  TableDefinition,
} from "../types";
import type {
  EngineCommand,
  EngineAction,
  MutationInvalidation,
} from "./types";
import {
  ActionTypes,
  emptyMutationInvalidation,
  getSerializedCellValueKind,
} from "./types";
import { getTableResourceKey } from "../resource-keys";
import { parseCellReference } from "../utils";

function getAddressKey(address: CellAddress): string {
  return `${address.workbookName}:${address.sheetName}:${address.rowIndex}:${address.colIndex}`;
}

function collectTableFootprintCells(
  workbookManager: WorkbookManager,
  table: TableDefinition
): Array<{
  address: CellAddress;
  content: SerializedCellValue | undefined;
}> {
  const cells = new Map<
    string,
    {
      address: CellAddress;
      content: SerializedCellValue | undefined;
    }
  >();
  const sheet = workbookManager.getSheet({
    workbookName: table.workbookName,
    sheetName: table.sheetName,
  });
  if (!sheet) {
    return [];
  }

  const startColIndex = table.start.colIndex;
  const endColIndex = startColIndex + table.headers.size - 1;

  if (table.endRow.type === "number") {
    for (let rowIndex = table.start.rowIndex; rowIndex <= table.endRow.value; rowIndex++) {
      for (let colIndex = startColIndex; colIndex <= endColIndex; colIndex++) {
        const address = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          rowIndex,
          colIndex,
        };
        cells.set(getAddressKey(address), {
          address,
          content: workbookManager.getCellContent(address),
        });
      }
    }
    return Array.from(cells.values());
  }

  for (const [ref, content] of sheet.content.entries()) {
    const { rowIndex, colIndex } = parseCellReference(ref);
    if (rowIndex < table.start.rowIndex) {
      continue;
    }
    if (colIndex < startColIndex || colIndex > endColIndex) {
      continue;
    }

    const address = {
      workbookName: table.workbookName,
      sheetName: table.sheetName,
      rowIndex,
      colIndex,
    };
    cells.set(getAddressKey(address), {
      address,
      content,
    });
  }

  return Array.from(cells.values());
}

function buildTableTouchedCells(
  workbookManager: WorkbookManager,
  tables: Array<TableDefinition | undefined>
): MutationInvalidation["touchedCells"] {
  const touchedCells = new Map<
    string,
    {
      address: CellAddress;
      beforeKind: ReturnType<typeof getSerializedCellValueKind>;
      afterKind: ReturnType<typeof getSerializedCellValueKind>;
    }
  >();

  for (const table of tables) {
    if (!table) {
      continue;
    }
    for (const cell of collectTableFootprintCells(workbookManager, table)) {
      touchedCells.set(getAddressKey(cell.address), {
        address: cell.address,
        beforeKind: getSerializedCellValueKind(cell.content),
        afterKind: getSerializedCellValueKind(cell.content),
      });
    }
  }

  return Array.from(touchedCells.values());
}

function buildTableContextChangedCells(
  workbookManager: WorkbookManager,
  tables: Array<TableDefinition | undefined>
): CellAddress[] {
  const changedCells = new Map<string, CellAddress>();

  for (const table of tables) {
    if (!table) {
      continue;
    }
    for (const cell of collectTableFootprintCells(workbookManager, table)) {
      changedCells.set(getAddressKey(cell.address), cell.address);
    }
  }

  return Array.from(changedCells.values());
}

function mergeTouchedCells(
  ...groups: MutationInvalidation["touchedCells"][]
): MutationInvalidation["touchedCells"] {
  const precedence = {
    empty: 0,
    scalar: 1,
    formula: 2,
  } as const;
  const merged = new Map<
    string,
    MutationInvalidation["touchedCells"][number]
  >();

  for (const group of groups) {
    for (const touchedCell of group) {
      const key = getAddressKey(touchedCell.address);
      const existing = merged.get(key);
      if (!existing) {
        merged.set(key, touchedCell);
        continue;
      }

      merged.set(key, {
        address: touchedCell.address,
        beforeKind:
          precedence[touchedCell.beforeKind] >= precedence[existing.beforeKind]
            ? touchedCell.beforeKind
            : existing.beforeKind,
        afterKind:
          precedence[touchedCell.afterKind] >= precedence[existing.afterKind]
            ? touchedCell.afterKind
            : existing.afterKind,
      });
    }
  }

  return Array.from(merged.values());
}

/**
 * Dependencies needed for table commands.
 */
export interface TableCommandDeps {
  tableManager: TableManager;
  namedExpressionManager: NamedExpressionManager;
  workbookManager: WorkbookManager;
  apiSchemaManager: SchemaManager;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

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
    const table = this.deps.tableManager.addTable({
      ...this.props,
      getCellValue: this.deps.getCellValue,
    });
    const resourceKey = getTableResourceKey({
      workbookName: this.props.workbookName,
      tableName: this.props.tableName,
    });
    this.executeFootprint = {
      touchedCells: buildTableTouchedCells(this.deps.workbookManager, [table]),
      tableContextChangedCells: buildTableContextChangedCells(
        this.deps.workbookManager,
        [table]
      ),
      resourceKeys: [resourceKey],
    };
    this.undoFootprint = this.executeFootprint;
  }

  undo(): void {
    this.deps.tableManager.removeTable({
      workbookName: this.props.workbookName,
      tableName: this.props.tableName,
    });
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

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
    const resourceKey = getTableResourceKey({
      workbookName: this.opts.workbookName,
      tableName: this.opts.tableName,
    });
    this.executeFootprint = {
      touchedCells: buildTableTouchedCells(this.deps.workbookManager, [
        this.removedTable,
      ]),
      tableContextChangedCells: buildTableContextChangedCells(
        this.deps.workbookManager,
        [this.removedTable]
      ),
      resourceKeys: [resourceKey],
    };
    this.undoFootprint = this.executeFootprint;
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

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: TableCommandDeps,
    private workbookName: string,
    private oldName: string,
    private newName: string
  ) {}

  execute(): void {
    const previousTable = this.deps.tableManager.getTable({
      workbookName: this.workbookName,
      name: this.oldName,
    });
    this.deps.tableManager.renameTable(this.workbookName, {
      oldName: this.oldName,
      newName: this.newName,
    });

    // Update formulas in sheet cells
    const changedCells = this.deps.workbookManager.updateAllFormulas((formula) =>
      this.deps.renameTableInFormula(formula, this.oldName, this.newName)
    );

    // Update named expressions
    const changedNamedExpressions =
      this.deps.namedExpressionManager.updateAllNamedExpressions((formula) =>
      this.deps.renameTableInFormula(formula, this.oldName, this.newName)
    );

    // Update API schemas
    this.deps.apiSchemaManager.updateForTableRename(
      this.workbookName,
      this.oldName,
      this.newName
    );
    const renamedTable = this.deps.tableManager.getTable({
      workbookName: this.workbookName,
      name: this.newName,
    });

    this.executeFootprint = {
      touchedCells: mergeTouchedCells(
        buildTableTouchedCells(this.deps.workbookManager, [
          previousTable,
          renamedTable,
        ]),
        changedCells.map((address) => ({
          address,
          beforeKind: "formula" as const,
          afterKind: "formula" as const,
        }))
      ),
      tableContextChangedCells: buildTableContextChangedCells(
        this.deps.workbookManager,
        [previousTable, renamedTable]
      ),
      resourceKeys: [
        getTableResourceKey({
          workbookName: this.workbookName,
          tableName: this.oldName,
        }),
        getTableResourceKey({
          workbookName: this.workbookName,
          tableName: this.newName,
        }),
        ...changedNamedExpressions,
      ],
    };
  }

  undo(): void {
    const currentTable = this.deps.tableManager.getTable({
      workbookName: this.workbookName,
      name: this.newName,
    });
    // Rename back
    this.deps.tableManager.renameTable(this.workbookName, {
      oldName: this.newName,
      newName: this.oldName,
    });

    // Update formulas back
    const changedCells = this.deps.workbookManager.updateAllFormulas((formula) =>
      this.deps.renameTableInFormula(formula, this.newName, this.oldName)
    );

    // Update named expressions back
    const changedNamedExpressions =
      this.deps.namedExpressionManager.updateAllNamedExpressions((formula) =>
      this.deps.renameTableInFormula(formula, this.newName, this.oldName)
    );

    // Update API schemas back
    this.deps.apiSchemaManager.updateForTableRename(
      this.workbookName,
      this.newName,
      this.oldName
    );
    const restoredTable = this.deps.tableManager.getTable({
      workbookName: this.workbookName,
      name: this.oldName,
    });

    this.undoFootprint = {
      touchedCells: mergeTouchedCells(
        buildTableTouchedCells(this.deps.workbookManager, [
          currentTable,
          restoredTable,
        ]),
        changedCells.map((address) => ({
          address,
          beforeKind: "formula" as const,
          afterKind: "formula" as const,
        }))
      ),
      tableContextChangedCells: buildTableContextChangedCells(
        this.deps.workbookManager,
        [currentTable, restoredTable]
      ),
      resourceKeys: [
        getTableResourceKey({
          workbookName: this.workbookName,
          tableName: this.oldName,
        }),
        getTableResourceKey({
          workbookName: this.workbookName,
          tableName: this.newName,
        }),
        ...changedNamedExpressions,
      ],
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

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
    const resourceKey = getTableResourceKey({
      workbookName: this.opts.workbookName,
      tableName: this.opts.tableName,
    });
    const nextTable = this.deps.tableManager.getTable({
      workbookName: this.opts.workbookName,
      name: this.opts.tableName,
    });
    this.executeFootprint = {
      touchedCells: buildTableTouchedCells(this.deps.workbookManager, [
        this.previousTable,
        nextTable,
      ]),
      tableContextChangedCells: buildTableContextChangedCells(
        this.deps.workbookManager,
        [this.previousTable, nextTable]
      ),
      resourceKeys: [resourceKey],
    };
    this.undoFootprint = this.executeFootprint;
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

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

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
    const resourceKeys = new Set<string>();
    for (const [workbookName, tables] of this.previousTables ?? []) {
      for (const tableName of tables.keys()) {
        resourceKeys.add(
          getTableResourceKey({
            workbookName,
            tableName,
          })
        );
      }
    }
    for (const [workbookName, tables] of this.newTables) {
      for (const tableName of tables.keys()) {
        resourceKeys.add(
          getTableResourceKey({
            workbookName,
            tableName,
          })
        );
      }
    }
    this.executeFootprint = {
      touchedCells: buildTableTouchedCells(this.deps.workbookManager, [
        ...Array.from(this.previousTables?.values() ?? []).flatMap((tables) =>
          Array.from(tables.values())
        ),
        ...Array.from(this.newTables.values()).flatMap((tables) =>
          Array.from(tables.values())
        ),
      ]),
      tableContextChangedCells: buildTableContextChangedCells(
        this.deps.workbookManager,
        [
          ...Array.from(this.previousTables?.values() ?? []).flatMap((tables) =>
            Array.from(tables.values())
          ),
          ...Array.from(this.newTables.values()).flatMap((tables) =>
            Array.from(tables.values())
          ),
        ]
      ),
      resourceKeys: Array.from(resourceKeys),
    };
    this.undoFootprint = this.executeFootprint;
  }

  undo(): void {
    if (!this.previousTables) return;
    this.deps.tableManager.resetTables(this.previousTables);
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
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
