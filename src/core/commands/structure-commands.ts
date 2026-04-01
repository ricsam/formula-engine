/**
 * Structure Commands - Commands that modify workbook/sheet structure
 *
 * These commands all require re-evaluation after execution.
 */

import type { SchemaManager } from "../managers/schema-manager";
import type { NamedExpressionManager } from "../managers/named-expression-manager";
import type { ReferenceManager } from "../managers/reference-manager";
import type { StyleManager } from "../managers/style-manager";
import type { TableManager } from "../managers/table-manager";
import type { WorkbookManager } from "../managers/workbook-manager";
import type {
  CellAddress,
  ConditionalStyle,
  DirectCellStyle,
  NamedExpression,
  Sheet,
  TableDefinition,
  Workbook,
} from "../types";
import type {
  EngineAction,
  EngineCommand,
  MutationInvalidation,
} from "./types";
import { ActionTypes, emptyMutationInvalidation } from "./types";
import {
  getNamedExpressionResourceKey,
  getSheetResourceKey,
  getTableResourceKey,
  getWorkbookResourceKey,
} from "../resource-keys";

/**
 * Dependencies needed for structure commands.
 */
export interface StructureCommandDeps {
  workbookManager: WorkbookManager;
  namedExpressionManager: NamedExpressionManager;
  tableManager: TableManager;
  styleManager: StyleManager;
  referenceManager: ReferenceManager;
  apiSchemaManager: SchemaManager;
  renameSheetInFormula: (opts: {
    formula: string;
    oldSheetName: string;
    newSheetName: string;
  }) => string;
  renameWorkbookInFormula: (opts: {
    formula: string;
    oldWorkbookName: string;
    newWorkbookName: string;
  }) => string;
}

// ============================================================================
// Workbook Commands
// ============================================================================

/**
 * Command to add a workbook.
 */
export class AddWorkbookCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private workbookName: string
  ) {}

  execute(): void {
    this.deps.workbookManager.addWorkbook(this.workbookName);
    this.deps.namedExpressionManager.addWorkbook(this.workbookName);
    this.deps.tableManager.addWorkbook(this.workbookName);
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [getWorkbookResourceKey(this.workbookName)],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: [getWorkbookResourceKey(this.workbookName)],
      removedScopes: [{ type: "workbook", workbookName: this.workbookName }],
    };
  }

  undo(): void {
    this.deps.workbookManager.removeWorkbook(this.workbookName);
    this.deps.namedExpressionManager.removeWorkbook(this.workbookName);
    this.deps.tableManager.removeWorkbook(this.workbookName);
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.ADD_WORKBOOK,
      payload: { workbookName: this.workbookName },
    };
  }
}

/**
 * Captured state of a workbook for undo purposes.
 */
interface WorkbookSnapshot {
  workbook: Workbook;
  namedExpressions: {
    workbookLevel: Map<string, NamedExpression>;
    sheetLevel: Map<string, Map<string, NamedExpression>>;
  };
  tables: Map<string, TableDefinition>;
  conditionalStyles: ConditionalStyle[];
  cellStyles: DirectCellStyle[];
}

/**
 * Capture the full state of a workbook for later restoration.
 */
function captureWorkbookSnapshot(
  deps: StructureCommandDeps,
  workbookName: string
): WorkbookSnapshot | undefined {
  const workbook = deps.workbookManager.getWorkbooks().get(workbookName);
  if (!workbook) return undefined;

  // Deep clone workbook sheets
  const clonedSheets = new Map<string, Sheet>();
  for (const [name, sheet] of workbook.sheets) {
    clonedSheets.set(name, {
      name: sheet.name,
      index: sheet.index,
      content: new Map(sheet.content),
      metadata: new Map(sheet.metadata),
      sheetMetadata: sheet.sheetMetadata,
    });
  }

  const namedExpressions = deps.namedExpressionManager.getNamedExpressions();

  const snapshot: WorkbookSnapshot = {
    workbook: {
      name: workbook.name,
      sheets: clonedSheets,
      workbookMetadata: workbook.workbookMetadata,
    },
    namedExpressions: {
      workbookLevel: new Map(
        namedExpressions.workbookExpressions.get(workbookName) || []
      ),
      sheetLevel: new Map(),
    },
    tables: new Map(deps.tableManager.getTables(workbookName)),
    conditionalStyles: deps.styleManager
      .getAllConditionalStyles()
      .filter((s) => s.areas.some((a) => a.workbookName === workbookName)),
    cellStyles: deps.styleManager
      .getAllCellStyles()
      .filter((s) => s.areas.some((a) => a.workbookName === workbookName)),
  };

  // Capture sheet-level named expressions
  const sheetExpressions = namedExpressions.sheetExpressions.get(workbookName);
  if (sheetExpressions) {
    for (const [sheetName, expressions] of sheetExpressions) {
      snapshot.namedExpressions.sheetLevel.set(sheetName, new Map(expressions));
    }
  }

  return snapshot;
}

/**
 * Remove a workbook and all its associated data.
 */
function removeWorkbook(
  deps: StructureCommandDeps,
  workbookName: string
): void {
  deps.workbookManager.removeWorkbook(workbookName);
  deps.namedExpressionManager.removeWorkbook(workbookName);
  deps.tableManager.removeWorkbook(workbookName);
  deps.styleManager.removeWorkbookStyles(workbookName);
  deps.referenceManager.invalidateWorkbook(workbookName);
}

/**
 * Rename a workbook across all managers.
 */
function renameWorkbookAcrossManagers(
  deps: StructureCommandDeps,
  oldName: string,
  newName: string
): CellAddress[] {
  deps.workbookManager.renameWorkbook({
    workbookName: oldName,
    newWorkbookName: newName,
  });
  deps.namedExpressionManager.renameWorkbook({
    workbookName: oldName,
    newWorkbookName: newName,
  });
  deps.tableManager.updateTablesForWorkbookRename({
    workbookName: oldName,
    newWorkbookName: newName,
  });
  deps.styleManager.updateWorkbookName(oldName, newName);
  const changedCells = deps.workbookManager.updateAllFormulas((formula) =>
    deps.renameWorkbookInFormula({
      formula,
      oldWorkbookName: oldName,
      newWorkbookName: newName,
    })
  );
  deps.referenceManager.updateWorkbookName(oldName, newName);
  deps.apiSchemaManager.updateForWorkbookRename(oldName, newName);
  return changedCells;
}

/**
 * Rename a sheet across all managers.
 */
function renameSheetAcrossManagers(
  deps: StructureCommandDeps,
  opts: { workbookName: string; sheetName: string; newSheetName: string }
): CellAddress[] {
  deps.workbookManager.renameSheet(opts);
  deps.namedExpressionManager.renameSheet(opts);
  deps.tableManager.updateTablesForSheetRename(opts);
  deps.styleManager.updateSheetName(
    opts.workbookName,
    opts.sheetName,
    opts.newSheetName
  );
  const changedCells = deps.workbookManager.updateAllFormulas((formula) =>
    deps.renameSheetInFormula({
      formula,
      oldSheetName: opts.sheetName,
      newSheetName: opts.newSheetName,
    })
  );
  deps.referenceManager.updateSheetName(
    opts.workbookName,
    opts.sheetName,
    opts.newSheetName
  );
  deps.apiSchemaManager.updateForSheetRename(
    opts.workbookName,
    opts.sheetName,
    opts.newSheetName
  );
  return changedCells;
}

function getWorkbookSnapshotResourceKeys(
  workbookName: string,
  snapshot: WorkbookSnapshot | undefined
): string[] {
  const resourceKeys = new Set<string>([getWorkbookResourceKey(workbookName)]);
  if (!snapshot) {
    return Array.from(resourceKeys);
  }

  for (const tableName of snapshot.tables.keys()) {
    resourceKeys.add(getTableResourceKey({ workbookName, tableName }));
  }
  for (const name of snapshot.namedExpressions.workbookLevel.keys()) {
    resourceKeys.add(
      getNamedExpressionResourceKey({ expressionName: name, workbookName })
    );
  }
  for (const [sheetName, expressions] of snapshot.namedExpressions.sheetLevel) {
    resourceKeys.add(getSheetResourceKey({ workbookName, sheetName }));
    for (const name of expressions.keys()) {
      resourceKeys.add(
        getNamedExpressionResourceKey({
          expressionName: name,
          workbookName,
          sheetName,
        })
      );
    }
  }

  return Array.from(resourceKeys);
}

function getSheetSnapshotResourceKeys(
  workbookName: string,
  sheetName: string,
  snapshot: SheetSnapshot | undefined
): string[] {
  const resourceKeys = new Set<string>([
    getWorkbookResourceKey(workbookName),
    getSheetResourceKey({ workbookName, sheetName }),
  ]);
  if (!snapshot) {
    return Array.from(resourceKeys);
  }

  for (const tableName of snapshot.tables.keys()) {
    resourceKeys.add(getTableResourceKey({ workbookName, tableName }));
  }
  for (const name of snapshot.namedExpressions.keys()) {
    resourceKeys.add(
      getNamedExpressionResourceKey({
        expressionName: name,
        workbookName,
        sheetName,
      })
    );
  }

  return Array.from(resourceKeys);
}

/**
 * Restore tables from a snapshot.
 */
function restoreTables(
  tableManager: TableManager,
  tables: Map<string, TableDefinition>,
  workbookName: string
): void {
  for (const [tableName, table] of tables) {
    tableManager.addTable({
      workbookName,
      tableName,
      sheetName: table.sheetName,
      start: `${String.fromCharCode(65 + table.start.colIndex)}${
        table.start.rowIndex + 1
      }`,
      numRows: table.endRow,
      numCols: table.headers.size,
      getCellValue: () => undefined,
    });
  }
}

/**
 * Restore styles from a snapshot.
 */
function restoreStyles(
  styleManager: StyleManager,
  conditionalStyles: ConditionalStyle[],
  cellStyles: DirectCellStyle[]
): void {
  for (const style of conditionalStyles) {
    styleManager.addConditionalStyle(style);
  }
  for (const style of cellStyles) {
    styleManager.addCellStyle(style);
  }
}

/**
 * Command to remove a workbook.
 */
export class RemoveWorkbookCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private snapshot: WorkbookSnapshot | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private workbookName: string
  ) {}

  execute(): void {
    this.snapshot = captureWorkbookSnapshot(this.deps, this.workbookName);
    removeWorkbook(this.deps, this.workbookName);
    const resourceKeys = getWorkbookSnapshotResourceKeys(
      this.workbookName,
      this.snapshot
    );
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys,
      removedScopes: [{ type: "workbook", workbookName: this.workbookName }],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys,
    };
  }

  undo(): void {
    if (!this.snapshot) return;

    // Restore workbook structure
    this.deps.workbookManager.addWorkbook(this.workbookName);
    this.deps.namedExpressionManager.addWorkbook(this.workbookName);
    this.deps.tableManager.addWorkbook(this.workbookName);

    // Restore sheets
    for (const [sheetName, sheet] of this.snapshot.workbook.sheets) {
      this.deps.workbookManager.addSheet({
        workbookName: this.workbookName,
        sheetName,
      });
      this.deps.namedExpressionManager.addSheet({
        workbookName: this.workbookName,
        sheetName,
      });
      this.deps.workbookManager.setSheetContent(
        { workbookName: this.workbookName, sheetName },
        sheet.content
      );
    }

    // Restore named expressions
    for (const [name, expr] of this.snapshot.namedExpressions.workbookLevel) {
      this.deps.namedExpressionManager.addNamedExpression({
        expressionName: name,
        expression: expr.expression,
        workbookName: this.workbookName,
      });
    }
    for (const [sheetName, expressions] of this.snapshot.namedExpressions.sheetLevel) {
      for (const [name, expr] of expressions) {
        this.deps.namedExpressionManager.addNamedExpression({
          expressionName: name,
          expression: expr.expression,
          workbookName: this.workbookName,
          sheetName,
        });
      }
    }

    // Restore tables and styles
    restoreTables(this.deps.tableManager, this.snapshot.tables, this.workbookName);
    restoreStyles(this.deps.styleManager, this.snapshot.conditionalStyles, this.snapshot.cellStyles);
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.REMOVE_WORKBOOK,
      payload: { workbookName: this.workbookName },
    };
  }
}

/**
 * Command to rename a workbook.
 */
export class RenameWorkbookCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private snapshot: WorkbookSnapshot | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private workbookName: string,
    private newWorkbookName: string
  ) {}

  execute(): void {
    this.snapshot = captureWorkbookSnapshot(this.deps, this.workbookName);
    const changedCells = renameWorkbookAcrossManagers(
      this.deps,
      this.workbookName,
      this.newWorkbookName
    );
    this.executeFootprint = {
      touchedCells: changedCells.map((address) => ({
        address,
        beforeKind: "formula" as const,
        afterKind: "formula" as const,
      })),
      resourceKeys: Array.from(
        new Set([
          ...getWorkbookSnapshotResourceKeys(this.workbookName, this.snapshot),
          ...getWorkbookSnapshotResourceKeys(this.newWorkbookName, this.snapshot),
        ])
      ),
    };
  }

  undo(): void {
    const changedCells = renameWorkbookAcrossManagers(
      this.deps,
      this.newWorkbookName,
      this.workbookName
    );
    this.undoFootprint = {
      touchedCells: changedCells.map((address) => ({
        address,
        beforeKind: "formula" as const,
        afterKind: "formula" as const,
      })),
      resourceKeys: Array.from(
        new Set([
          ...getWorkbookSnapshotResourceKeys(this.workbookName, this.snapshot),
          ...getWorkbookSnapshotResourceKeys(this.newWorkbookName, this.snapshot),
        ])
      ),
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.RENAME_WORKBOOK,
      payload: {
        workbookName: this.workbookName,
        newWorkbookName: this.newWorkbookName,
      },
    };
  }
}

/**
 * Command to clone a workbook.
 * Undo simply removes the cloned workbook.
 */
export class CloneWorkbookCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private fromWorkbookName: string,
    private toWorkbookName: string
  ) {}

  execute(): void {
    const {
      workbookManager,
      namedExpressionManager,
      tableManager,
      styleManager,
      renameWorkbookInFormula,
    } = this.deps;

    // Check if source workbook exists
    const sourceWorkbook = workbookManager
      .getWorkbooks()
      .get(this.fromWorkbookName);
    if (!sourceWorkbook) {
      throw new Error(`Source workbook "${this.fromWorkbookName}" not found`);
    }

    // Check if target workbook name already exists
    if (workbookManager.getWorkbooks().has(this.toWorkbookName)) {
      throw new Error(
        `Target workbook "${this.toWorkbookName}" already exists`
      );
    }

    // Create new workbook
    workbookManager.addWorkbook(this.toWorkbookName);
    namedExpressionManager.addWorkbook(this.toWorkbookName);
    tableManager.addWorkbook(this.toWorkbookName);

    // Clone all sheets from source workbook
    for (const [sheetName, sheet] of sourceWorkbook.sheets) {
      // Add sheet to target workbook
      workbookManager.addSheet({
        workbookName: this.toWorkbookName,
        sheetName: sheetName,
      });
      namedExpressionManager.addSheet({
        workbookName: this.toWorkbookName,
        sheetName: sheetName,
      });

      // Copy all cell content
      workbookManager.setSheetContent(
        { workbookName: this.toWorkbookName, sheetName: sheetName },
        new Map(sheet.content)
      );

      // Copy all cell metadata
      const targetSheet = workbookManager.getSheet({
        workbookName: this.toWorkbookName,
        sheetName: sheetName,
      });
      if (targetSheet) {
        targetSheet.metadata = new Map(sheet.metadata);

        // Copy sheet metadata
        if (sheet.sheetMetadata !== undefined) {
          targetSheet.sheetMetadata = structuredClone(sheet.sheetMetadata);
        }
      }
    }

    // Copy workbook metadata
    const targetWorkbook = workbookManager
      .getWorkbooks()
      .get(this.toWorkbookName);
    if (targetWorkbook && sourceWorkbook.workbookMetadata !== undefined) {
      targetWorkbook.workbookMetadata = structuredClone(
        sourceWorkbook.workbookMetadata
      );
    }

    // Clone workbook-scoped named expressions
    const allExpressions = namedExpressionManager.getNamedExpressions();
    const sourceWorkbookExpressions = allExpressions.workbookExpressions.get(
      this.fromWorkbookName
    );
    if (sourceWorkbookExpressions) {
      for (const [expressionName, expression] of sourceWorkbookExpressions) {
        namedExpressionManager.addNamedExpression({
          expressionName,
          expression: expression.expression,
          workbookName: this.toWorkbookName,
        });
      }
    }

    // Clone sheet-scoped named expressions
    const sourceSheetExpressions = allExpressions.sheetExpressions.get(
      this.fromWorkbookName
    );
    if (sourceSheetExpressions) {
      for (const [sheetName, sheetExpressions] of sourceSheetExpressions) {
        for (const [expressionName, expression] of sheetExpressions) {
          namedExpressionManager.addNamedExpression({
            expressionName,
            expression: expression.expression,
            workbookName: this.toWorkbookName,
            sheetName,
          });
        }
      }
    }

    // Clone tables
    const sourceTables = tableManager.tables.get(this.fromWorkbookName);
    if (sourceTables) {
      for (const [tableName] of sourceTables) {
        tableManager.copyTable(
          { workbookName: this.fromWorkbookName, tableName },
          { workbookName: this.toWorkbookName, tableName }
        );
      }
    }

    // Clone conditional styles
    const allConditionalStyles = styleManager.getAllConditionalStyles();
    for (const style of allConditionalStyles) {
      if (
        style.areas.some((area) => area.workbookName === this.fromWorkbookName)
      ) {
        const newStyle: ConditionalStyle = {
          ...style,
          areas: style.areas.map((area) =>
            area.workbookName === this.fromWorkbookName
              ? { ...area, workbookName: this.toWorkbookName }
              : area
          ),
        };
        styleManager.addConditionalStyle(newStyle);
      }
    }

    // Clone cell styles
    const allCellStyles = styleManager.getAllCellStyles();
    for (const style of allCellStyles) {
      if (
        style.areas.some((area) => area.workbookName === this.fromWorkbookName)
      ) {
        const newStyle: DirectCellStyle = {
          ...style,
          areas: style.areas.map((area) =>
            area.workbookName === this.fromWorkbookName
              ? { ...area, workbookName: this.toWorkbookName }
              : area
          ),
        };
        styleManager.addCellStyle(newStyle);
      }
    }

    // Update formulas in cloned workbook that reference the source workbook
    workbookManager.updateFormulasForWorkbook(this.toWorkbookName, (formula) =>
      renameWorkbookInFormula({
        formula,
        oldWorkbookName: this.fromWorkbookName,
        newWorkbookName: this.toWorkbookName,
      })
    );
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [getWorkbookResourceKey(this.toWorkbookName)],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: [getWorkbookResourceKey(this.toWorkbookName)],
      removedScopes: [{ type: "workbook", workbookName: this.toWorkbookName }],
    };
  }

  undo(): void {
    // Undo clone = remove the cloned workbook
    removeWorkbook(this.deps, this.toWorkbookName);
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.CLONE_WORKBOOK,
      payload: {
        fromWorkbookName: this.fromWorkbookName,
        toWorkbookName: this.toWorkbookName,
      },
    };
  }
}

// ============================================================================
// Sheet Commands
// ============================================================================

/**
 * Command to add a sheet.
 */
export class AddSheetCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private opts: { workbookName: string; sheetName: string }
  ) {}

  execute(): void {
    this.deps.workbookManager.addSheet(this.opts);
    this.deps.namedExpressionManager.addSheet(this.opts);
    const resourceKey = getSheetResourceKey(this.opts);
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey, getWorkbookResourceKey(this.opts.workbookName)],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys: [resourceKey, getWorkbookResourceKey(this.opts.workbookName)],
      removedScopes: [{ type: "sheet", ...this.opts }],
    };
  }

  undo(): void {
    this.deps.workbookManager.removeSheet(this.opts);
    this.deps.namedExpressionManager.removeSheet(this.opts);
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.ADD_SHEET,
      payload: this.opts,
    };
  }
}

/**
 * Captured state of a sheet for undo purposes.
 */
interface SheetSnapshot {
  sheet: Sheet;
  namedExpressions: Map<string, NamedExpression>;
  tables: Map<string, TableDefinition>;
  conditionalStyles: ConditionalStyle[];
  cellStyles: DirectCellStyle[];
}

/**
 * Command to remove a sheet.
 */
export class RemoveSheetCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private snapshot: SheetSnapshot | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private opts: { workbookName: string; sheetName: string }
  ) {}

  execute(): void {
    // Capture sheet state before removal
    const sheet = this.deps.workbookManager.getSheet(this.opts);
    if (sheet) {
      const namedExpressions =
        this.deps.namedExpressionManager.getNamedExpressions();
      const sheetExpressions =
        namedExpressions.sheetExpressions
          .get(this.opts.workbookName)
          ?.get(this.opts.sheetName) || new Map();

      // Get tables on this sheet
      const allTables = this.deps.tableManager.getTables(
        this.opts.workbookName
      );
      const sheetTables = new Map<string, TableDefinition>();
      for (const [name, table] of allTables) {
        if (table.sheetName === this.opts.sheetName) {
          sheetTables.set(name, table);
        }
      }

      this.snapshot = {
        sheet: {
          name: sheet.name,
          index: sheet.index,
          content: new Map(sheet.content),
          metadata: new Map(sheet.metadata),
          sheetMetadata: sheet.sheetMetadata,
        },
        namedExpressions: new Map(sheetExpressions),
        tables: sheetTables,
        conditionalStyles: this.deps.styleManager
          .getAllConditionalStyles()
          .filter((s) =>
            s.areas.some(
              (a) =>
                a.workbookName === this.opts.workbookName &&
                a.sheetName === this.opts.sheetName
            )
          ),
        cellStyles: this.deps.styleManager
          .getAllCellStyles()
          .filter((s) =>
            s.areas.some(
              (a) =>
                a.workbookName === this.opts.workbookName &&
                a.sheetName === this.opts.sheetName
            )
          ),
      };
    }

    // Execute removal
    this.deps.workbookManager.removeSheet(this.opts);
    this.deps.namedExpressionManager.removeSheet(this.opts);
    this.deps.tableManager.removeSheet(this.opts);
    this.deps.styleManager.removeSheetStyles(
      this.opts.workbookName,
      this.opts.sheetName
    );
    this.deps.referenceManager.invalidateSheet(
      this.opts.workbookName,
      this.opts.sheetName
    );
    const resourceKeys = getSheetSnapshotResourceKeys(
      this.opts.workbookName,
      this.opts.sheetName,
      this.snapshot
    );
    this.executeFootprint = {
      touchedCells: [],
      resourceKeys,
      removedScopes: [{ type: "sheet", ...this.opts }],
    };
    this.undoFootprint = {
      touchedCells: [],
      resourceKeys,
    };
  }

  undo(): void {
    if (!this.snapshot) return;

    // Restore sheet structure
    this.deps.workbookManager.addSheet(this.opts);
    this.deps.namedExpressionManager.addSheet(this.opts);
    this.deps.workbookManager.setSheetContent(this.opts, this.snapshot.sheet.content);

    // Restore named expressions
    for (const [name, expr] of this.snapshot.namedExpressions) {
      this.deps.namedExpressionManager.addNamedExpression({
        expressionName: name,
        expression: expr.expression,
        workbookName: this.opts.workbookName,
        sheetName: this.opts.sheetName,
      });
    }

    // Restore tables and styles
    restoreTables(this.deps.tableManager, this.snapshot.tables, this.opts.workbookName);
    restoreStyles(this.deps.styleManager, this.snapshot.conditionalStyles, this.snapshot.cellStyles);
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.REMOVE_SHEET,
      payload: this.opts,
    };
  }
}

/**
 * Command to rename a sheet.
 */
export class RenameSheetCommand implements EngineCommand {
  readonly requiresReevaluation = true;
  private snapshot: SheetSnapshot | undefined;
  private executeFootprint = emptyMutationInvalidation();
  private undoFootprint = emptyMutationInvalidation();

  constructor(
    private deps: StructureCommandDeps,
    private opts: {
      workbookName: string;
      sheetName: string;
      newSheetName: string;
    }
  ) {}

  execute(): void {
    const namedExpressions =
      this.deps.namedExpressionManager.getNamedExpressions().sheetExpressions
        .get(this.opts.workbookName)
        ?.get(this.opts.sheetName) || new Map();
    const tables = new Map<string, TableDefinition>();
    for (const [name, table] of this.deps.tableManager.getTables(
      this.opts.workbookName
    )) {
      if (table.sheetName === this.opts.sheetName) {
        tables.set(name, table);
      }
    }
    const sheet = this.deps.workbookManager.getSheet(this.opts);
    this.snapshot = {
      sheet: sheet!,
      namedExpressions: new Map(namedExpressions),
      tables,
      conditionalStyles: [],
      cellStyles: [],
    };
    const changedCells = renameSheetAcrossManagers(this.deps, this.opts);
    this.executeFootprint = {
      touchedCells: changedCells.map((address) => ({
        address,
        beforeKind: "formula" as const,
        afterKind: "formula" as const,
      })),
      resourceKeys: Array.from(
        new Set([
          ...getSheetSnapshotResourceKeys(
            this.opts.workbookName,
            this.opts.sheetName,
            this.snapshot
          ),
          ...getSheetSnapshotResourceKeys(
            this.opts.workbookName,
            this.opts.newSheetName,
            this.snapshot
          ),
        ])
      ),
    };
  }

  undo(): void {
    const changedCells = renameSheetAcrossManagers(this.deps, {
      workbookName: this.opts.workbookName,
      sheetName: this.opts.newSheetName,
      newSheetName: this.opts.sheetName,
    });
    this.undoFootprint = {
      touchedCells: changedCells.map((address) => ({
        address,
        beforeKind: "formula" as const,
        afterKind: "formula" as const,
      })),
      resourceKeys: Array.from(
        new Set([
          ...getSheetSnapshotResourceKeys(
            this.opts.workbookName,
            this.opts.sheetName,
            this.snapshot
          ),
          ...getSheetSnapshotResourceKeys(
            this.opts.workbookName,
            this.opts.newSheetName,
            this.snapshot
          ),
        ])
      ),
    };
  }

  getInvalidationFootprint(phase: "execute" | "undo"): MutationInvalidation {
    return phase === "execute" ? this.executeFootprint : this.undoFootprint;
  }

  toAction(): EngineAction {
    return {
      type: ActionTypes.RENAME_SHEET,
      payload: this.opts,
    };
  }
}
