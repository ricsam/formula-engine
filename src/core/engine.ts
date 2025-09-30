/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  type CellAddress,
  type NamedExpression,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "./types";

import type { FillDirection } from "@ricsam/selection-manager";
import { FormulaEvaluator } from "src/evaluator/formula-evaluator";
import { AutoFill } from "./autofill-utils";
import {
  DependencyManager,
  EvaluationManager,
  EventManager,
  NamedExpressionManager,
  TableManager,
} from "./managers";
import { WorkbookManager } from "./managers/workbook-manager";
import { deserialize, serialize } from "./map-serializer";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameSheetInFormula } from "./sheet-renamer";
import { renameTableInFormula } from "./table-renamer";
import { renameWorkbookInFormula } from "./workbook-renamer";
import { cellAddressToKey, keyToCellAddress } from "./utils";

/**
 * Main FormulaEngine class
 */
export class FormulaEngine {
  private workbookManager: WorkbookManager;
  private namedExpressionManager: NamedExpressionManager;
  private tableManager: TableManager;
  private eventManager: EventManager;
  private evaluationManager: EvaluationManager;
  private autoFillManager: AutoFill;
  private dependencyManager: DependencyManager;

  /**
   * Public access to the store manager for testing
   */
  public _workbookManager: WorkbookManager;
  public _namedExpressionManager: NamedExpressionManager;
  public _tableManager: TableManager;
  public _eventManager: EventManager;
  public _evaluationManager: EvaluationManager;
  public _autoFillManager: AutoFill;
  public _dependencyManager: DependencyManager;

  constructor() {
    this.eventManager = new EventManager();
    this.workbookManager = new WorkbookManager();
    this.namedExpressionManager = new NamedExpressionManager();
    this.tableManager = new TableManager(this.workbookManager);
    this.dependencyManager = new DependencyManager();

    const formulaEvaluator = new FormulaEvaluator(
      this.tableManager,
      this.dependencyManager,
      this.namedExpressionManager,
      this.workbookManager
    );

    this.evaluationManager = new EvaluationManager(
      this.workbookManager,
      formulaEvaluator,
      this.dependencyManager
    );

    this.autoFillManager = new AutoFill(this.workbookManager, this);

    this._workbookManager = this.workbookManager;
    this._namedExpressionManager = this.namedExpressionManager;
    this._tableManager = this.tableManager;
    this._eventManager = this.eventManager;
    this._evaluationManager = this.evaluationManager;
    this._autoFillManager = this.autoFillManager;
    this._dependencyManager = this.dependencyManager;
  }

  /**
   * Static factory method to build an empty engine
   */
  static buildEmpty(): FormulaEngine {
    return new FormulaEngine();
  }

  //#region Cell
  getCellEvaluationResult(
    cellAddress: CellAddress
  ): SingleEvaluationResult | undefined {
    return this.evaluationManager.getCellEvaluationResult(cellAddress);
  }

  getCellValue(cellAddress: CellAddress, debug?: boolean): SerializedCellValue {
    const result = this.getCellEvaluationResult(cellAddress);
    if (!result) {
      return "";
    }

    return this.evaluationManager.evaluationResultToSerializedValue(
      result,
      debug
    );
  }

  getCellDependents(
    address: CellAddress | SpreadsheetRange
  ): (SpreadsheetRange | CellAddress)[] {
    throw new Error("Not implemented");
  }

  getCellPrecedents(
    address: CellAddress | SpreadsheetRange
  ): (SpreadsheetRange | CellAddress)[] {
    throw new Error("Not implemented");
  }

  //#endregion

  //#region Named Expressions
  addNamedExpression({
    expression,
    expressionName,
    sheetName,
    workbookName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }) {
    this.namedExpressionManager.addNamedExpression({
      expression,
      expressionName,
      sheetName,
      workbookName,
    });

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  removeNamedExpression({
    expressionName,
    sheetName,
    workbookName,
  }: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }) {
    const found = this.namedExpressionManager.removeNamedExpression({
      expressionName,
      sheetName,
      workbookName,
    });

    if (found) {
      // Re-evaluate all sheets since named expressions can be referenced from anywhere
      this.reevaluate();
      this.eventManager.emitUpdate();
    }

    return found;
  }

  updateNamedExpression({
    expression,
    expressionName,
    sheetName,
    workbookName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }) {
    this.namedExpressionManager.updateNamedExpression({
      expression,
      expressionName,
      sheetName,
      workbookName,
    });

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  renameNamedExpression({
    expressionName,
    sheetName,
    workbookName,
    newName,
  }: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
    newName: string;
  }) {
    const result = this.namedExpressionManager.renameNamedExpression({
      expressionName,
      sheetName,
      workbookName,
      newName,
    });

    // Update all formulas that reference this named expression in sheet cells
    this.workbookManager.updateAllFormulas((formula) =>
      renameNamedExpressionInFormula(formula, expressionName, newName)
    );

    // Update named expressions that reference this named expression
    this.namedExpressionManager.updateAllNamedExpressions((formula) =>
      renameNamedExpressionInFormula(formula, expressionName, newName)
    );

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.eventManager.emitUpdate();

    return result;
  }

  setNamedExpressions(
    opts: (
      | {
          type: "global";
        }
      | {
          type: "sheet";
          sheetName: string;
          workbookName: string;
        }
      | {
          type: "workbook";
          workbookName: string;
        }
    ) & {
      expressions: Map<string, NamedExpression>;
    }
  ) {
    this.namedExpressionManager.setNamedExpressions(opts);

    this.reevaluate();
    this.eventManager.emitUpdate();
  }
  //#endregion

  //#region Tables
  addTable(props: {
    tableName: string;
    sheetName: string;
    workbookName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    const table = this.tableManager.addTable({
      ...props,
      getCellValue: (cellAddress: CellAddress) =>
        this.getCellValue(cellAddress),
    });

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.eventManager.emitUpdate();

    return table;
  }

  renameTable(
    workbookName: string,
    names: { oldName: string; newName: string }
  ) {
    this.tableManager.renameTable(workbookName, names);

    // Update all formulas that reference this table in sheet cells
    this.workbookManager.updateAllFormulas((formula) =>
      renameTableInFormula(formula, names.oldName, names.newName)
    );

    // Update named expressions that reference this table
    this.namedExpressionManager.updateAllNamedExpressions((formula) =>
      renameTableInFormula(formula, names.oldName, names.newName)
    );

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  updateTable(opts: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
    workbookName: string;
  }) {
    this.tableManager.updateTable({
      ...opts,
      getCellValue: (cellAddress: CellAddress) =>
        this.getCellValue(cellAddress),
    });

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  removeTable(opts: { tableName: string; workbookName: string }) {
    const found = this.tableManager.removeTable(opts);

    if (found) {
      // Re-evaluate all sheets since structured references might depend on this table
      this.reevaluate();
      this.eventManager.emitUpdate();
    }

    return found;
  }

  resetTables(tables: Map<string, Map<string, TableDefinition>>) {
    this.tableManager.resetTables(tables);
    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  getTables(workbookName: string) {
    return this.tableManager.getTables(workbookName);
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.evaluationManager.isCellInTable(cellAddress);
  }

  //#endregion

  //#region Sheets
  addSheet(opts: { workbookName: string; sheetName: string }) {
    const newSheet = this.workbookManager.addSheet(opts);
    const wbLevel = this.namedExpressionManager.addSheet(opts);
    this.reevaluate();
    this.eventManager.emitUpdate();
    return newSheet;
  }

  removeSheet(opts: { workbookName: string; sheetName: string }) {
    const sheet = this.workbookManager.removeSheet(opts);

    // Clean up related data
    this.namedExpressionManager.removeSheet(opts);
    this.tableManager.removeSheet(opts);

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.eventManager.emitUpdate();

    return sheet;
  }

  renameSheet(opts: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }) {
    const sheet = this.workbookManager.renameSheet(opts);

    // Update scoped named expressions
    this.namedExpressionManager.renameSheet(opts);

    // Update tables that belong to the renamed sheet
    this.tableManager.updateTablesForSheetRename(opts);

    // Update all formulas that reference this sheet
    this.workbookManager.updateAllFormulas((formula) =>
      renameSheetInFormula({
        formula,
        oldSheetName: opts.sheetName,
        newSheetName: opts.newSheetName,
      })
    );

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.eventManager.emitUpdate();

    return sheet;
  }

  getSheets(workbookName: string) {
    return this.workbookManager.getSheets(workbookName);
  }

  getSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }) {
    return this.workbookManager.getSheet({ workbookName, sheetName });
  }

  getSheetSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, SerializedCellValue> {
    return this.workbookManager.getSheetSerialized(opts);
  }

  //#endregion

  //#region Workbook
  addWorkbook(workbookName: string) {
    this.workbookManager.addWorkbook(workbookName);
    this.namedExpressionManager.addWorkbook(workbookName);
    this.tableManager.addWorkbook(workbookName);

    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  removeWorkbook(workbookName: string) {
    this.workbookManager.removeWorkbook(workbookName);
    this.namedExpressionManager.removeWorkbook(workbookName);
    this.tableManager.removeWorkbook(workbookName);

    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  cloneWorkbook(fromWorkbookName: string, toWorkbookName: string) {
    // Check if source workbook exists
    const sourceWorkbook = this.workbookManager.getWorkbooks().get(fromWorkbookName);
    if (!sourceWorkbook) {
      throw new Error(`Source workbook "${fromWorkbookName}" not found`);
    }

    // Check if target workbook name already exists
    if (this.workbookManager.getWorkbooks().has(toWorkbookName)) {
      throw new Error(`Target workbook "${toWorkbookName}" already exists`);
    }

    // Create new workbook
    this.addWorkbook(toWorkbookName);

    // Clone all sheets from source workbook
    for (const [sheetName, sheet] of sourceWorkbook.sheets) {
      // Add sheet to target workbook
      this.addSheet({
        workbookName: toWorkbookName,
        sheetName: sheetName,
      });

      // Copy all cell content using setSheetContent for efficiency
      this.setSheetContent({
        workbookName: toWorkbookName,
        sheetName: sheetName,
      }, new Map(sheet.content));
    }

    // Clone workbook-scoped named expressions
    const sourceWorkbookExpressions = this.namedExpressionManager.getNamedExpressions().workbookExpressions.get(fromWorkbookName);
    if (sourceWorkbookExpressions) {
      for (const [name, expression] of sourceWorkbookExpressions) {
        this.addNamedExpression({
          expressionName: name,
          expression: expression.expression,
          workbookName: toWorkbookName,
        });
      }
    }

    // Clone sheet-scoped named expressions
    const sourceSheetExpressions = this.namedExpressionManager.getNamedExpressions().sheetExpressions.get(fromWorkbookName);
    if (sourceSheetExpressions) {
      for (const [sheetName, sheetExpressions] of sourceSheetExpressions) {
        for (const [name, expression] of sheetExpressions) {
          this.addNamedExpression({
            expressionName: name,
            expression: expression.expression,
            workbookName: toWorkbookName,
            sheetName: sheetName,
          });
        }
      }
    }

    // Clone tables
    const sourceTables = this.tableManager.tables.get(fromWorkbookName);
    if (sourceTables) {
      for (const [tableName, table] of sourceTables) {
        // Convert start position to cell reference string
        const startCellRef = `${String.fromCharCode(65 + table.start.colIndex)}${table.start.rowIndex + 1}`;
        
        // Calculate numRows and numCols from the original table
        const numCols = table.headers.size;
        const numRows = table.endRow;
        
        this.addTable({
          workbookName: toWorkbookName,
          tableName: tableName,
          sheetName: table.sheetName,
          start: startCellRef,
          numRows: numRows,
          numCols: numCols,
        });
      }
    }

    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  renameWorkbook(opts: { workbookName: string; newWorkbookName: string }) {
    this.workbookManager.renameWorkbook(opts);

    // Update scoped named expressions
    this.namedExpressionManager.renameWorkbook(opts);

    // Update tables that belong to the renamed sheet
    this.tableManager.updateTablesForWorkbookRename(opts);

    // Update all formulas that reference this workbook
    this.workbookManager.updateAllFormulas((formula) =>
      renameWorkbookInFormula({
        formula,
        oldWorkbookName: opts.workbookName,
        newWorkbookName: opts.newWorkbookName,
      })
    );

    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  getWorkbooks() {
    return this.workbookManager.getWorkbooks();
  }
  //#endregion

  //#region CRUD Operations
  /**
   * Overrides the content of a sheet.
   * @param sheetName - The name of the sheet to set the content of
   * @param content - A map of cell addresses to their serialized values
   * @remarks This method is used to set the content of a sheet. It will re-evaluate all sheets to ensure all dependencies are resolved correctly.
   */
  setSheetContent(
    opts: { sheetName: string; workbookName: string },
    content: Map<string, SerializedCellValue>
  ) {
    this.workbookManager.setSheetContent(opts, content);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  setCellContent(address: CellAddress, content: SerializedCellValue) {
    this.workbookManager.setCellContent(address, content);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.eventManager.emitUpdate();
  }
  //#endregion

  //#region Evaluation

  /**
   * Re-evaluates all sheets to ensure all dependencies are resolved correctly
   *
   * but just clears the evaluation cache
   */
  reevaluate() {
    this.evaluationManager.clearEvaluationCache();
  }
  //#endregion

  //#region Auto-fill
  /**
   * Auto-fills the fillRange based on the seedRange and the direction.
   */
  autoFill(
    opts: { sheetName: string; workbookName: string },
    /**
     * The user's original selection that defines the pattern/series.
     */
    seedRange: SpreadsheetRange,
    /**
     * the new cells populated by the drag, excluding the seed
     */
    fillRange: SpreadsheetRange,
    /**
     * The direction of the fill.
     */
    direction: FillDirection
  ) {
    this.autoFillManager.fill(opts, seedRange, fillRange, direction);
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   */
  clearSpreadsheetRange(
    opts: { sheetName: string; workbookName: string },
    range: SpreadsheetRange
  ) {
    this.workbookManager.clearSpreadsheetRange(opts, range);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.eventManager.emitUpdate();
  }
  //#endregion

  //#region State - UI library integration
  getState() {
    return {
      workbooks: this.workbookManager.getWorkbooks(),
      namedExpressions: this.namedExpressionManager.getNamedExpressions(),
      tables: this.tableManager.tables,
    };
  }

  onUpdate(listener: () => void) {
    return this.eventManager.onUpdate(listener);
  }

  serializeEngine(): string {
    return serialize(this.getState());
  }

  resetToSerializedEngine(data: string) {
    const deserialized = deserialize(data) as ReturnType<typeof this.getState>;

    this.workbookManager.resetWorkbooks(deserialized.workbooks);

    deserialized.workbooks.forEach((workbook) => {
      this.namedExpressionManager.addWorkbook(workbook.name);
      this.tableManager.addWorkbook(workbook.name);
      workbook.sheets.forEach((sheet) => {
        this.namedExpressionManager.addSheet({
          workbookName: workbook.name,
          sheetName: sheet.name,
        });
      });
    });

    this.namedExpressionManager.resetNamedExpressions(
      deserialized.namedExpressions
    );
    this.tableManager.resetTables(deserialized.tables);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.eventManager.emitUpdate();
  }
  //#endregion
}
