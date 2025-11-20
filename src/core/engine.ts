/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  type CellAddress,
  type CellStyle,
  type ConditionalStyle,
  type CopyCellsOptions,
  type DirectCellStyle,
  type NamedExpression,
  type RangeAddress,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "./types";

import type { FillDirection } from "@ricsam/selection-manager";
import { FormulaEvaluator } from "../evaluator/formula-evaluator";
import { AutoFill } from "./autofill-utils";
import { WorkbookManager } from "./managers/workbook-manager";
import { deserialize, serialize } from "./map-serializer";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameSheetInFormula } from "./sheet-renamer";
import { renameTableInFormula } from "./table-renamer";
import { renameWorkbookInFormula } from "./workbook-renamer";
import { cellAddressToKey, keyToCellAddress } from "./utils";
import { CacheManager } from "./managers/cache-manager";
import { NamedExpressionManager } from "./managers/named-expression-manager";
import { TableManager } from "./managers/table-manager";
import { EventManager } from "./managers/event-manager";
import { EvaluationManager } from "./managers/evaluation-manager";
import { DependencyManager } from "./managers/dependency-manager";
import { StyleManager } from "./managers/style-manager";
import { CopyManager } from "./managers/copy-manager";

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
  private styleManager: StyleManager;
  private copyManager: CopyManager;

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
  public _styleManager: StyleManager;

  constructor() {
    this.eventManager = new EventManager();
    this.workbookManager = new WorkbookManager();
    this.namedExpressionManager = new NamedExpressionManager();
    this.tableManager = new TableManager(this.workbookManager);
    const cacheManager = new CacheManager();
    this.dependencyManager = new DependencyManager(
      cacheManager,
      this.workbookManager
    );

    const formulaEvaluator = new FormulaEvaluator(
      this.tableManager,
      this.dependencyManager,
      this.namedExpressionManager,
    );

    this.evaluationManager = new EvaluationManager(
      this.workbookManager,
      this.tableManager,
      formulaEvaluator,
      this.dependencyManager
    );

    this.styleManager = new StyleManager(
      this.workbookManager,
      this.evaluationManager
    );
    this.copyManager = new CopyManager(
      this.workbookManager,
      this.evaluationManager,
      this.styleManager
    );

    this.autoFillManager = new AutoFill(this.workbookManager, this);

    this._workbookManager = this.workbookManager;
    this._namedExpressionManager = this.namedExpressionManager;
    this._tableManager = this.tableManager;
    this._eventManager = this.eventManager;
    this._evaluationManager = this.evaluationManager;
    this._autoFillManager = this.autoFillManager;
    this._dependencyManager = this.dependencyManager;
    this._styleManager = this.styleManager;
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
      cellAddress,
      debug
    );
  }

  evaluateFormula(
    /**
     * formula without the leading = sign
     */
    formula: string, cellAddress: CellAddress): SerializedCellValue {
    return this.evaluationManager.evaluateFormula(formula, cellAddress);
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
    return this.tableManager.isCellInTable(cellAddress);
  }

  //#endregion

  //#region Conditional Styling
  /**
   * Add a conditional style rule
   */
  addConditionalStyle(style: ConditionalStyle): void {
    this.styleManager.addConditionalStyle(style);
    this.eventManager.emitUpdate();
  }

  /**
   * Remove a conditional style rule by index
   */
  removeConditionalStyle(workbookName: string, index: number): boolean {
    const removed = this.styleManager.removeConditionalStyle(workbookName, index);
    if (removed) {
      this.eventManager.emitUpdate();
    }
    return removed;
  }

  /**
   * Get all conditional styles for a workbook
   */
  getConditionalStyles(workbookName: string): ConditionalStyle[] {
    return this.styleManager.getConditionalStyles(workbookName);
  }

  /**
   * Get the computed style for a specific cell
   */
  getCellStyle(cellAddress: CellAddress): CellStyle | undefined {
    return this.styleManager.getCellStyle(cellAddress);
  }

  /**
   * Add a direct cell style rule
   */
  addCellStyle(style: DirectCellStyle): void {
    this.styleManager.addCellStyle(style);
    this.eventManager.emitUpdate();
  }

  /**
   * Remove a direct cell style rule by index
   */
  removeCellStyle(workbookName: string, index: number): boolean {
    const removed = this.styleManager.removeCellStyle(workbookName, index);
    if (removed) {
      this.eventManager.emitUpdate();
    }
    return removed;
  }

  /**
   * Get all direct cell styles for a workbook
   */
  getCellStyles(workbookName: string): DirectCellStyle[] {
    return this.styleManager.getCellStyles(workbookName);
  }

  /**
   * Clear all cell styles and conditional styles for a given range
   * Adjusts existing style ranges rather than deleting them entirely
   */
  clearCellStyles(range: RangeAddress): void {
    this.styleManager.clearCellStyles(range);
    this.eventManager.emitUpdate();
  }

  //#endregion

  //#region Copy/Paste
  /**
   * Copy cells from source to target
   */
  copyCells(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): void {
    this.copyManager.copyCells(source, target, options);
    this.reevaluate();
    this.eventManager.emitUpdate();
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
    this.styleManager.removeSheetStyles(opts.workbookName, opts.sheetName);

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

    // Update conditional styles that reference this sheet
    this.styleManager.updateSheetName(
      opts.workbookName,
      opts.sheetName,
      opts.newSheetName
    );

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
    this.styleManager.removeWorkbookStyles(workbookName);

    this.reevaluate();
    this.eventManager.emitUpdate();
  }

  cloneWorkbook(fromWorkbookName: string, toWorkbookName: string) {
    // Check if source workbook exists
    const sourceWorkbook = this.workbookManager
      .getWorkbooks()
      .get(fromWorkbookName);
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
      this.setSheetContent(
        {
          workbookName: toWorkbookName,
          sheetName: sheetName,
        },
        new Map(sheet.content)
      );
    }

    // Clone workbook-scoped named expressions
    const sourceWorkbookExpressions = this.namedExpressionManager
      .getNamedExpressions()
      .workbookExpressions.get(fromWorkbookName);
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
    const sourceSheetExpressions = this.namedExpressionManager
      .getNamedExpressions()
      .sheetExpressions.get(fromWorkbookName);
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
        this.tableManager.copyTable(
          {
            workbookName: fromWorkbookName,
            tableName: tableName,
          },
          {
            workbookName: toWorkbookName,
            tableName: tableName,
          }
        );
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

    // Update conditional styles that reference this workbook
    this.styleManager.updateWorkbookName(
      opts.workbookName,
      opts.newWorkbookName
    );

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
  clearSpreadsheetRange(address: RangeAddress) {
    this.workbookManager.clearSpreadsheetRange(address);

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
      conditionalStyles: this.styleManager.getAllConditionalStyles(),
      cellStyles: this.styleManager.getAllCellStyles(),
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
    
    // Reset styles if present
    // Handle backward compatibility: if conditionalStyles is a Map, convert it
    let conditionalStylesArray: ConditionalStyle[] | undefined;
    let cellStylesArray: DirectCellStyle[] | undefined;
    
    if (deserialized.conditionalStyles) {
      if (deserialized.conditionalStyles instanceof Map) {
        // Old format: Map<string, ConditionalStyle[]>
        conditionalStylesArray = Array.from(deserialized.conditionalStyles.values()).flat();
      } else if (Array.isArray(deserialized.conditionalStyles)) {
        // New format: ConditionalStyle[]
        conditionalStylesArray = deserialized.conditionalStyles;
      }
    }
    
    if (deserialized.cellStyles) {
      if (Array.isArray(deserialized.cellStyles)) {
        cellStylesArray = deserialized.cellStyles;
      }
    }
    
    this.styleManager.resetStyles(conditionalStylesArray, cellStylesArray);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.eventManager.emitUpdate();
  }
  //#endregion
}
