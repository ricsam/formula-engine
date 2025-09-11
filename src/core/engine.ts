/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  type CellAddress,
  type EvaluationContext,
  type NamedExpression,
  type SerializedCellValue,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
} from "./types";

import {
  EvaluationManager,
  EventManager,
  NamedExpressionManager,
  TableManager,
} from "./managers";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameTableInFormula } from "./table-renamer";
import { type FunctionEvaluationResult } from "./types";
import type { FillDirection } from "@ricsam/selection-manager";
import { getCellReference } from "./utils";
import { AutoFill } from "./autofill-utils";
import { WorkbookManager } from "./managers/workbook-manager";
import { renameSheetInFormula } from "./sheet-renamer";
import { FormulaEvaluator } from "src/evaluator/formula-evaluator";

/**
 * Main FormulaEngine class
 */
export class FormulaEngine {
  public workbookManager: WorkbookManager;
  public namedExpressionManager: NamedExpressionManager;
  public tableManager: TableManager;
  public eventManager: EventManager;
  public evaluationManager: EvaluationManager;
  public autoFillManager: AutoFill;

  constructor() {
    this.eventManager = new EventManager();
    this.workbookManager = new WorkbookManager(this.eventManager);
    this.namedExpressionManager = new NamedExpressionManager(this.eventManager);
    this.tableManager = new TableManager(
      this.eventManager,
      this.workbookManager
    );
    const formulaEvaluator = new FormulaEvaluator(
      this.tableManager,
      (...args) => this.evaluationManager.evalTimeSafeEvaluateCell(...args),
      (...args) => this.evaluationManager.evalTimeSafeEvaluateNamedExpression(...args)
    );
    this.evaluationManager = new EvaluationManager(
      this.workbookManager,
      this.namedExpressionManager,
      formulaEvaluator
    );
    this.autoFillManager = new AutoFill(this.workbookManager, this);
  }

  /**
   * Static factory method to build an empty engine
   */
  static buildEmpty(): FormulaEngine {
    return new FormulaEngine();
  }

  getCellEvaluationResult(
    cellAddress: CellAddress
  ): SingleEvaluationResult | undefined {
    return this.evaluationManager.getCellEvaluationResult(cellAddress);
  }

  getCellValue(cellAddress: CellAddress, debug?: boolean): SerializedCellValue {
    const result = this.getCellEvaluationResult(cellAddress);
    if (!result) {
      return undefined;
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

  addNamedExpression({
    expression,
    expressionName,
    sheetName,
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
  }) {
    this.namedExpressionManager.addNamedExpression({
      expression,
      expressionName,
      sheetName,
    });

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  removeNamedExpression({
    expressionName,
    sheetName,
  }: {
    expressionName: string;
    sheetName?: string;
  }) {
    const found = this.namedExpressionManager.removeNamedExpression({
      expressionName,
      sheetName,
    });

    if (found) {
      // Re-evaluate all sheets since named expressions can be referenced from anywhere
      this.reevaluate();
      this.triggerCellsUpdateEvent();
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
    this.triggerCellsUpdateEvent();
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
    this.triggerCellsUpdateEvent();

    return result;
  }

  makeTable(props: {
    tableName: string;
    sheetName: string;
    workbookName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    return this.tableManager.makeTable(props);
  }

  addTable(props: {
    tableName: string;
    sheetName: string;
    workbookName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    const table = this.tableManager.addTable(props);

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return table;
  }

  renameTable(names: { oldName: string; newName: string }) {
    this.tableManager.renameTable(names);

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
    this.triggerCellsUpdateEvent();
  }

  updateTable(opts: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
    workbookName?: string;
  }) {
    this.tableManager.updateTable(opts);

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  removeTable(opts: { tableName: string }) {
    const found = this.tableManager.removeTable(opts);

    if (found) {
      // Re-evaluate all sheets since structured references might depend on this table
      this.reevaluate();
      this.triggerCellsUpdateEvent();
    }

    return found;
  }

  addSheet(opts: { workbookName: string; sheetName: string }) {
    return this.workbookManager.addSheet(opts);
  }

  removeSheet(opts: { workbookName: string; sheetName: string }) {
    const sheet = this.workbookManager.removeSheet(opts);

    // Clean up related data
    this.namedExpressionManager.removeSheetExpressions(opts);
    this.tableManager.removeTablesForSheet(opts);
    this.eventManager.removeCellsUpdateListenersForSheet(opts);

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return sheet;
  }

  renameSheet(opts: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }) {
    const sheet = this.workbookManager.renameSheet(opts);

    // Update scoped named expressions
    this.namedExpressionManager.renameSheetExpressions(opts);

    // Update tables that belong to the renamed sheet
    this.tableManager.updateTablesForSheetRename(opts);

    // Update cell update listeners
    this.eventManager.renameCellsUpdateListenersForSheet(opts);

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
    this.triggerCellsUpdateEvent();

    return sheet;
  }

  getTablesSerialized(): Map<string, TableDefinition> {
    return this.tableManager.getTablesSerialized();
  }

  getSheetExpressionsSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, NamedExpression> {
    return this.namedExpressionManager.getSheetExpressionsSerialized(opts);
  }

  getGlobalNamedExpressionsSerialized(): Map<string, NamedExpression> {
    return this.namedExpressionManager.getGlobalNamedExpressionsSerialized();
  }

  setNamedExpressions(opts: {
    sheetName: string;
    workbookName: string;
    expressions: Map<string, NamedExpression>;
  }) {
    this.namedExpressionManager.setNamedExpressions({
      ...opts,
      newExpressions: opts.expressions,
    });

    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  setGlobalNamedExpressions(namedExpressions: Map<string, NamedExpression>) {
    this.namedExpressionManager.setGlobalNamedExpressions(namedExpressions);
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  setTables(tables: Map<string, TableDefinition>) {
    this.tableManager.setTables(tables);
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  /**
   * Overrides the content of a sheet.
   * @param sheetName - The name of the sheet to set the content of
   * @param content - A map of cell addresses to their serialized values
   * @remarks This method is used to set the content of a sheet. It will re-evaluate all sheets to ensure all dependencies are resolved correctly.
   */
  public setSheetContent(
    opts: { sheetName: string; workbookName: string },
    content: Map<string, SerializedCellValue>
  ) {
    this.workbookManager.setSheetContent(opts, content);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  setCellContent(address: CellAddress, content: SerializedCellValue) {
    this.workbookManager.setCellContent(address, content);
    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  triggerCellsUpdateEvent() {
    this.eventManager.triggerCellsUpdateEvent();
  }

  reevaluateSheet(opts: { sheetName: string; workbookName: string }) {
    this.workbookManager.reevaluateSheet(opts, (address) => {
      this.evaluationManager.evaluateCell(address);
    });
  }

  /**
   * Re-evaluates all sheets to ensure all dependencies are resolved correctly
   */
  reevaluate() {
    this.evaluationManager.clearEvaluationCache();
    for (const workbook of this.workbookManager.getWorkbooks().values()) {
      for (const sheet of workbook.sheets.values()) {
        this.reevaluateSheet({
          sheetName: sheet.name,
          workbookName: workbook.name,
        });
      }
    }
  }

  // ===== Event System Delegation =====

  /**
   * Subscribe to FormulaEngine events
   */
  on<K extends keyof import("./types").FormulaEngineEvents>(
    event: K,
    listener: (data: import("./types").FormulaEngineEvents[K]) => void
  ): () => void {
    return this.eventManager.on(event, listener);
  }

  /**
   * Subscribe to FormulaEngine events (alias for on)
   */
  subscribe<K extends keyof import("./types").FormulaEngineEvents>(
    event: K,
    listener: (data: import("./types").FormulaEngineEvents[K]) => void
  ): () => void {
    return this.eventManager.subscribe(event, listener);
  }

  /**
   * Remove all event listeners
   */
  removeAllListeners(): void {
    this.eventManager.removeAllListeners();
  }

  /**
   * Register listener for batched sheet updates. Returns an unsubscribe function.
   */
  onCellsUpdate(
    opts: { sheetName: string; workbookName: string },
    listener: () => void
  ): () => void {
    return this.eventManager.onCellsUpdate(opts, listener);
  }

  getSheetSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, SerializedCellValue> {
    return this.workbookManager.getSheetSerialized(opts);
  }

  // ===== Evaluation System Delegation =====

  /**
   * Access to evaluation manager for functions that need it
   */
  get evaluator() {
    return this.evaluationManager;
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.evaluationManager.isCellInTable(cellAddress);
  }

  // Additional methods that might be needed
  getSheets(workbookName: string) {
    return this.workbookManager.getSheets(workbookName);
  }

  get tables() {
    return this.tableManager.getTables();
  }

  get globalNamedExpressions() {
    return this.namedExpressionManager.getGlobalNamedExpressions();
  }

  get evaluatedNodes() {
    return this.evaluationManager.getEvaluatedNodes();
  }

  get spilledValues() {
    return this.evaluationManager.getSpilledValues();
  }

  getTransitiveDeps(nodeKey: string): Set<string> {
    return this.evaluationManager.getTransitiveDeps(nodeKey);
  }

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
    this.workbookManager.clearSpreadsheetRange(opts, range, (content) =>
      this.setSheetContent(opts, content)
    );
  }

  addWorkbook(workbookName: string) {}
  getWorkbooks(): unknown {
    throw new Error("Not implemented");
  }
}
