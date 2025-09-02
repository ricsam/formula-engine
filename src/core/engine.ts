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
  SheetManager,
  TableManager,
} from "./managers";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameTableInFormula } from "./table-renamer";
import { type FunctionEvaluationResult } from "./types";

/**
 * Main FormulaEngine class
 */
export class FormulaEngine {
  public sheetManager: SheetManager;
  public namedExpressionManager: NamedExpressionManager;
  public tableManager: TableManager;
  public eventManager: EventManager;
  public evaluationManager: EvaluationManager;

  constructor() {
    this.eventManager = new EventManager();
    this.sheetManager = new SheetManager(this.eventManager);
    this.namedExpressionManager = new NamedExpressionManager(this.eventManager);
    this.tableManager = new TableManager(this.eventManager);
    this.evaluationManager = new EvaluationManager(
      this.sheetManager.getSheets(),
      this.namedExpressionManager.getScopedNamedExpressions(),
      this.namedExpressionManager.getGlobalNamedExpressions(),
      this.tableManager.getTables()
    );
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
  }: {
    expression: string;
    expressionName: string;
    sheetName?: string;
  }) {
    this.namedExpressionManager.updateNamedExpression({
      expression,
      expressionName,
      sheetName,
    });

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  renameNamedExpression({
    expressionName,
    sheetName,
    newName,
  }: {
    expressionName: string;
    sheetName?: string;
    newName: string;
  }) {
    const result = this.namedExpressionManager.renameNamedExpression({
      expressionName,
      sheetName,
      newName,
    });

    // Update all formulas that reference this named expression in sheet cells
    this.sheetManager.updateFormulasForSheetRename(
      expressionName, // This is a bit of a hack - we're reusing the sheet rename method
      newName,
      (formula) =>
        renameNamedExpressionInFormula(formula, expressionName, newName)
    );

    // Update named expressions that reference this named expression
    this.namedExpressionManager.updateFormulasForNamedExpressionRename(
      expressionName,
      newName
    );

    // Re-evaluate all sheets since named expressions can be referenced from anywhere
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return result;
  }

  makeTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
  }: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    return this.tableManager.makeTable({
      tableName,
      sheetName,
      start,
      numRows,
      numCols,
      getSheetContent: (sheetName) =>
        this.sheetManager.getSheet(sheetName)?.content,
    });
  }

  addTable(props: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
  }) {
    const table = this.tableManager.addTable({
      ...props,
      getSheetContent: (sheetName) =>
        this.sheetManager.getSheet(sheetName)?.content,
    });

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return table;
  }

  renameTable(names: { oldName: string; newName: string }) {
    this.tableManager.renameTable(names);

    // Update all formulas that reference this table in sheet cells
    this.sheetManager.updateFormulasForSheetRename(
      names.oldName, // Reusing sheet rename method
      names.newName,
      (formula) => renameTableInFormula(formula, names.oldName, names.newName)
    );

    // Update named expressions that reference this table
    this.namedExpressionManager.updateFormulasForNamedExpressionRename(
      names.oldName, // Reusing named expression rename method
      names.newName,
      (formula) => renameTableInFormula(formula, names.oldName, names.newName)
    );

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  updateTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
  }: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
  }) {
    this.tableManager.updateTable({
      tableName,
      sheetName,
      start,
      numRows,
      numCols,
      getSheetContent: (sheetName) =>
        this.sheetManager.getSheet(sheetName)?.content,
    });

    // Re-evaluate all sheets since structured references might depend on this table
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  removeTable({ tableName }: { tableName: string }) {
    const found = this.tableManager.removeTable({ tableName });

    if (found) {
      // Re-evaluate all sheets since structured references might depend on this table
      this.reevaluate();
      this.triggerCellsUpdateEvent();
    }

    return found;
  }

  addSheet(name: string) {
    return this.sheetManager.addSheet(name);
  }

  removeSheet(sheetName: string) {
    const sheet = this.sheetManager.removeSheet(sheetName);

    // Clean up related data
    this.namedExpressionManager.removeSheetNamedExpressions(sheetName);
    this.tableManager.removeTablesForSheet(sheetName);
    this.eventManager.removeCellsUpdateListenersForSheet(sheetName);

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return sheet;
  }

  renameSheet(sheetName: string, newName: string) {
    const sheet = this.sheetManager.renameSheet(sheetName, newName);

    // Update scoped named expressions
    this.namedExpressionManager.renameSheetNamedExpressions(sheetName, newName);

    // Update tables that belong to the renamed sheet
    this.tableManager.updateTablesForSheetRename(sheetName, newName);

    // Update cell update listeners
    this.eventManager.renameCellsUpdateListenersForSheet(sheetName, newName);

    // Update all formulas that reference this sheet
    this.sheetManager.updateFormulasForSheetRename(sheetName, newName);

    // Add engine-specific logic: re-evaluate since references might be affected
    this.reevaluate();
    this.triggerCellsUpdateEvent();

    return sheet;
  }

  getTablesSerialized(): Map<string, TableDefinition> {
    return this.tableManager.getTablesSerialized();
  }

  getNamedExpressionsSerialized(
    sheetName: string
  ): Map<string, NamedExpression> {
    return this.namedExpressionManager.getNamedExpressionsSerialized(sheetName);
  }

  getGlobalNamedExpressionsSerialized(): Map<string, NamedExpression> {
    return this.namedExpressionManager.getGlobalNamedExpressionsSerialized();
  }

  setNamedExpressions(
    sheetName: string,
    namedExpressions: Map<string, NamedExpression>
  ) {
    this.namedExpressionManager.setNamedExpressions(
      sheetName,
      namedExpressions
    );
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
    sheetName: string,
    content: Map<string, SerializedCellValue>
  ) {
    this.sheetManager.setSheetContent(sheetName, content);

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  setCellContent(address: CellAddress, content: SerializedCellValue) {
    this.sheetManager.setCellContent(address, content);
    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.triggerCellsUpdateEvent();
  }

  triggerCellsUpdateEvent() {
    this.eventManager.triggerCellsUpdateEvent();
  }

  reevaluateSheet(sheetName: string) {
    this.sheetManager.reevaluateSheet(sheetName, (address) => {
      this.evaluationManager.evaluateCell(address);
    });
  }

  /**
   * Re-evaluates all sheets to ensure all dependencies are resolved correctly
   */
  reevaluate() {
    this.evaluationManager.clearEvaluationCache();
    for (const sheet of this.sheetManager.getSheets().values()) {
      this.reevaluateSheet(sheet.name);
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
  onCellsUpdate(sheetName: string, listener: () => void): () => void {
    return this.eventManager.onCellsUpdate(sheetName, listener);
  }

  getSheetSerialized(sheetName: string): Map<string, SerializedCellValue> {
    return this.sheetManager.getSheetSerialized(sheetName);
  }

  // ===== Evaluation System Delegation =====

  /**
   * Access to evaluation manager for functions that need it
   */
  get evaluator() {
    return this.evaluationManager;
  }

  isCellInRange(cellAddress: CellAddress, range: SpreadsheetRange): boolean {
    return this.evaluationManager.isCellInRange(cellAddress, range);
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    return this.evaluationManager.isCellInTable(cellAddress);
  }

  // Additional methods that might be needed
  get sheets() {
    return this.sheetManager.getSheets();
  }

  get tables() {
    return this.tableManager.getTables();
  }

  get globalNamedExpressions() {
    return this.namedExpressionManager.getGlobalNamedExpressions();
  }

  get scopedNamedExpressions() {
    return this.namedExpressionManager.getScopedNamedExpressions();
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
}
