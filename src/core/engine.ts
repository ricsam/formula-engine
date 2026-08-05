/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  type CellAddress,
  type CellDataType,
  type CellStyle,
  type ConditionalStyle,
  type CopyCellsOptions,
  type DirectCellDataType,
  type DirectCellStyle,
  type FormulaEngineOptions,
  type NamedExpression,
  type RangeAddress,
  type RangeMetadata,
  type RangeMetadataInput,
  type ReplaceChange,
  type ReplaceTarget,
  type SearchMatch,
  type SearchOptions,
  type SerializedCellValue,
  type Sheet,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
  type UndoRedoState,
} from "./types";

import type { FillDirection } from "@ricsam/selection-manager";
import { FormulaEvaluator } from "../evaluator/formula-evaluator";
import { AutoFill } from "./autofill-utils";
import { WorkbookManager } from "./managers/workbook-manager";
import { deserialize, serialize } from "./map-serializer";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameSheetInFormula } from "./sheet-renamer";
import {
  renameTableColumnsInFormula,
  renameTableInFormula,
} from "./table-renamer";
import { renameWorkbookInFormula } from "./workbook-renamer";
import { getCellReference, parseCellReference } from "./utils";
import { CacheManager } from "./managers/cache-manager";
import { NamedExpressionManager } from "./managers/named-expression-manager";
import {
  TableManager,
  type TableHeaderUpdate,
} from "./managers/table-manager";
import { EventManager } from "./managers/event-manager";
import { EvaluationManager } from "./managers/evaluation-manager";
import { DependencyManager } from "./managers/dependency-manager";
import { StyleManager } from "./managers/style-manager";
import { CopyManager } from "./managers/copy-manager";
import { ReferenceManager } from "./managers/reference-manager";
import { RangeMetadataManager } from "./managers/range-metadata-manager";
import { UndoRedoManager } from "./managers/undo-redo-manager";
import {
  ENGINE_SNAPSHOT_VERSION,
  LEGACY_ENGINE_SNAPSHOT_VERSION,
  type EngineHistorySnapshot,
  type EngineSnapshot,
  type StyleManagerSnapshot,
} from "./engine-snapshot";
import {
  buildFormulaTouchedCells,
  buildSheetContentTouchedCells,
  buildTableContextChangedCells,
  buildTableTouchedCells,
  buildTouchedCells,
  captureCellContents,
  getFiniteRangeAddresses,
  getMutationAddressKey,
  getNamedExpressionScopeResourceKeys,
  mergeTouchedCells,
  type MutationInvalidation,
} from "./mutation-invalidation";
import {
  getNamedExpressionResourceKey,
  getSheetResourceKey,
  getTableResourceKey,
  getWorkbookResourceKey,
} from "./resource-keys";

type Metadata = {
  cell?: unknown;
  range?: unknown;
  sheet?: unknown;
  workbook?: unknown;
};

type MetadataType<
  TMetadata extends Metadata,
  TKey extends keyof Metadata
> = TMetadata[TKey];

 /**
  * Main FormulaEngine class
 * @template TMetadata - Consumer-defined metadata shape with optional cell, sheet, and workbook entries.
 */
export class FormulaEngine<TMetadata extends Metadata = Metadata> {
  private workbookManager: WorkbookManager;
  private namedExpressionManager: NamedExpressionManager;
  private tableManager: TableManager;
  private eventManager: EventManager;
  private evaluationManager: EvaluationManager;
  private autoFillManager: AutoFill;
  private dependencyManager: DependencyManager;
  private styleManager: StyleManager;
  private rangeMetadataManager: RangeMetadataManager<MetadataType<TMetadata, "range">>;
  private copyManager: CopyManager;
  private referenceManager: ReferenceManager;
  private undoRedoManager: UndoRedoManager;
  private historyCheckpointDepth = 0;

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
  public _rangeMetadataManager: RangeMetadataManager<MetadataType<TMetadata, "range">>;

  constructor(options: FormulaEngineOptions = {}) {
    this.undoRedoManager = new UndoRedoManager(options.undoRedo);
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
      this.namedExpressionManager
    );

    this.evaluationManager = new EvaluationManager(
      this.workbookManager,
      this.tableManager,
      formulaEvaluator,
      this.dependencyManager
    );

    this.styleManager = new StyleManager(this.evaluationManager);
    this.rangeMetadataManager = new RangeMetadataManager<
      MetadataType<TMetadata, "range">
    >();
    this.copyManager = new CopyManager(
      this.workbookManager,
      this.evaluationManager,
      this.styleManager,
      this.rangeMetadataManager
    );

    this.autoFillManager = new AutoFill(
      this.workbookManager,
      this.styleManager,
      this.rangeMetadataManager
    );

    this.referenceManager = new ReferenceManager();

    this._workbookManager = this.workbookManager;
    this._namedExpressionManager = this.namedExpressionManager;
    this._tableManager = this.tableManager;
    this._eventManager = this.eventManager;
    this._evaluationManager = this.evaluationManager;
    this._autoFillManager = this.autoFillManager;
    this._dependencyManager = this.dependencyManager;
    this._styleManager = this.styleManager;
    this._rangeMetadataManager = this.rangeMetadataManager;
  }

  /**
   * Static factory method to build an empty engine
   * @template TMetadata - Consumer-defined metadata shape with optional cell, sheet, and workbook entries.
   */
  static buildEmpty<TMetadata extends Metadata = Metadata>(
    options?: FormulaEngineOptions
  ) {
    return new FormulaEngine<TMetadata>(options);
  }

  undo(): boolean {
    if (!this.undoRedoManager.canUndo()) {
      return false;
    }

    const target = this.undoRedoManager.popUndo();
    if (!target) {
      return false;
    }

    const current = this.serializeHistorySnapshot();
    this.restoreHistorySnapshot(target);
    this.undoRedoManager.pushRedo(current);
    this.emitUpdate();
    return true;
  }

  redo(): boolean {
    if (!this.undoRedoManager.canRedo()) {
      return false;
    }

    const target = this.undoRedoManager.popRedo();
    if (!target) {
      return false;
    }

    const current = this.serializeHistorySnapshot();
    this.restoreHistorySnapshot(target);
    this.undoRedoManager.pushUndo(current);
    this.emitUpdate();
    return true;
  }

  canUndo(): boolean {
    return this.undoRedoManager.canUndo();
  }

  canRedo(): boolean {
    return this.undoRedoManager.canRedo();
  }

  getUndoRedoState(): UndoRedoState {
    return this.undoRedoManager.getState();
  }

  clearUndoRedoHistory(): void {
    this.undoRedoManager.clear();
  }

  transact<T>(callback: () => T): T {
    return this.withUndoRedoCheckpoint(callback);
  }

  private emitMutation(footprint: MutationInvalidation): void {
    this.evaluationManager.invalidateFromMutation(footprint);
    this.eventManager.emitUpdate();
  }

  private emitUpdate(): void {
    this.eventManager.emitUpdate();
  }

  private withUndoRedoCheckpoint<T>(callback: () => T): T {
    if (this.historyCheckpointDepth > 0) {
      return callback();
    }

    const before = this.serializeHistorySnapshot();
    this.historyCheckpointDepth++;
    try {
      const result = callback();
      const after = this.serializeHistorySnapshot();
      this.undoRedoManager.recordMutation(before, after);
      return result;
    } finally {
      this.historyCheckpointDepth--;
    }
  }

  private getExistingSheetContent(opts: {
    workbookName: string;
    sheetName: string;
  }): Map<string, SerializedCellValue> | undefined {
    const sheet = this.workbookManager.getSheet(opts);
    return sheet ? new Map(sheet.content) : undefined;
  }

  private materializeDefaultTableHeaders(
    table: TableDefinition
  ): MutationInvalidation["touchedCells"] {
    const changes: Array<{
      address: CellAddress;
      before: SerializedCellValue;
      after: SerializedCellValue;
    }> = [];

    for (const header of table.headers.values()) {
      const address = {
        workbookName: table.workbookName,
        sheetName: table.sheetName,
        rowIndex: table.start.rowIndex,
        colIndex: table.start.colIndex + header.index,
      };
      const before = this.workbookManager.getCellContent(address);
      if (before !== undefined && before !== "") {
        continue;
      }

      this.workbookManager.setCellContent(address, header.name);
      changes.push({ address, before, after: header.name });
    }

    return buildTouchedCells(changes);
  }

  private renameTableHeaderReferences(updates: TableHeaderUpdate[]): {
    changedFormulaCells: CellAddress[];
    resourceKeys: string[];
  } {
    const renamesByTable = new Map<
      TableDefinition,
      Map<string, string>
    >();
    for (const update of updates) {
      if (update.oldName === update.newName) {
        continue;
      }
      let renames = renamesByTable.get(update.table);
      if (!renames) {
        renames = new Map();
        renamesByTable.set(update.table, renames);
      }
      renames.set(update.oldName, update.newName);
    }

    const changedFormulaCells: CellAddress[] = [];
    const resourceKeys: string[] = [];
    for (const [table, columnRenames] of renamesByTable) {
      changedFormulaCells.push(
        ...this.workbookManager.updateAllFormulas(
          (formula, formulaAddress) =>
            renameTableColumnsInFormula({
              formula,
              tableName: table.name,
              tableWorkbookName: table.workbookName,
              formulaWorkbookName: formulaAddress.workbookName,
              columnRenames,
              includeImplicitReferences:
                this.tableManager.isCellInTable(formulaAddress) === table,
            })
        )
      );
      resourceKeys.push(
        getTableResourceKey({
          workbookName: table.workbookName,
          tableName: table.name,
        }),
        ...this.namedExpressionManager.updateAllNamedExpressions(
          (formula, scope) =>
            renameTableColumnsInFormula({
              formula,
              tableName: table.name,
              tableWorkbookName: table.workbookName,
              formulaWorkbookName: scope.workbookName,
              columnRenames,
            })
        )
      );
    }

    return { changedFormulaCells, resourceKeys };
  }

  private getWorkbookResourceKeys(workbookName: string): string[] {
    const resourceKeys = new Set<string>([getWorkbookResourceKey(workbookName)]);

    for (const sheetName of this.workbookManager
      .getWorkbooks()
      .get(workbookName)
      ?.sheets.keys() ?? []) {
      resourceKeys.add(getSheetResourceKey({ workbookName, sheetName }));
    }

    for (const tableName of this.tableManager.getTables(workbookName).keys()) {
      resourceKeys.add(getTableResourceKey({ workbookName, tableName }));
    }

    const namedExpressions = this.namedExpressionManager.getNamedExpressions();
    for (const name of namedExpressions.workbookExpressions
      .get(workbookName)
      ?.keys() ?? []) {
      resourceKeys.add(
        getNamedExpressionResourceKey({ expressionName: name, workbookName })
      );
    }
    for (const [sheetName, expressions] of namedExpressions.sheetExpressions.get(
      workbookName
    ) ?? []) {
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

  private getSheetResourceKeys(opts: {
    workbookName: string;
    sheetName: string;
  }): string[] {
    const resourceKeys = new Set<string>([
      getWorkbookResourceKey(opts.workbookName),
      getSheetResourceKey(opts),
    ]);

    for (const [tableName, table] of this.tableManager.getTables(
      opts.workbookName
    )) {
      if (table.sheetName === opts.sheetName) {
        resourceKeys.add(
          getTableResourceKey({ workbookName: opts.workbookName, tableName })
        );
      }
    }

    const sheetExpressions = this.namedExpressionManager
      .getNamedExpressions()
      .sheetExpressions.get(opts.workbookName)
      ?.get(opts.sheetName);
    for (const name of sheetExpressions?.keys() ?? []) {
      resourceKeys.add(
        getNamedExpressionResourceKey({
          expressionName: name,
          workbookName: opts.workbookName,
          sheetName: opts.sheetName,
        })
      );
    }

    return Array.from(resourceKeys);
  }

  private assertNamedExpressionScopeExists(opts: {
    workbookName?: string;
    sheetName?: string;
  }): void {
    if (opts.sheetName && !opts.workbookName) {
      throw new Error("Missing workbookName");
    }

    if (!opts.workbookName) {
      return;
    }

    if (!this.workbookManager.getWorkbooks().has(opts.workbookName)) {
      throw new Error(`Workbook not found: ${opts.workbookName}`);
    }

    if (
      opts.sheetName &&
      !this.workbookManager.getSheet({
        workbookName: opts.workbookName,
        sheetName: opts.sheetName,
      })
    ) {
      throw new Error(`Sheet not found: ${opts.sheetName}`);
    }
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

  /**
   * Set metadata for a cell
   * Metadata can contain rich text, links, comments, or any consumer-defined data
   */
  setCellMetadata(
    address: CellAddress,
    metadata: MetadataType<TMetadata, "cell"> | undefined
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.setCellMetadata(address, metadata);
      this.emitUpdate();
    });
  }

  /**
   * Get metadata for a cell
   */
  getCellMetadata(
    address: CellAddress
  ): MetadataType<TMetadata, "cell"> | undefined {
    const metadata = this.workbookManager.getCellMetadata(address);
    return metadata as MetadataType<TMetadata, "cell"> | undefined;
  }

  /**
   * Get all cell metadata for a sheet (serialized as Map)
   */
  getSheetMetadataSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, MetadataType<TMetadata, "sheet">> {
    return this.workbookManager.getSheetMetadataSerialized(opts) as Map<
      string,
      MetadataType<TMetadata, "sheet">
    >;
  }

  /**
   * Set metadata for a sheet
   * Sheet metadata can contain text boxes, frozen panes, print settings, or any consumer-defined data
   */
  setSheetMetadata(
    opts: { workbookName: string; sheetName: string },
    metadata: MetadataType<TMetadata, "sheet">
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.setSheetMetadata(opts, metadata);
      this.emitUpdate();
    });
  }

  /**
   * Get metadata for a sheet
   */
  getSheetMetadata(opts: {
    workbookName: string;
    sheetName: string;
  }): MetadataType<TMetadata, "sheet"> | undefined {
    return this.workbookManager.getSheetMetadata(opts) as
      | MetadataType<TMetadata, "sheet">
      | undefined;
  }

  /**
   * Set metadata for a workbook
   * Workbook metadata can contain themes, document properties, settings, or any consumer-defined data
   */
  setWorkbookMetadata(
    workbookName: string,
    metadata: MetadataType<TMetadata, "workbook">
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.setWorkbookMetadata(workbookName, metadata);
      this.emitUpdate();
    });
  }

  /**
   * Get metadata for a workbook
   */
  getWorkbookMetadata(
    workbookName: string
  ): MetadataType<TMetadata, "workbook"> | undefined {
    return this.workbookManager.getWorkbookMetadata(workbookName) as
      | MetadataType<TMetadata, "workbook">
      | undefined;
  }

  /**
   * Add metadata attached to one or more ranges.
   * Range metadata is arbitrary consumer-defined data that follows range
   * copy/paste, autofill, workbook/sheet rename, and serialization flows.
   */
  addRangeMetadata(
    metadata: RangeMetadataInput<MetadataType<TMetadata, "range">>
  ): string {
    return this.withUndoRedoCheckpoint(() => {
      const id = this.rangeMetadataManager.addRangeMetadata(metadata);
      this.emitUpdate();
      return id;
    });
  }

  /**
   * Remove a range metadata entry by id.
   */
  removeRangeMetadata(id: string): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.rangeMetadataManager.removeRangeMetadata(id);
      if (removed) {
        this.emitUpdate();
      }
    });
  }

  /**
   * Get all range metadata entries.
   */
  getAllRangeMetadata(): RangeMetadata<MetadataType<TMetadata, "range">>[] {
    return this.rangeMetadataManager.getAllRangeMetadata();
  }

  /**
   * Get range metadata entries that apply to a specific cell.
   */
  getRangeMetadataForCell(
    address: CellAddress
  ): RangeMetadata<MetadataType<TMetadata, "range">>[] {
    return this.rangeMetadataManager.getRangeMetadataForCell(address);
  }

  /**
   * Get range metadata entries intersecting with a range.
   */
  getRangeMetadataIntersectingWithRange(
    range: RangeAddress
  ): RangeMetadata<MetadataType<TMetadata, "range">>[] {
    return this.rangeMetadataManager.getRangeMetadataIntersectingWithRange(
      range
    );
  }

  /**
   * Clear range metadata from a range, preserving non-overlapping portions.
   */
  clearRangeMetadata(range: RangeAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      this.rangeMetadataManager.clearRangeMetadataInRange(range);
      this.emitUpdate();
    });
  }

  //#endregion

  //#region Reference Tracking
  /**
   * Create a tracked reference to a range
   * Returns a stable UUID that can be used to retrieve the address later
   * The reference automatically updates when sheets/workbooks are renamed
   */
  createRef(address: RangeAddress): string {
    return this.withUndoRedoCheckpoint(() => {
      const id = this.referenceManager.createRef(address);
      this.emitUpdate();
      return id;
    });
  }

  /**
   * Get the current address for a tracked reference
   * Returns undefined if reference doesn't exist or has been invalidated
   */
  getRefAddress(refId: string): RangeAddress | undefined {
    return this.referenceManager.getRefAddress(refId);
  }

  /**
   * Delete a tracked reference
   * Returns true if the reference was deleted, false if it didn't exist
   */
  deleteRef(refId: string): boolean {
    return this.withUndoRedoCheckpoint(() => {
      const deleted = this.referenceManager.deleteRef(refId);
      if (deleted) {
        this.emitUpdate();
      }
      return deleted;
    });
  }

  /**
   * Get all invalid reference IDs
   * Useful for cleanup after sheet/workbook deletions
   */
  getInvalidRefs(): string[] {
    return this.referenceManager.getInvalidRefs();
  }
  //#endregion

  evaluateFormula(
    /**
     * formula without the leading = sign
     */
    formula: string,
    cellAddress: CellAddress
  ): SerializedCellValue {
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
  addNamedExpression(opts: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }) {
    return this.withUndoRedoCheckpoint(() => {
      this.assertNamedExpressionScopeExists(opts);
      this.namedExpressionManager.addNamedExpression(opts);
      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getNamedExpressionResourceKey(opts)],
      });
    });
  }

  removeNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.namedExpressionManager.removeNamedExpression(opts);
      if (removed) {
        this.emitMutation({
          touchedCells: [],
          resourceKeys: [getNamedExpressionResourceKey(opts)],
        });
      }
    });
  }

  /**
   * Check if a named expression exists
   */
  hasNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): boolean {
    const scope =
      opts.sheetName && opts.workbookName
        ? {
            type: "sheet" as const,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }
        : opts.workbookName
        ? { type: "workbook" as const, workbookName: opts.workbookName }
        : { type: "global" as const };

    return !!this.namedExpressionManager.getNamedExpression({
      name: opts.expressionName,
      scope,
    });
  }

  updateNamedExpression(opts: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      this.namedExpressionManager.updateNamedExpression(opts);
      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getNamedExpressionResourceKey(opts)],
      });
    });
  }

  renameNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
    newName: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      this.namedExpressionManager.renameNamedExpression(opts);

      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameNamedExpressionInFormula(
          formula,
          opts.expressionName,
          opts.newName
        )
      );

      const changedNamedExpressions =
        this.namedExpressionManager.updateAllNamedExpressions((formula) =>
          renameNamedExpressionInFormula(
            formula,
            opts.expressionName,
            opts.newName
          )
        );

      this.emitMutation({
        touchedCells: buildFormulaTouchedCells(changedCells),
        resourceKeys: [
          getNamedExpressionResourceKey({
            expressionName: opts.expressionName,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }),
          getNamedExpressionResourceKey({
            expressionName: opts.newName,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }),
          ...changedNamedExpressions,
        ],
      });
    });
  }

  setNamedExpressions(
    opts: (
      | { type: "global" }
      | { type: "sheet"; sheetName: string; workbookName: string }
      | { type: "workbook"; workbookName: string }
    ) & {
      expressions: Map<string, NamedExpression>;
    }
  ) {
    return this.withUndoRedoCheckpoint(() => {
      const allExpressions = this.namedExpressionManager.getNamedExpressions();
      let previousExpressions: Map<string, NamedExpression> | undefined;

      if (opts.type === "global") {
        previousExpressions = new Map(allExpressions.globalExpressions);
      } else if (opts.type === "workbook") {
        previousExpressions = new Map(
          allExpressions.workbookExpressions.get(opts.workbookName) || []
        );
      } else {
        const sheetExpressions = allExpressions.sheetExpressions
          .get(opts.workbookName)
          ?.get(opts.sheetName);
        previousExpressions = new Map(sheetExpressions || []);
      }

      this.namedExpressionManager.setNamedExpressions(opts);

      const scope =
        opts.type === "global"
          ? {}
          : opts.type === "workbook"
          ? { workbookName: opts.workbookName }
          : {
              workbookName: opts.workbookName,
              sheetName: opts.sheetName,
            };

      this.emitMutation({
        touchedCells: [],
        resourceKeys: [
          ...getNamedExpressionScopeResourceKeys(
            previousExpressions.keys(),
            scope
          ),
          ...getNamedExpressionScopeResourceKeys(opts.expressions.keys(), scope),
        ],
      });
    });
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
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const table = this.tableManager.addTable({
        ...props,
        getCellValue: (cellAddress: CellAddress) =>
          this.getCellValue(cellAddress),
      });
      const generatedHeaderCells = this.materializeDefaultTableHeaders(table);

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, [table]),
          generatedHeaderCells
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [table]
        ),
        resourceKeys: [
          getTableResourceKey({
            workbookName: props.workbookName,
            tableName: props.tableName,
          }),
        ],
      });
    });
  }

  renameTable(
    workbookName: string,
    names: { oldName: string; newName: string }
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTable = this.tableManager.getTable({
        workbookName,
        name: names.oldName,
      });
      const oldTableSnapshot = oldTable
        ? { ...oldTable, headers: new Map(oldTable.headers) }
        : undefined;

      this.tableManager.renameTable(workbookName, names);

      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameTableInFormula(formula, names.oldName, names.newName)
      );

      const changedNamedExpressions =
        this.namedExpressionManager.updateAllNamedExpressions((formula) =>
          renameTableInFormula(formula, names.oldName, names.newName)
        );

      const newTable = this.tableManager.getTable({
        workbookName,
        name: names.newName,
      });

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, [oldTableSnapshot]),
          buildTableTouchedCells(this.workbookManager, [newTable]),
          buildFormulaTouchedCells(changedCells)
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [oldTableSnapshot, newTable]
        ),
        resourceKeys: [
          getTableResourceKey({
            workbookName,
            tableName: names.oldName,
          }),
          getTableResourceKey({
            workbookName,
            tableName: names.newName,
          }),
          ...changedNamedExpressions,
        ],
      });
    });
  }

  updateTable(opts: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
    workbookName: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTable = this.tableManager.getTable({
        workbookName: opts.workbookName,
        name: opts.tableName,
      });
      const oldTableSnapshot = oldTable
        ? { ...oldTable, headers: new Map(oldTable.headers) }
        : undefined;

      this.tableManager.updateTable({
        ...opts,
        getCellValue: (cellAddress: CellAddress) =>
          this.getCellValue(cellAddress),
      });

      const newTable = this.tableManager.getTable({
        workbookName: opts.workbookName,
        name: opts.tableName,
      });
      const generatedHeaderCells = newTable
        ? this.materializeDefaultTableHeaders(newTable)
        : [];

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, [oldTableSnapshot]),
          buildTableTouchedCells(this.workbookManager, [newTable]),
          generatedHeaderCells
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [oldTableSnapshot, newTable]
        ),
        resourceKeys: [
          getTableResourceKey({
            workbookName: opts.workbookName,
            tableName: opts.tableName,
          }),
        ],
      });
    });
  }

  removeTable(opts: { tableName: string; workbookName: string }): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTable = this.tableManager.getTable({
        workbookName: opts.workbookName,
        name: opts.tableName,
      });
      const oldTableSnapshot = oldTable
        ? { ...oldTable, headers: new Map(oldTable.headers) }
        : undefined;

      const found = this.tableManager.removeTable(opts);
      if (found) {
        this.emitMutation({
          touchedCells: buildTableTouchedCells(this.workbookManager, [
            oldTableSnapshot,
          ]),
          tableContextChangedCells: buildTableContextChangedCells(
            this.workbookManager,
            [oldTableSnapshot]
          ),
          resourceKeys: [
            getTableResourceKey({
              workbookName: opts.workbookName,
              tableName: opts.tableName,
            }),
          ],
        });
      }
    });
  }

  private getAllTables(): TableDefinition[] {
    return Array.from(this.tableManager.tables.values()).flatMap((tables) =>
      Array.from(tables.values()).map((table) => ({
        ...table,
        headers: new Map(table.headers),
      }))
    );
  }

  /**
   * Check if a table exists
   */
  hasTable(opts: { tableName: string; workbookName: string }): boolean {
    return !!this.tableManager.getTable({
      workbookName: opts.workbookName,
      name: opts.tableName,
    });
  }

  /**
   * Get a table definition by name
   */
  getTable(opts: {
    tableName: string;
    workbookName: string;
  }): TableDefinition | undefined {
    return this.tableManager.getTable({
      workbookName: opts.workbookName,
      name: opts.tableName,
    });
  }

  resetTables(tables: Map<string, Map<string, TableDefinition>>): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTables = this.getAllTables();
      const newTables = Array.from(tables.values()).flatMap((workbookTables) =>
        Array.from(workbookTables.values())
      );
      const resourceKeys = new Set<string>();
      for (const table of [...oldTables, ...newTables]) {
        resourceKeys.add(
          getTableResourceKey({
            workbookName: table.workbookName,
            tableName: table.name,
          })
        );
      }

      this.tableManager.resetTables(tables);
      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, oldTables),
          buildTableTouchedCells(this.workbookManager, newTables)
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [...oldTables, ...newTables]
        ),
        resourceKeys: Array.from(resourceKeys),
      });
    });
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
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.addConditionalStyle(style);
      this.emitUpdate();
    });
  }

  /**
   * Remove a conditional style rule by index
   */
  removeConditionalStyle(workbookName: string, index: number): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.styleManager.removeConditionalStyle(
        workbookName,
        index
      );
      if (removed) {
        this.emitUpdate();
      }
    });
  }

  /**
   * Get the count of conditional styles for a workbook
   */
  getConditionalStyleCount(workbookName: string): number {
    const allStyles = this.styleManager.getAllConditionalStyles();
    return allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === workbookName)
    ).length;
  }

  /**
   * Get all conditional styles intersecting with a range
   */
  getConditionalStylesIntersectingWithRange(
    range: RangeAddress
  ): ConditionalStyle[] {
    return this.styleManager.getConditionalStylesIntersectingWithRange(range);
  }

  /**
   * Get the computed style for a specific cell
   */
  getCellStyle(cellAddress: CellAddress): CellStyle | undefined {
    return this.styleManager.getCellStyle(cellAddress);
  }

  getCellDataType(cellAddress: CellAddress): CellDataType {
    return this.styleManager.getCellDataType(cellAddress);
  }

  getDataTypeForRange(range: RangeAddress): CellDataType | undefined {
    return this.styleManager.getDataTypeForRange(range);
  }

  getAllCellDataTypes(): DirectCellDataType[] {
    return this.styleManager.getAllCellDataTypes();
  }

  addCellDataType(dataType: DirectCellDataType): void {
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.addCellDataType(dataType);
      this.emitUpdate();
    });
  }

  clearCellDataTypes(range: RangeAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.clearCellDataTypesInRange(range);
      this.emitUpdate();
    });
  }

  /**
   * Get all cell styles (for testing and serialization)
   */
  getAllCellStyles(): DirectCellStyle[] {
    return this.styleManager.getAllCellStyles();
  }

  /**
   * Get all conditional styles (for testing and serialization)
   */
  getAllConditionalStyles(): ConditionalStyle[] {
    return this.styleManager.getAllConditionalStyles();
  }

  /**
   * Add a direct cell style rule
   */
  addCellStyle(style: DirectCellStyle): void {
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.addCellStyle(style);
      this.emitUpdate();
    });
  }

  /**
   * Remove a direct cell style rule by index
   */
  removeCellStyle(workbookName: string, index: number): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.styleManager.removeCellStyle(workbookName, index);
      if (removed) {
        this.emitUpdate();
      }
    });
  }

  /**
   * Get the count of direct cell styles for a workbook
   */
  getCellStyleCount(workbookName: string): number {
    const allStyles = this.styleManager.getAllCellStyles();
    return allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === workbookName)
    ).length;
  }

  /**
   * Get all direct cell styles intersecting with a range
   */
  getStylesIntersectingWithRange(range: RangeAddress): DirectCellStyle[] {
    return this.styleManager.getStylesIntersectingWithRange(range);
  }

  /**
   * Get the style for a range if all cells in the range have the same style
   * Returns the DirectCellStyle if the range is completely contained within a single style's area
   * Returns undefined if multiple styles, partial coverage, or no styles apply
   */
  getStyleForRange(range: RangeAddress): DirectCellStyle | undefined {
    return this.styleManager.getStyleForRange(range);
  }

  /**
   * Clear all cell styles and conditional styles for a given range
   * Adjusts existing style ranges rather than deleting them entirely
   */
  clearCellStyles(range: RangeAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.clearCellStyles(range);
      this.emitUpdate();
    });
  }

  //#endregion

  //#region Copy/Paste
  private getTopLeftCell(cells: CellAddress[]): CellAddress {
    let topLeft = cells[0]!;
    for (const cell of cells) {
      if (
        cell.rowIndex < topLeft.rowIndex ||
        (cell.rowIndex === topLeft.rowIndex &&
          cell.colIndex < topLeft.colIndex)
      ) {
        topLeft = cell;
      }
    }
    return topLeft;
  }

  private dedupeAddresses(addresses: CellAddress[]): CellAddress[] {
    return Array.from(
      new Map(addresses.map((address) => [getMutationAddressKey(address), address]))
        .values()
    );
  }

  private getPasteTouchedAddresses(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): CellAddress[] {
    if (source.length === 0) {
      return [];
    }

    const topLeft = this.getTopLeftCell(source);
    const colOffset = target.colIndex - topLeft.colIndex;
    const rowOffset = target.rowIndex - topLeft.rowIndex;
    const targetCells = source.map((sourceCell) => ({
      workbookName: target.workbookName,
      sheetName: target.sheetName,
      colIndex: sourceCell.colIndex + colOffset,
      rowIndex: sourceCell.rowIndex + rowOffset,
    }));

    return this.dedupeAddresses(
      options.cut ? [...source, ...targetCells] : targetCells
    );
  }

  /**
   * Paste cells from source to target
   */
  pasteCells(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      if (source.length === 0) {
        return;
      }

      const touchedAddresses = this.getPasteTouchedAddresses(
        source,
        target,
        options
      );
      const before = captureCellContents(this.workbookManager, touchedAddresses);

      this.copyManager.pasteCells(source, target, options);

      const after = captureCellContents(this.workbookManager, touchedAddresses);
      this.emitMutation({
        touchedCells: buildTouchedCells(
          touchedAddresses.map((address) => ({
            address,
            before: before.get(getMutationAddressKey(address)),
            after: after.get(getMutationAddressKey(address)),
          }))
        ),
        resourceKeys: [],
      });
    });
  }

  /**
   * Fill one or more areas with a seed range's content/style
   * Uses column-first strategy: fills down, then replicates right
   * Formulas are adjusted based on each target cell's offset from the seed
   *
   * @param seedRange - The range to use as a template/pattern
   * @param targetRanges - One or more range addresses to fill
   * @param options - Copy options (target: 'all'|'content'|'style', type: 'value'|'formula', cut: boolean)
   *
   * @example
   * // Fill F6:J10 with A1:B2 seed (2x2 pattern fills 5x5 area)
   * engine.fillAreas(
   *   {
   *     workbookName,
   *     sheetName,
   *     range: {
   *       start: { col: 0, row: 0 },
   *       end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
   *     }
   *   },
   *   [{
   *     workbookName,
   *     sheetName,
   *     range: {
   *       start: { col: 5, row: 5 },
   *       end: { col: { type: "number", value: 9 }, row: { type: "number", value: 9 } }
   *     }
   *   }],
   *   { cut: false, type: "formula", target: "all" }
   * );
   */
  fillAreas(
    seedRange: RangeAddress,
    targetRanges: RangeAddress[],
    options: CopyCellsOptions
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      const touchedAddresses = this.dedupeAddresses([
        ...targetRanges.flatMap((targetRange) =>
          getFiniteRangeAddresses(targetRange)
        ),
        ...(options.cut ? getFiniteRangeAddresses(seedRange) : []),
      ]);
      const before = captureCellContents(this.workbookManager, touchedAddresses);

      this.copyManager.fillAreas(seedRange, targetRanges, options);

      const after = captureCellContents(this.workbookManager, touchedAddresses);
      this.emitMutation({
        touchedCells: buildTouchedCells(
          touchedAddresses.map((address) => ({
            address,
            before: before.get(getMutationAddressKey(address)),
            after: after.get(getMutationAddressKey(address)),
          }))
        ),
        resourceKeys: [],
      });
    });
  }

  /**
   * Smart paste that automatically determines whether to paste or fill
   * Handles multiple selection areas - each area is independently pasted or filled
   * - If area is larger than source, uses fillAreas() to fill the area
   * - If area is same size or smaller, uses pasteCells() for normal paste
   *
   * @param sourceCells - The copied cells
   * @param pasteSelection - One or more selection areas where user is pasting
   * @param options - Copy options
   *
   * @example
   * // Copy A1, paste into two areas B1:C2 and E5:F6 - both get filled
   * engine.smartPaste(
   *   [{ workbookName, sheetName, colIndex: 0, rowIndex: 0 }],
   *   {
   *     workbookName,
   *     sheetName,
   *     areas: [
   *       { start: { col: 1, row: 0 }, end: { col: { type: "number", value: 2 }, row: { type: "number", value: 1 } } },
   *       { start: { col: 4, row: 4 }, end: { col: { type: "number", value: 5 }, row: { type: "number", value: 5 } } }
   *     ]
   *   },
   *   { cut: false, type: "formula", target: "all" }
   * );
   */
  smartPaste(
    sourceCells: CellAddress[],
    pasteSelection: {
      workbookName: string;
      sheetName: string;
      areas: SpreadsheetRange[];
    },
    options: CopyCellsOptions
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      if (sourceCells.length === 0) {
        return;
      }

      // If cut operation, always use pasteCells (never fillAreas)
      // Cut should be a simple move operation, not a fill
      if (options.cut === true) {
        for (const area of pasteSelection.areas) {
          const target: CellAddress = {
            workbookName: pasteSelection.workbookName,
            sheetName: pasteSelection.sheetName,
            colIndex: area.start.col,
            rowIndex: area.start.row,
          };
          this.pasteCells(sourceCells, target, options);
        }
        return;
      }

      // For copy operations (not cut), use smart paste/fill logic
      // Calculate source bounds once
      const sourceBounds = this.getBoundsFromCells(sourceCells);
      const sourceWidth = sourceBounds.maxCol - sourceBounds.minCol + 1;
      const sourceHeight = sourceBounds.maxRow - sourceBounds.minRow + 1;

      // Create seed range for fill operations
      const seedRange: RangeAddress = {
        workbookName: sourceCells[0]!.workbookName,
        sheetName: sourceCells[0]!.sheetName,
        range: {
          start: { col: sourceBounds.minCol, row: sourceBounds.minRow },
          end: {
            col: { type: "number", value: sourceBounds.maxCol },
            row: { type: "number", value: sourceBounds.maxRow },
          },
        },
      };

      // Process each selected area independently
      for (const area of pasteSelection.areas) {
        const pasteStartCol = area.start.col;
        const pasteStartRow = area.start.row;
        const pasteEndCol =
          area.end.col.type === "number" ? area.end.col.value : pasteStartCol;
        const pasteEndRow =
          area.end.row.type === "number" ? area.end.row.value : pasteStartRow;

        const pasteWidth = pasteEndCol - pasteStartCol + 1;
        const pasteHeight = pasteEndRow - pasteStartRow + 1;

        // Decide per area: paste or fill?
        const shouldFill =
          pasteWidth > sourceWidth || pasteHeight > sourceHeight;

        if (shouldFill) {
          // Use fillAreas for this area
          const targetRange: RangeAddress = {
            workbookName: pasteSelection.workbookName,
            sheetName: pasteSelection.sheetName,
            range: {
              start: { col: pasteStartCol, row: pasteStartRow },
              end: {
                col: { type: "number", value: pasteEndCol },
                row: { type: "number", value: pasteEndRow },
              },
            },
          };

          this.fillAreas(seedRange, [targetRange], options);
        } else {
          // Use pasteCells for this area
          const target: CellAddress = {
            workbookName: pasteSelection.workbookName,
            sheetName: pasteSelection.sheetName,
            colIndex: pasteStartCol,
            rowIndex: pasteStartRow,
          };

          this.pasteCells(sourceCells, target, options);
        }
      }
    });
  }

  /**
   * Get bounds (min/max row/col) from an array of cell addresses
   */
  private getBoundsFromCells(cells: CellAddress[]): {
    minCol: number;
    minRow: number;
    maxCol: number;
    maxRow: number;
  } {
    if (cells.length === 0) {
      throw new Error("Cannot get bounds from empty cell array");
    }

    let minCol = Infinity;
    let minRow = Infinity;
    let maxCol = -Infinity;
    let maxRow = -Infinity;

    for (const cell of cells) {
      minCol = Math.min(minCol, cell.colIndex);
      minRow = Math.min(minRow, cell.rowIndex);
      maxCol = Math.max(maxCol, cell.colIndex);
      maxRow = Math.max(maxRow, cell.rowIndex);
    }

    return { minCol, minRow, maxCol, maxRow };
  }

  /**
   * Move a single cell to a new location
   * Updates all formula references that point to the moved cell
   *
   * @param source - The cell to move
   * @param target - The destination cell address
   *
   * @example
   * // Move A1 to D5. If B1 contains =A1, it will be updated to =D5
   * engine.moveCell(
   *   { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
   *   { workbookName, sheetName, colIndex: 3, rowIndex: 4 }
   * );
   */
  moveCell(source: CellAddress, target: CellAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      this.pasteCells([source], target, {
        cut: true,
        type: "formula",
        include: "all",
      });
    });
  }

  /**
   * Move a range of cells to a new location
   * Updates all formula references that point to the moved cells
   *
   * @param sourceRange - The range to move
   * @param target - The top-left destination cell address
   *
   * @example
   * // Move A1:D5 to F10. If E1 contains =SUM(A1:D5), it will be updated to =SUM(F10:I14)
   * engine.moveRange(
   *   {
   *     workbookName,
   *     sheetName,
   *     range: {
   *       start: { col: 0, row: 0 },
   *       end: { col: { type: "number", value: 3 }, row: { type: "number", value: 4 } }
   *     }
   *   },
   *   { workbookName, sheetName, colIndex: 5, rowIndex: 9 }
   * );
   */
  moveRange(sourceRange: RangeAddress, target: CellAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      const cells = this.copyManager.expandRangeToCells(sourceRange);
      this.pasteCells(cells, target, {
        cut: true,
        type: "formula",
        include: "all",
      });
    });
  }
  //#endregion

  //#region Sheets
  addSheet(opts: { workbookName: string; sheetName: string }): Sheet {
    return this.withUndoRedoCheckpoint(() => {
      const sheet = this.workbookManager.addSheet(opts);
      this.namedExpressionManager.addSheet(opts);
      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getSheetResourceKey(opts)],
      });
      return sheet;
    });
  }

  createSheet(opts: {
    workbookName: string;
    sheetName?: string;
    baseName?: string;
  }): Sheet {
    return this.withUndoRedoCheckpoint(() => {
      const sheetName =
        opts.sheetName ??
        this.workbookManager.getNextAvailableSheetName(
          opts.workbookName,
          opts.baseName
        );

      return this.addSheet({
        workbookName: opts.workbookName,
        sheetName,
      });
    });
  }

  removeSheet(opts: { workbookName: string; sheetName: string }): void {
    return this.withUndoRedoCheckpoint(() => {
      const resourceKeys = this.getSheetResourceKeys(opts);
      this.workbookManager.removeSheet(opts);
      this.namedExpressionManager.removeSheet(opts);
      this.tableManager.removeSheet(opts);
      this.styleManager.removeSheetStyles(opts.workbookName, opts.sheetName);
      this.rangeMetadataManager.removeSheetRangeMetadata(
        opts.workbookName,
        opts.sheetName
      );
      this.referenceManager.invalidateSheet(opts.workbookName, opts.sheetName);
      this.emitMutation({
        touchedCells: [],
        resourceKeys,
        removedScopes: [{ type: "sheet", ...opts }],
      });
    });
  }

  renameSheet(opts: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldResourceKeys = this.getSheetResourceKeys(opts);

      this.workbookManager.renameSheet(opts);
      this.namedExpressionManager.renameSheet(opts);
      this.tableManager.updateTablesForSheetRename(opts);
      this.styleManager.updateSheetName(
        opts.workbookName,
        opts.sheetName,
        opts.newSheetName
      );
      this.rangeMetadataManager.updateSheetName(
        opts.workbookName,
        opts.sheetName,
        opts.newSheetName
      );
      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameSheetInFormula({
          formula,
          oldSheetName: opts.sheetName,
          newSheetName: opts.newSheetName,
        })
      );
      this.referenceManager.updateSheetName(
        opts.workbookName,
        opts.sheetName,
        opts.newSheetName
      );

      this.emitMutation({
        touchedCells: buildFormulaTouchedCells(changedCells),
        resourceKeys: Array.from(
          new Set([
            ...oldResourceKeys,
            ...this.getSheetResourceKeys({
              workbookName: opts.workbookName,
              sheetName: opts.newSheetName,
            }),
          ])
        ),
      });
    });
  }

  /**
   * Check if a sheet exists
   */
  hasSheet(opts: { workbookName: string; sheetName: string }): boolean {
    return !!this.workbookManager.getSheet(opts);
  }

  getSheets(workbookName: string) {
    return this.workbookManager.getSheets(workbookName);
  }

  getOrderedSheets(workbookName: string) {
    return this.workbookManager.getOrderedSheets(workbookName);
  }

  getOrderedSheetNames(workbookName: string) {
    return this.workbookManager.getOrderedSheetNames(workbookName);
  }

  getNextAvailableSheetName(workbookName: string, baseName?: string) {
    return this.workbookManager.getNextAvailableSheetName(
      workbookName,
      baseName
    );
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

  /**
   * Search raw stored string content without evaluating cell values.
   */
  search(
    query: string,
    options?: SearchOptions
  ): SearchMatch[] {
    return this.workbookManager.search(query, options);
  }

  /**
   * Replace one specific search occurrence inside one addressed cell.
   */
  replace(
    query: string,
    replacement: string,
    target: ReplaceTarget,
    options?: { caseSensitive?: boolean }
  ): ReplaceChange {
    return this.withUndoRedoCheckpoint(() => {
      const prepared = this.workbookManager.prepareReplace(
        query,
        replacement,
        target,
        options
      );

      this.workbookManager.setCellContent(
        prepared.address,
        prepared.afterContent
      );
      this.emitMutation({
        touchedCells: buildTouchedCells([
          {
            address: prepared.address,
            before: prepared.beforeContent,
            after: prepared.afterContent,
          },
        ]),
        resourceKeys: [],
      });

      return prepared.change;
    });
  }

  /**
   * Replace all matching raw string occurrences within the requested scope.
   */
  replaceAll(
    query: string,
    replacement: string,
    options?: SearchOptions
  ): ReplaceChange[] {
    return this.withUndoRedoCheckpoint(() => {
      const preparedReplacements = this.workbookManager.prepareReplaceAll(
        query,
        replacement,
        options
      );

      if (preparedReplacements.length === 0) {
        return [];
      }

      for (const prepared of preparedReplacements) {
        this.workbookManager.setCellContent(
          prepared.address,
          prepared.afterContent
        );
      }

      this.emitMutation({
        touchedCells: buildTouchedCells(
          preparedReplacements.map((prepared) => ({
            address: prepared.address,
            before: prepared.beforeContent,
            after: prepared.afterContent,
          }))
        ),
        resourceKeys: [],
      });

      return preparedReplacements.flatMap((prepared) => prepared.changes);
    });
  }

  //#endregion

  //#region Workbook
  addWorkbook(workbookName: string): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.addWorkbook(workbookName);
      this.namedExpressionManager.addWorkbook(workbookName);
      this.tableManager.addWorkbook(workbookName);
      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getWorkbookResourceKey(workbookName)],
      });
    });
  }

  removeWorkbook(workbookName: string): void {
    return this.withUndoRedoCheckpoint(() => {
      const resourceKeys = this.getWorkbookResourceKeys(workbookName);
      this.workbookManager.removeWorkbook(workbookName);
      this.namedExpressionManager.removeWorkbook(workbookName);
      this.tableManager.removeWorkbook(workbookName);
      this.styleManager.removeWorkbookStyles(workbookName);
      this.rangeMetadataManager.removeWorkbookRangeMetadata(workbookName);
      this.referenceManager.invalidateWorkbook(workbookName);
      this.emitMutation({
        touchedCells: [],
        resourceKeys,
        removedScopes: [{ type: "workbook", workbookName }],
      });
    });
  }

  /**
   * Check if a workbook exists
   */
  hasWorkbook(workbookName: string): boolean {
    return this.workbookManager.getWorkbooks().has(workbookName);
  }

  cloneWorkbook(fromWorkbookName: string, toWorkbookName: string): void {
    return this.withUndoRedoCheckpoint(() => {
      const sourceWorkbook = this.workbookManager
        .getWorkbooks()
        .get(fromWorkbookName);
      if (!sourceWorkbook) {
        throw new Error(`Source workbook "${fromWorkbookName}" not found`);
      }
      if (this.workbookManager.getWorkbooks().has(toWorkbookName)) {
        throw new Error(`Target workbook "${toWorkbookName}" already exists`);
      }

      this.workbookManager.addWorkbook(toWorkbookName);
      this.namedExpressionManager.addWorkbook(toWorkbookName);
      this.tableManager.addWorkbook(toWorkbookName);

      for (const [sheetName, sheet] of sourceWorkbook.sheets) {
        this.workbookManager.addSheet({
          workbookName: toWorkbookName,
          sheetName,
        });
        this.namedExpressionManager.addSheet({
          workbookName: toWorkbookName,
          sheetName,
        });
        this.workbookManager.setSheetContent(
          { workbookName: toWorkbookName, sheetName },
          new Map(sheet.content)
        );

        const targetSheet = this.workbookManager.getSheet({
          workbookName: toWorkbookName,
          sheetName,
        });
        if (targetSheet) {
          targetSheet.metadata = new Map(sheet.metadata);
          if (sheet.sheetMetadata !== undefined) {
            targetSheet.sheetMetadata = structuredClone(sheet.sheetMetadata);
          }
        }
      }

      const targetWorkbook = this.workbookManager
        .getWorkbooks()
        .get(toWorkbookName);
      if (targetWorkbook && sourceWorkbook.workbookMetadata !== undefined) {
        targetWorkbook.workbookMetadata = structuredClone(
          sourceWorkbook.workbookMetadata
        );
      }

      const namedExpressions = this.namedExpressionManager.getNamedExpressions();
      const sourceWorkbookExpressions =
        namedExpressions.workbookExpressions.get(fromWorkbookName);
      if (sourceWorkbookExpressions) {
        for (const [expressionName, expression] of sourceWorkbookExpressions) {
          this.namedExpressionManager.addNamedExpression({
            expressionName,
            expression: expression.expression,
            workbookName: toWorkbookName,
          });
        }
      }

      const sourceSheetExpressions =
        namedExpressions.sheetExpressions.get(fromWorkbookName);
      if (sourceSheetExpressions) {
        for (const [sheetName, expressions] of sourceSheetExpressions) {
          for (const [expressionName, expression] of expressions) {
            this.namedExpressionManager.addNamedExpression({
              expressionName,
              expression: expression.expression,
              workbookName: toWorkbookName,
              sheetName,
            });
          }
        }
      }

      const sourceTables = this.tableManager.tables.get(fromWorkbookName);
      if (sourceTables) {
        for (const [tableName] of sourceTables) {
          this.tableManager.copyTable(
            { workbookName: fromWorkbookName, tableName },
            { workbookName: toWorkbookName, tableName }
          );
        }
      }

      for (const style of this.styleManager.getAllConditionalStyles()) {
        if (style.areas.some((area) => area.workbookName === fromWorkbookName)) {
          this.styleManager.addConditionalStyle({
            ...style,
            areas: style.areas.map((area) =>
              area.workbookName === fromWorkbookName
                ? { ...area, workbookName: toWorkbookName }
                : area
            ),
          });
        }
      }

      for (const style of this.styleManager.getAllCellStyles()) {
        if (style.areas.some((area) => area.workbookName === fromWorkbookName)) {
          this.styleManager.addCellStyle({
            ...style,
            areas: style.areas.map((area) =>
              area.workbookName === fromWorkbookName
                ? { ...area, workbookName: toWorkbookName }
                : area
            ),
          });
        }
      }

      for (const dataType of this.styleManager.getAllCellDataTypes()) {
        const clonedAreas = dataType.areas
          .filter((area) => area.workbookName === fromWorkbookName)
          .map((area) => ({ ...area, workbookName: toWorkbookName }));
        if (clonedAreas.length > 0) {
          this.styleManager.addCellDataType({
            ...dataType,
            areas: clonedAreas,
          });
        }
      }

      for (const metadata of this.rangeMetadataManager.getAllRangeMetadata()) {
        if (
          metadata.areas.some((area) => area.workbookName === fromWorkbookName)
        ) {
          this.rangeMetadataManager.addRangeMetadata({
            metadata: metadata.metadata,
            areas: metadata.areas.map((area) =>
              area.workbookName === fromWorkbookName
                ? { ...area, workbookName: toWorkbookName }
                : area
            ),
          });
        }
      }

      this.workbookManager.updateFormulasForWorkbook(
        toWorkbookName,
        (formula) =>
          renameWorkbookInFormula({
            formula,
            oldWorkbookName: fromWorkbookName,
            newWorkbookName: toWorkbookName,
          })
      );

      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getWorkbookResourceKey(toWorkbookName)],
      });
    });
  }

  renameWorkbook(opts: { workbookName: string; newWorkbookName: string }) {
    return this.withUndoRedoCheckpoint(() => {
      const oldResourceKeys = this.getWorkbookResourceKeys(opts.workbookName);

      this.workbookManager.renameWorkbook(opts);
      this.namedExpressionManager.renameWorkbook(opts);
      this.tableManager.updateTablesForWorkbookRename(opts);
      this.styleManager.updateWorkbookName(
        opts.workbookName,
        opts.newWorkbookName
      );
      this.rangeMetadataManager.updateWorkbookName(
        opts.workbookName,
        opts.newWorkbookName
      );
      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameWorkbookInFormula({
          formula,
          oldWorkbookName: opts.workbookName,
          newWorkbookName: opts.newWorkbookName,
        })
      );
      this.referenceManager.updateWorkbookName(
        opts.workbookName,
        opts.newWorkbookName
      );

      this.emitMutation({
        touchedCells: buildFormulaTouchedCells(changedCells),
        resourceKeys: Array.from(
          new Set([
            ...oldResourceKeys,
            ...this.getWorkbookResourceKeys(opts.newWorkbookName),
          ])
        ),
      });
    });
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
    return this.withUndoRedoCheckpoint(() => {
      const normalizedContent = new Map(content);
      const preparedHeaderUpdates =
        this.tableManager.prepareHeaderUpdatesForSheet({
          ...opts,
          getCellContent: (address) =>
            normalizedContent.get(getCellReference(address)),
        });
      for (const generatedHeader of preparedHeaderUpdates.generatedHeaders) {
        normalizedContent.set(
          getCellReference(generatedHeader.address),
          generatedHeader.name
        );
      }

      const previousContent = this.getExistingSheetContent(opts);
      this.workbookManager.setSheetContent(opts, normalizedContent);
      this.tableManager.applyHeaderUpdates(preparedHeaderUpdates.updates);
      const renamedReferences = this.renameTableHeaderReferences(
        preparedHeaderUpdates.updates
      );
      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildSheetContentTouchedCells(
            opts,
            previousContent,
            normalizedContent
          ),
          buildFormulaTouchedCells(renamedReferences.changedFormulaCells)
        ),
        resourceKeys: Array.from(new Set(renamedReferences.resourceKeys)),
      });
    });
  }

  /**
   * Set the content of a single cell.
   */
  setCellContent(address: CellAddress, content: SerializedCellValue) {
    return this.withUndoRedoCheckpoint(() => {
      const preparedHeaderUpdate = this.tableManager.prepareHeaderUpdate(
        address,
        content
      );
      const previousValue = this.workbookManager.getCellContent(address);
      this.workbookManager.setCellContent(
        address,
        preparedHeaderUpdate.content
      );
      this.tableManager.applyHeaderUpdates(preparedHeaderUpdate.updates);
      const renamedReferences = this.renameTableHeaderReferences(
        preparedHeaderUpdate.updates
      );

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTouchedCells([
            {
              address,
              before: previousValue,
              after: preparedHeaderUpdate.content,
            },
          ]),
          buildFormulaTouchedCells(renamedReferences.changedFormulaCells)
        ),
        resourceKeys: Array.from(new Set(renamedReferences.resourceKeys)),
      });
    });
  }
  //#endregion

  //#region Auto-fill
  /**
   * Auto-fills one or more ranges based on the seedRange and the direction.
   * Supports pattern detection and style copying.
   */
  autoFill(
    opts: { sheetName: string; workbookName: string },
    /**
     * The user's original selection that defines the pattern/series.
     */
    seedRange: SpreadsheetRange,
    /**
     * One or more ranges to fill (the new cells populated by the drag, excluding the seed)
     */
    fillRanges: SpreadsheetRange[],
    /**
     * The direction of the fill.
     */
    direction: FillDirection
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      const touchedAddresses = this.dedupeAddresses(
        fillRanges.flatMap((range) =>
          getFiniteRangeAddresses({
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
            range,
          })
        )
      );
      const before = captureCellContents(this.workbookManager, touchedAddresses);

      this.autoFillManager.fill(opts, seedRange, fillRanges, direction);

      const after = captureCellContents(this.workbookManager, touchedAddresses);
      this.emitMutation({
        touchedCells: buildTouchedCells(
          touchedAddresses.map((address) => ({
            address,
            before: before.get(getMutationAddressKey(address)),
            after: after.get(getMutationAddressKey(address)),
          }))
        ),
        resourceKeys: [],
      });
    });
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   */
  clearSpreadsheetRange(address: RangeAddress) {
    return this.withUndoRedoCheckpoint(() => {
      const clearedCells = Array.from(
        this.workbookManager.iterateCellsInRange(address)
      );
      const before = captureCellContents(this.workbookManager, clearedCells);

      this.workbookManager.clearSpreadsheetRange(address);

      this.emitMutation({
        touchedCells: buildTouchedCells(
          clearedCells.map((cellAddress) => ({
            address: cellAddress,
            before: before.get(getMutationAddressKey(cellAddress)),
            after: undefined,
          }))
        ),
        resourceKeys: [],
      });
    });
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
      cellDataTypes: this.styleManager.getAllCellDataTypes(),
      rangeMetadata: this.rangeMetadataManager.getAllRangeMetadata(),
      references: this.referenceManager.getAllReferences(),
    };
  }

  onUpdate(listener: () => void) {
    return this.eventManager.onUpdate(listener);
  }

  private buildSerializedSnapshot(): EngineSnapshot {
    const evaluationSnapshots = this.dependencyManager.toSnapshot(
      this.evaluationManager
    );
    const historySnapshot = this.buildHistorySnapshot();

    return {
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        ...historySnapshot.managers,
        dependency: evaluationSnapshots.dependency,
        cache: evaluationSnapshots.cache,
      },
    };
  }

  private buildHistorySnapshot(): EngineHistorySnapshot {
    const workbookSnapshot = this.workbookManager.toSnapshot();

    return {
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        workbook: workbookSnapshot,
        namedExpression: this.buildNamedExpressionSnapshot(workbookSnapshot),
        table: this.tableManager.toSnapshot(),
        style: this.styleManager.toSnapshot(),
        rangeMetadata: this.rangeMetadataManager.toSnapshot(),
        reference: this.referenceManager.toSnapshot(),
      },
    };
  }

  private serializeHistorySnapshot(): string {
    return serialize(this.buildHistorySnapshot());
  }

  private buildNamedExpressionSnapshot(
    workbookSnapshot: EngineSnapshot["managers"]["workbook"]
  ): EngineSnapshot["managers"]["namedExpression"] {
    const namedExpressions = this.namedExpressionManager.toSnapshot();
    const workbookExpressions = new Map<string, Map<string, NamedExpression>>();
    const sheetExpressions = new Map<
      string,
      Map<string, Map<string, NamedExpression>>
    >();

    workbookSnapshot.forEach((workbook) => {
      const sourceWorkbookExpressions =
        namedExpressions.workbookExpressions.get(workbook.name);
      workbookExpressions.set(
        workbook.name,
        new Map(sourceWorkbookExpressions ?? [])
      );

      const sourceSheets = namedExpressions.sheetExpressions.get(
        workbook.name
      );
      const workbookSheetExpressions = new Map<
        string,
        Map<string, NamedExpression>
      >();
      workbook.sheets.forEach((_, sheetName) => {
        workbookSheetExpressions.set(
          sheetName,
          new Map(sourceSheets?.get(sheetName) ?? [])
        );
      });
      sheetExpressions.set(workbook.name, workbookSheetExpressions);
    });

    return {
      sheetExpressions,
      workbookExpressions,
      globalExpressions: new Map(namedExpressions.globalExpressions),
    };
  }

  serializeEngine(): string {
    return serialize(this.buildSerializedSnapshot());
  }

  private normalizeSupportedSnapshot(
    snapshot: unknown
  ): EngineHistorySnapshot | EngineSnapshot {
    const candidate = snapshot as {
      version?: unknown;
      managers?: Record<string, unknown> & {
        style?: Partial<StyleManagerSnapshot>;
      };
    };

    if (
      !candidate ||
      typeof candidate !== "object" ||
      (candidate.version !== ENGINE_SNAPSHOT_VERSION &&
        candidate.version !== LEGACY_ENGINE_SNAPSHOT_VERSION) ||
      !candidate.managers ||
      !candidate.managers.style
    ) {
      throw new Error(
        `Unsupported serialized engine format. Expected EngineSnapshot version ${LEGACY_ENGINE_SNAPSHOT_VERSION} or ${ENGINE_SNAPSHOT_VERSION}.`
      );
    }

    const style = candidate.managers.style;
    if (
      candidate.version === ENGINE_SNAPSHOT_VERSION &&
      !Array.isArray(style.cellDataTypes)
    ) {
      throw new Error(
        `Unsupported serialized engine format. Expected EngineSnapshot version ${LEGACY_ENGINE_SNAPSHOT_VERSION} or ${ENGINE_SNAPSHOT_VERSION}.`
      );
    }

    return {
      ...candidate,
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        ...candidate.managers,
        style: {
          ...style,
          cellDataTypes: style.cellDataTypes ?? [],
        },
      },
    } as unknown as EngineHistorySnapshot | EngineSnapshot;
  }

  private restoreDataManagersFromSnapshot(
    managers: EngineHistorySnapshot["managers"]
  ): void {
    this.namedExpressionManager.clear();
    this.workbookManager.restoreFromSnapshot(managers.workbook);

    managers.workbook.forEach((workbook) => {
      this.namedExpressionManager.addWorkbook(workbook.name);
      workbook.sheets.forEach((sheet) => {
        this.namedExpressionManager.addSheet({
          workbookName: workbook.name,
          sheetName: sheet.name,
        });
      });
    });

    this.namedExpressionManager.restoreFromSnapshot(managers.namedExpression);
    this.tableManager.restoreFromSnapshot(managers.table);
    this.styleManager.restoreFromSnapshot(managers.style);
    this.rangeMetadataManager.restoreFromSnapshot(managers.rangeMetadata);
    this.referenceManager.restoreFromSnapshot(managers.reference);
  }

  private restoreHistorySnapshot(data: string): void {
    const deserialized = this.normalizeSupportedSnapshot(deserialize(data));
    this.restoreDataManagersFromSnapshot(deserialized.managers);
    this.evaluationManager.clearEvaluationCache();
  }

  resetToSerializedEngine(data: string) {
    const deserialized = this.normalizeSupportedSnapshot(
      deserialize(data)
    ) as EngineSnapshot;
    this.restoreDataManagersFromSnapshot(deserialized.managers);
    this.dependencyManager.restoreFromSnapshot(
      {
        dependency: deserialized.managers.dependency,
        cache: deserialized.managers.cache,
      },
      this.evaluationManager
    );

    this.clearUndoRedoHistory();
    this.eventManager.emitUpdate();
  }
  //#endregion
}
