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
import { ReferenceManager } from "./managers/reference-manager";

/**
 * Main FormulaEngine class
 * @template TCellMetadata - Consumer-defined type for cell metadata (rich text, links, custom data, etc.)
 * @template TSheetMetadata - Consumer-defined type for sheet metadata (text boxes, frozen panes, etc.)
 * @template TWorkbookMetadata - Consumer-defined type for workbook metadata (themes, document properties, etc.)
 */
export class FormulaEngine<
  TCellMetadata = unknown,
  TSheetMetadata = unknown,
  TWorkbookMetadata = unknown
> {
  private workbookManager: WorkbookManager;
  private namedExpressionManager: NamedExpressionManager;
  private tableManager: TableManager;
  private eventManager: EventManager;
  private evaluationManager: EvaluationManager;
  private autoFillManager: AutoFill;
  private dependencyManager: DependencyManager;
  private styleManager: StyleManager;
  private copyManager: CopyManager;
  private referenceManager: ReferenceManager;

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
      this.namedExpressionManager
    );

    this.evaluationManager = new EvaluationManager(
      this.workbookManager,
      this.tableManager,
      formulaEvaluator,
      this.dependencyManager
    );

    this.styleManager = new StyleManager(this.evaluationManager);
    this.copyManager = new CopyManager(
      this.workbookManager,
      this.evaluationManager,
      this.styleManager
    );

    this.autoFillManager = new AutoFill(
      this.workbookManager,
      this.styleManager,
      this
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
  }

  /**
   * Static factory method to build an empty engine
   * @template TC - Consumer-defined cell metadata type
   * @template TS - Consumer-defined sheet metadata type
   * @template TW - Consumer-defined workbook metadata type
   */
  static buildEmpty<TC = unknown, TS = unknown, TW = unknown>(): FormulaEngine<
    TC,
    TS,
    TW
  > {
    return new FormulaEngine<TC, TS, TW>();
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
    metadata: TCellMetadata | undefined
  ): void {
    this.workbookManager.setCellMetadata(address, metadata);
    this.eventManager.emitUpdate();
  }

  /**
   * Get metadata for a cell
   */
  getCellMetadata(address: CellAddress): TCellMetadata | undefined {
    return this.workbookManager.getCellMetadata(address) as
      | TCellMetadata
      | undefined;
  }

  /**
   * Get all cell metadata for a sheet (serialized as Map)
   */
  getSheetMetadataSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, TCellMetadata> {
    return this.workbookManager.getSheetMetadataSerialized(opts) as Map<
      string,
      TCellMetadata
    >;
  }

  /**
   * Set metadata for a sheet
   * Sheet metadata can contain text boxes, frozen panes, print settings, or any consumer-defined data
   */
  setSheetMetadata(
    opts: { workbookName: string; sheetName: string },
    metadata: TSheetMetadata
  ): void {
    this.workbookManager.setSheetMetadata(opts, metadata);
    this.eventManager.emitUpdate();
  }

  /**
   * Get metadata for a sheet
   */
  getSheetMetadata(opts: {
    workbookName: string;
    sheetName: string;
  }): TSheetMetadata | undefined {
    return this.workbookManager.getSheetMetadata(opts) as
      | TSheetMetadata
      | undefined;
  }

  /**
   * Set metadata for a workbook
   * Workbook metadata can contain themes, document properties, settings, or any consumer-defined data
   */
  setWorkbookMetadata(workbookName: string, metadata: TWorkbookMetadata): void {
    this.workbookManager.setWorkbookMetadata(workbookName, metadata);
    this.eventManager.emitUpdate();
  }

  /**
   * Get metadata for a workbook
   */
  getWorkbookMetadata(workbookName: string): TWorkbookMetadata | undefined {
    return this.workbookManager.getWorkbookMetadata(workbookName) as
      | TWorkbookMetadata
      | undefined;
  }

  //#endregion

  //#region Reference Tracking
  /**
   * Create a tracked reference to a range
   * Returns a stable UUID that can be used to retrieve the address later
   * The reference automatically updates when sheets/workbooks are renamed
   */
  createRef(address: RangeAddress): string {
    return this.referenceManager.createRef(address);
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
    return this.referenceManager.deleteRef(refId);
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
    const removed = this.styleManager.removeConditionalStyle(
      workbookName,
      index
    );
    if (removed) {
      this.eventManager.emitUpdate();
    }
    return removed;
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
    this.styleManager.clearCellStyles(range);
    this.eventManager.emitUpdate();
  }

  //#endregion

  //#region Copy/Paste
  /**
   * Paste cells from source to target
   */
  pasteCells(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): void {
    this.copyManager.pasteCells(source, target, options);
    this.reevaluate();
    this.eventManager.emitUpdate();
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
    this.copyManager.fillAreas(seedRange, targetRanges, options);
    this.reevaluate();
    this.eventManager.emitUpdate();
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
      const shouldFill = pasteWidth > sourceWidth || pasteHeight > sourceHeight;

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
    this.pasteCells([source], target, {
      cut: true,
      type: "formula",
      include: "all",
    });
    this.reevaluate();
    this.eventManager.emitUpdate();
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
    const cells = this.copyManager.expandRangeToCells(sourceRange);
    this.pasteCells(cells, target, {
      cut: true,
      type: "formula",
      include: "all",
    });
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

    // Invalidate tracked references to this sheet
    this.referenceManager.invalidateSheet(opts.workbookName, opts.sheetName);

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

    // Update tracked references to this sheet
    this.referenceManager.updateSheetName(
      opts.workbookName,
      opts.sheetName,
      opts.newSheetName
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

    // Invalidate tracked references to this workbook
    this.referenceManager.invalidateWorkbook(workbookName);

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

      // Copy all cell metadata
      const targetSheet = this.workbookManager.getSheet({
        workbookName: toWorkbookName,
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
    const targetWorkbook = this.workbookManager
      .getWorkbooks()
      .get(toWorkbookName);
    if (targetWorkbook && sourceWorkbook.workbookMetadata !== undefined) {
      targetWorkbook.workbookMetadata = structuredClone(
        sourceWorkbook.workbookMetadata
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

    // Clone conditional styles
    const allConditionalStyles = this.styleManager.getAllConditionalStyles();
    for (const style of allConditionalStyles) {
      if (style.areas.some((area) => area.workbookName === fromWorkbookName)) {
        const newStyle: ConditionalStyle = {
          ...style,
          areas: style.areas.map((area) =>
            area.workbookName === fromWorkbookName
              ? { ...area, workbookName: toWorkbookName }
              : area
          ),
        };
        this.styleManager.addConditionalStyle(newStyle);
      }
    }

    // Clone cell styles
    const allCellStyles = this.styleManager.getAllCellStyles();
    for (const style of allCellStyles) {
      if (style.areas.some((area) => area.workbookName === fromWorkbookName)) {
        const newStyle: DirectCellStyle = {
          ...style,
          areas: style.areas.map((area) =>
            area.workbookName === fromWorkbookName
              ? { ...area, workbookName: toWorkbookName }
              : area
          ),
        };
        this.styleManager.addCellStyle(newStyle);
      }
    }

    // Update formulas in cloned workbook that reference the source workbook
    this.workbookManager.updateFormulasForWorkbook(toWorkbookName, (formula) =>
      renameWorkbookInFormula({
        formula,
        oldWorkbookName: fromWorkbookName,
        newWorkbookName: toWorkbookName,
      })
    );

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

    // Update tracked references to this workbook
    this.referenceManager.updateWorkbookName(
      opts.workbookName,
      opts.newWorkbookName
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
  ) {
    this.autoFillManager.fill(opts, seedRange, fillRanges, direction);
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
      references: this.referenceManager.getAllReferences(),
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
        conditionalStylesArray = Array.from(
          deserialized.conditionalStyles.values()
        ).flat();
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

    // Reset references if present
    if (deserialized.references) {
      this.referenceManager.resetReferences(deserialized.references);
    }

    // Re-evaluate all sheets to ensure all dependencies are resolved correctly
    this.reevaluate();
    this.eventManager.emitUpdate();
  }
  //#endregion
}
