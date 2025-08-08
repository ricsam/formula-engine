/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import type {
  SimpleCellAddress,
  SimpleCellRange,
  CellValue,
  RawCellContent,
  BoundingRect,
  ExportedChange,
  NamedExpression,
  SerializedNamedExpression,
  NamedExpressionOptions,
  CellType,
  CellValueType,
  CellValueDetailedType,
  FormatInfo,
  Sheet,
  Cell,
  FormulaEngineOptions,
  ArrayFormula,
  FormulaEngineEvents,
} from "./types";

import {
  addressToKey,
  getCellValueType,
  isFormulaError,
  isCellEmpty,
  parseDependencyKey,
} from "./types";

import {
  parseCellAddress,
  parseCellRange,
  addressToA1,
  rangeToA1,
  isValidAddress,
  iterateRange,
} from "./address";

import {
  createValueCell,
  createFormulaCell,
  createArrayCell,
  createEmptyCell,
  isFormula,
  extractFormula,
  parseInputValue,
  serializeCell,
  deserializeCell,
  hasFormula,
  isArrayCell,
  isEmptyCell,
  getCellDisplayValue,
} from "./cell";

import {
  createSheet,
  getCell,
  setCell,
  removeCell,
  clearSheet as clearSheetCells,
  getCellsInRange,
  getRangeValues as getSheetRangeValues,
  setRangeValues,
  getBoundingRect as getSheetBoundingRect,
  getAllCells,
  shiftCells,
  getFormulaCells,
  parseA1Key,
} from "./sheet";

import { parseFormula, type SheetResolver, ParseError } from "../parser/parser";
import { Evaluator, type EvaluationContext } from "../evaluator/evaluator";
import { DependencyGraph } from "../evaluator/dependency-graph";
import { ErrorHandler } from "../evaluator/error-handler";
import { functionRegistry } from "../functions";

// Simple event emitter for internal use
type EventListener<T = any> = (data: T) => void;

class EventEmitter<T extends Record<string, any>> {
  private listeners: { [K in keyof T]?: EventListener<T[K]>[] } = {};

  on<K extends keyof T>(event: K, listener: EventListener<T[K]>): () => void {
    if (!this.listeners[event]) {
      this.listeners[event] = [];
    }
    this.listeners[event]!.push(listener);

    // Return unsubscribe function
    return () => {
      const listeners = this.listeners[event];
      if (listeners) {
        const index = listeners.indexOf(listener);
        if (index > -1) {
          listeners.splice(index, 1);
        }
      }
    };
  }

  emit<K extends keyof T>(event: K, data: T[K]): void {
    const listeners = this.listeners[event];
    if (listeners) {
      listeners.forEach((listener) => listener(data));
    }
  }

  removeAllListeners(): void {
    this.listeners = {};
  }
}

/**
 * Main FormulaEngine class
 */
export class FormulaEngine {
  private sheets: Map<number, Sheet> = new Map();
  private namedExpressions: Map<string, NamedExpression> = new Map();
  private nextSheetId: number = 0;
  private evaluationSuspended: boolean = false;
  private pendingChanges: ExportedChange[] = [];
  private clipboard: {
    range: SimpleCellRange;
    data: RawCellContent[][];
    isValues?: boolean;
  } | null = null;
  private undoStack: Command[] = [];
  private redoStack: Command[] = [];
  private options: FormulaEngineOptions;
  private dependencyGraph: DependencyGraph;
  private evaluator: Evaluator;
  private errorHandler: ErrorHandler;
  private eventEmitter: EventEmitter<FormulaEngineEvents>;
  // New event system: cell-level immediate and sheet-level batched updates
  private cellUpdateListeners: Map<
    string,
    Set<
      (event: {
        address: SimpleCellAddress;
        oldValue: CellValue;
        newValue: CellValue;
      }) => void
    >
  > = new Map();
  private cellsUpdateListeners: Map<
    number,
    Set<
      (
        events: {
          address: SimpleCellAddress;
          oldValue: CellValue;
          newValue: CellValue;
        }[]
      ) => void
    >
  > = new Map();
  private pendingSheetUpdates: Map<
    number,
    { address: SimpleCellAddress; oldValue: CellValue; newValue: CellValue }[]
  > = new Map();
  private batchDepth: number = 0;
  private deferredEvaluations: Map<
    string,
    { address: SimpleCellAddress; formula: string; ast: any }
  > = new Map();

  constructor(options: FormulaEngineOptions = {}) {
    this.options = {
      evaluationMode: "eager",
      maxIterations: 100,
      cacheSize: 1000,
      enableArrayFormulas: true,
      enableNamedExpressions: true,
      locale: "en-US",
      ...options,
    };

    // Initialize evaluation components
    this.dependencyGraph = new DependencyGraph();
    this.errorHandler = new ErrorHandler();
    this.evaluator = new Evaluator(
      this.dependencyGraph,
      functionRegistry.getAllFunctions(),
      this.errorHandler
    );
    this.eventEmitter = new EventEmitter<FormulaEngineEvents>();
  }

  /**
   * Static factory method to build an empty engine
   */
  static buildEmpty(options?: FormulaEngineOptions): FormulaEngine {
    return new FormulaEngine(options);
  }

  // ===== Core Data Access =====

  getCellValue(cellAddress: SimpleCellAddress): CellValue {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return undefined;

    const cell = getCell(sheet, cellAddress);
    return cell?.value;
  }

  getCellFormula(cellAddress: SimpleCellAddress): string {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return "";

    const cell = getCell(sheet, cellAddress);
    if (!cell?.formula) return "";

    // Return formula with '=' prefix
    return "=" + cell.formula;
  }

  getCellSerialized(cellAddress: SimpleCellAddress): RawCellContent {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return undefined;

    const cell = getCell(sheet, cellAddress);
    return serializeCell(cell, cellAddress);
  }

  getSheetValues(sheetId: number): Map<string, CellValue> {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return new Map();

    const result = new Map<string, CellValue>();
    for (const [key, cell] of sheet.cells) {
      result.set(key, cell.value);
    }
    return result;
  }

  getSheetFormulas(sheetId: number): Map<string, string> {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return new Map();

    const result = new Map<string, string>();
    for (const [key, cell] of sheet.cells) {
      if (cell.formula) {
        result.set(key, cell.formula);
      }
    }
    return result;
  }

  getSheetSerialized(sheetId: number): Map<string, RawCellContent> {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return new Map();

    const result = new Map<string, RawCellContent>();
    for (const [key, cell] of sheet.cells) {
      const address = parseCellAddress(key, sheetId);
      const serialized = serializeCell(cell, address || undefined);
      // Only include cells that have content when serialized
      // Spilled array cells return undefined and should not be included
      if (serialized !== undefined) {
        result.set(key, serialized);
      }
    }
    return result;
  }

  getSheetBoundingRect(sheetId: number): BoundingRect | undefined {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return undefined;

    return getSheetBoundingRect(sheet);
  }

  getAllSheetsBoundingRects(): Record<string, BoundingRect> {
    const result: Record<string, BoundingRect> = {};

    for (const [id, sheet] of this.sheets) {
      const rect = getSheetBoundingRect(sheet);
      if (rect) {
        result[id.toString()] = rect;
      }
    }

    return result;
  }

  getAllSheetsValues(): Record<string, Map<string, CellValue>> {
    const result: Record<string, Map<string, CellValue>> = {};

    for (const [id, sheet] of this.sheets) {
      result[id.toString()] = this.getSheetValues(id);
    }

    return result;
  }

  getAllSheetsFormulas(): Record<string, Map<string, string>> {
    const result: Record<string, Map<string, string>> = {};

    for (const [id, sheet] of this.sheets) {
      result[id.toString()] = this.getSheetFormulas(id);
    }

    return result;
  }

  getAllSheetsSerialized(): Record<string, Map<string, RawCellContent>> {
    const result: Record<string, Map<string, RawCellContent>> = {};

    for (const [id, sheet] of this.sheets) {
      result[id.toString()] = this.getSheetSerialized(id);
    }

    return result;
  }

  // ===== Data Manipulation =====

  public setCellContent(
    topLeftCornerAddress: SimpleCellAddress,
    cellContents: RawCellContent[][] | RawCellContent
  ): ExportedChange[] {
    // Don't defer evaluation for setCellContent - only for setSheetContent batch operations
    const changes: ExportedChange[] = [];

    if (Array.isArray(cellContents)) {
      // Handle 2D array
      for (let row = 0; row < cellContents.length; row++) {
        const rowData = cellContents[row];
        if (rowData) {
          for (let col = 0; col < rowData.length; col++) {
            const address: SimpleCellAddress = {
              sheet: topLeftCornerAddress.sheet,
              col: topLeftCornerAddress.col + col,
              row: topLeftCornerAddress.row + row,
            };

            const change = this.setCellValue(address, rowData[col]);
            if (change) {
              changes.push(change);
            }
          }
        }
      }
    } else {
      // Handle single value
      const change = this.setCellValue(topLeftCornerAddress, cellContents);
      if (change) {
        changes.push(change);
      }
    }

    return this.evaluationSuspended ? [] : changes;
  }

  public setSheetContent(
    sheetId: number,
    contents: Map<string, RawCellContent>
  ): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];

    this.beginBatch();
    const changes: ExportedChange[] = [];
    const processedKeys = new Set<string>();

    // Update or add cells
    for (const [key, content] of contents) {
      const address = parseCellAddress(key, sheetId);
      if (address) {
        const change = this.setCellValue(address, content);
        if (change) {
          changes.push(change);
        }
        processedKeys.add(key);
      }
    }

    // Remove cells not in the new contents, but preserve spilled array cells
    const keysToRemove: string[] = [];
    for (const key of sheet.cells.keys()) {
      if (!processedKeys.has(key)) {
        const address = parseCellAddress(key, sheetId);
        if (address) {
          const cell = getCell(sheet, address);
          // Don't remove spilled array cells (they are auto-generated)
          if (cell && cell.type === "ARRAY" && cell.arrayFormula) {
            const isOrigin =
              cell.arrayFormula.originAddress.sheet === address.sheet &&
              cell.arrayFormula.originAddress.row === address.row &&
              cell.arrayFormula.originAddress.col === address.col;

            if (!isOrigin) {
              // This is a spilled cell, don't remove it
              continue;
            }
          }
        }
        keysToRemove.push(key);
      }
    }

    for (const key of keysToRemove) {
      const address = parseCellAddress(key, sheetId);
      if (address) {
        const oldCell = getCell(sheet, address);
        if (oldCell) {
          const change = this.setCellValue(address, undefined);
          if (change) {
            changes.push(change);
          }
        }
      }
    }

    this.endBatch();
    return this.evaluationSuspended ? [] : changes;
  }

  private setCellValue(
    address: SimpleCellAddress,
    content: RawCellContent
  ): ExportedChange | null {
    const sheet = this.sheets.get(address.sheet);
    if (!sheet) return null;

    const oldCell = getCell(sheet, address);
    const oldValue = oldCell?.value;

    // Parse the content and create appropriate cell
    let newCell: Cell;

    if (content === undefined || content === null || content === "") {
      // Empty cell
      removeCell(sheet, address);

      // Clear dependencies in the graph
      const addressKey = addressToKey(address);
      this.dependencyGraph.clearDependencies(addressKey);
      this.dependencyGraph.removeNode(addressKey);

      if (oldValue !== undefined) {
        // Notify listeners about single cell update and collect for sheet batch
        this.emitCellUpdate({ address, oldValue, newValue: undefined });
        return {
          address,
          oldValue,
          newValue: undefined,
          type: "cell-change",
        };
      }
      return null;
    } else if (typeof content === "string" && isFormula(content)) {
      // Formula cell
      const formula = extractFormula(content);
      const addressKey = addressToKey(address);
      newCell = createFormulaCell(formula);

      // Parse the formula and handle evaluation
      try {
        const ast = parseFormula(formula, address.sheet, (sheetName) =>
          this.getSheetId(sheetName)
        );

        // Setup cell and dependencies
        this.dependencyGraph.addCell(address);
        this.dependencyGraph.clearDependencies(addressKey);

        // Defer evaluation if in batch mode, otherwise evaluate immediately
        if (this.batchDepth > 0) {
          // Store for deferred evaluation
          this.deferredEvaluations.set(addressKey, { address, formula, ast });

          // Set initial value to undefined for now
          newCell.value = undefined;
        } else {
          // Evaluate immediately (non-batch mode)
          // Create evaluation context
          const context: EvaluationContext = {
            currentSheet: address.sheet,
            currentCell: address,
            namedExpressions: this.namedExpressions,
            getCellValue: (addr: SimpleCellAddress) =>
              this.getCellValueInternal(addr, context.evaluationStack),
            getRangeValues: (range: SimpleCellRange) =>
              this.getRangeValuesInternal(range, context.evaluationStack),
            getFunction: (name: string) => functionRegistry.get(name),
            errorHandler: this.errorHandler,
            evaluationStack: new Set<string>(),
            sheetResolver: (sheetName: string) => this.getSheetId(sheetName),
          };

          // Add current cell to evaluation stack to detect cycles
          context.evaluationStack.add(addressKey);

          // Evaluate the formula
          const result = this.evaluator.evaluate(ast, context);

          // Register dependencies
          for (const dep of result.dependencies) {
            try {
              const parsed = parseDependencyKey(dep);

              switch (parsed.type) {
                case "cell":
                  this.dependencyGraph.addCell(parsed.address);
                  this.dependencyGraph.addDependency(addressKey, dep);
                  break;

                case "range":
                  this.dependencyGraph.addRange(parsed.range);
                  this.dependencyGraph.addDependency(addressKey, dep);
                  break;

                case "named":
                  this.dependencyGraph.addNamedExpression(
                    parsed.name,
                    parsed.scope
                  );
                  this.dependencyGraph.addDependency(addressKey, dep);
                  break;
              }
            } catch (error) {
              // Log the error but still add the dependency to avoid breaking evaluation
              console.warn(`Failed to parse dependency key: ${dep}`, error);
              this.dependencyGraph.addDependency(addressKey, dep);
            }
          }

          // Handle array results
          if (result.type === "2d-array") {
            // Implement array spilling
            const arrayValue = result.value;
            const { rows, cols } = result.dimensions;

            // Calculate spill range
            const spillRange: SimpleCellRange = {
              start: address,
              end: {
                sheet: address.sheet,
                row: address.row + rows - 1,
                col: address.col + cols - 1,
              },
            };

            // Check if spill range is available (no non-empty cells)
            if (this.canSpillArray(spillRange, address)) {
              // Clear any previous spill range for this formula
              this.clearPreviousSpillRange(address);

              // Create array formula info
              const arrayFormula: ArrayFormula = {
                formula,
                originAddress: address,
                spillRange,
              };

              // Set the origin cell
              newCell = createArrayCell(
                formula,
                arrayValue[0]?.[0] ?? undefined,
                arrayFormula
              );
              newCell.dependencies = new Set();

              // Spill the array values
              this.spillArray(arrayValue, spillRange, arrayFormula);
            } else {
              // Spill blocked - return #SPILL! error
              newCell.value = "#SPILL!";
            }
          } else {
            newCell.value = result.value;
          }
        } // End of else clause for immediate evaluation
      } catch (error) {
        // If parsing or evaluation fails, store the error
        if (
          error instanceof ParseError &&
          error.message.includes("not found")
        ) {
          // Sheet not found error
          newCell.value = "#REF!";
        } else if (error instanceof Error && error.message.startsWith("#")) {
          newCell.value = error.message as any;
        } else {
          newCell.value = "#ERROR!";
        }
      }
    } else {
      // Value cell
      const value = parseInputValue(content);
      newCell = createValueCell(value);
    }

    setCell(sheet, address, newCell);

    if (oldValue !== newCell.value) {
      const change: ExportedChange = {
        address,
        oldValue,
        newValue: newCell.value,
        type: "cell-change",
      };

      // Notify listeners for the primary changed cell
      this.emitCellUpdate({
        address,
        oldValue: oldValue ?? undefined,
        newValue: newCell.value,
      });

      // Then recalculate dependent cells (which will emit their own events)
      if (!this.evaluationSuspended) {
        this.recalculateDependents(address);
      }

      if (this.evaluationSuspended) {
        this.pendingChanges.push(change);
        return null;
      }

      return change;
    }

    return null;
  }

  getSheetContents(sheetId: number): Map<string, CellValue> {
    return this.getSheetValues(sheetId);
  }

  getRangeValues(source: SimpleCellRange): CellValue[][] {
    // Use the internal method that handles infinite ranges
    return this.getRangeValuesInternal(source);
  }

  getRangeFormulas(source: SimpleCellRange): (string | undefined)[][] {
    const sheet = this.sheets.get(source.start.sheet);
    if (!sheet) return [];

    const width = source.end.col - source.start.col + 1;
    const height = source.end.row - source.start.row + 1;
    const result: (string | undefined)[][] = [];

    for (let row = 0; row < height; row++) {
      const rowData: (string | undefined)[] = [];
      for (let col = 0; col < width; col++) {
        const address: SimpleCellAddress = {
          sheet: source.start.sheet,
          col: source.start.col + col,
          row: source.start.row + row,
        };
        const cell = getCell(sheet, address);
        rowData.push(cell?.formula);
      }
      result.push(rowData);
    }

    return result;
  }

  getRangeSerialized(source: SimpleCellRange): RawCellContent[][] {
    const sheet = this.sheets.get(source.start.sheet);
    if (!sheet) return [];

    const width = source.end.col - source.start.col + 1;
    const height = source.end.row - source.start.row + 1;
    const result: RawCellContent[][] = [];

    for (let row = 0; row < height; row++) {
      const rowData: RawCellContent[] = [];
      for (let col = 0; col < width; col++) {
        const address: SimpleCellAddress = {
          sheet: source.start.sheet,
          col: source.start.col + col,
          row: source.start.row + row,
        };
        const cell = getCell(sheet, address);
        rowData.push(serializeCell(cell, address));
      }
      result.push(rowData);
    }

    return result;
  }

  getFillRangeData(
    source: SimpleCellRange,
    target: SimpleCellRange,
    offsetsFromTarget: boolean
  ): RawCellContent[][] {
    // TODO: Implement smart fill logic (patterns, series, etc.)
    // For now, just copy the source data
    return this.getRangeSerialized(source);
  }

  // ===== Sheet Management =====

  isItPossibleToAddSheet(sheetName: string): boolean {
    if (!sheetName || typeof sheetName !== "string") return false;

    // Check if name already exists
    for (const sheet of this.sheets.values()) {
      if (sheet.name === sheetName) return false;
    }

    return true;
  }

  addSheet(sheetName?: string): string {
    const id = this.nextSheetId++;
    const name = sheetName || `Sheet${id + 1}`;

    // Ensure unique name
    let finalName = name;
    let suffix = 1;
    while (!this.isItPossibleToAddSheet(finalName)) {
      finalName = `${name}_${suffix}`;
      suffix++;
    }

    const sheet = createSheet(id, finalName);
    this.sheets.set(id, sheet);

    // Emit sheet-added event
    this.emit("sheet-added", {
      sheetId: id,
      sheetName: finalName,
    });

    return finalName;
  }

  isItPossibleToRemoveSheet(sheetId: number): boolean {
    // Must have at least one sheet
    return this.sheets.has(sheetId) && this.sheets.size > 1;
  }

  removeSheet(sheetId: number): ExportedChange[] {
    if (!this.isItPossibleToRemoveSheet(sheetId)) return [];

    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];

    const deletedSheetName = sheet.name;

    // Collect all cell changes from the deleted sheet
    const changes: ExportedChange[] = [];
    for (const [key, cell] of sheet.cells) {
      const address = parseCellAddress(key, sheetId);
      if (address && cell.value !== undefined) {
        changes.push({
          address,
          oldValue: cell.value,
          newValue: undefined,
          type: "cell-change",
        });
      }
    }

    // Delete the sheet
    this.sheets.delete(sheetId);

    // Force re-evaluation of all formulas in remaining sheets that reference the deleted sheet
    for (const [remainingSheetId, remainingSheet] of this.sheets) {
      const formulaCells = getFormulaCells(remainingSheet);

      for (const [key, cell] of formulaCells) {
        if (cell.formula) {
          // Check if formula references the deleted sheet
          // Simple check - look for sheet name followed by !
          const sheetRefPattern = new RegExp(`\\b${deletedSheetName}!`, "i");
          const quotedSheetRefPattern = new RegExp(
            `'${deletedSheetName.replace(/'/g, "''")}!'`,
            "i"
          );

          if (
            sheetRefPattern.test(cell.formula) ||
            quotedSheetRefPattern.test(cell.formula)
          ) {
            const address = parseCellAddress(key, remainingSheetId);
            if (address) {
              const oldValue = cell.value;

              // Force recalculation of this cell
              const cellKey = addressToKey(address);
              this.recalculateCell(cellKey);

              // The cell should now have #REF! error
              const newCell = getCell(remainingSheet, address);
              if (newCell) {
                changes.push({
                  address,
                  oldValue,
                  newValue: newCell.value,
                  type: "cell-change",
                });
              }
            }
          }
        }
      }
    }

    // Emit sheet-removed event
    this.emit("sheet-removed", {
      sheetId,
      sheetName: deletedSheetName,
    });

    return changes;
  }

  isItPossibleToClearSheet(sheetId: number): boolean {
    return this.sheets.has(sheetId);
  }

  clearSheet(sheetId: number): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];

    const changes: ExportedChange[] = [];

    // Collect all cell changes
    for (const [key, cell] of sheet.cells) {
      const address = parseCellAddress(key, sheetId);
      if (address && cell.value !== undefined) {
        changes.push({
          address,
          oldValue: cell.value,
          newValue: undefined,
          type: "cell-change",
        });
      }
    }

    clearSheetCells(sheet);

    return changes;
  }

  isItPossibleToReplaceSheetContent(
    sheetId: number,
    values: RawCellContent[][]
  ): boolean {
    return this.sheets.has(sheetId);
  }

  getSheetName(sheetId: number): string {
    const sheet = this.sheets.get(sheetId);
    return sheet ? sheet.name : "";
  }

  getSheetNames(): string[] {
    return Array.from(this.sheets.values()).map((sheet) => sheet.name);
  }

  getSheetId(sheetName: string): number {
    for (const [id, sheet] of this.sheets) {
      if (sheet.name === sheetName) {
        return id;
      }
    }
    return -1;
  }

  doesSheetExist(sheetName: string): boolean {
    return this.getSheetId(sheetName) !== -1;
  }

  countSheets(): number {
    return this.sheets.size;
  }

  isItPossibleToRenameSheet(sheetId: number, newName: string): boolean {
    if (!this.sheets.has(sheetId)) return false;
    if (!newName || typeof newName !== "string") return false;

    // Check if name already exists (except for current sheet)
    for (const [id, sheet] of this.sheets) {
      if (id !== sheetId && sheet.name === newName) {
        return false;
      }
    }

    return true;
  }

  renameSheet(sheetId: number, newName: string): void {
    if (!this.isItPossibleToRenameSheet(sheetId, newName)) return;

    const sheet = this.sheets.get(sheetId);
    if (sheet) {
      const oldName = sheet.name;
      sheet.name = newName;

      // Emit sheet-renamed event
      this.emit("sheet-renamed", {
        sheetId,
        oldName,
        newName,
      });
    }
  }

  // ===== Operations =====

  removeRows(sheetId: number, ...indexes: number[]): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet || indexes.length === 0) return [];

    const changes: ExportedChange[] = [];
    const sortedIndexes = [...new Set(indexes)].sort((a, b) => b - a); // Remove duplicates and sort descending

    for (const rowIndex of sortedIndexes) {
      // Collect cells in the row to be removed
      for (const [key, cell] of sheet.cells) {
        const address = parseCellAddress(key, sheetId);
        if (address && address.row === rowIndex && cell.value !== undefined) {
          changes.push({
            address,
            oldValue: cell.value,
            newValue: undefined,
            type: "cell-change",
          });
        }
      }

      // Shift cells down
      shiftCells(sheet, "row", rowIndex + 1, -1);
    }

    return changes;
  }

  removeColumns(sheetId: number, ...indexes: number[]): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet || indexes.length === 0) return [];

    const changes: ExportedChange[] = [];
    const sortedIndexes = [...new Set(indexes)].sort((a, b) => b - a);

    for (const colIndex of sortedIndexes) {
      // Collect cells in the column to be removed
      for (const [key, cell] of sheet.cells) {
        const address = parseCellAddress(key, sheetId);
        if (address && address.col === colIndex && cell.value !== undefined) {
          changes.push({
            address,
            oldValue: cell.value,
            newValue: undefined,
            type: "cell-change",
          });
        }
      }

      // Shift cells left
      shiftCells(sheet, "col", colIndex + 1, -1);
    }

    return changes;
  }

  isItPossibleToMoveCells(
    source: SimpleCellRange,
    destinationLeftCorner: SimpleCellAddress
  ): boolean {
    // Check if sheets match
    if (source.start.sheet !== destinationLeftCorner.sheet) return false;

    // Check if destination is valid
    return isValidAddress(destinationLeftCorner);
  }

  moveCells(
    source: SimpleCellRange,
    destinationLeftCorner: SimpleCellAddress
  ): ExportedChange[] {
    if (!this.isItPossibleToMoveCells(source, destinationLeftCorner)) return [];

    const sheet = this.sheets.get(source.start.sheet);
    if (!sheet) return [];

    const changes: ExportedChange[] = [];

    // Cut operation
    const cutData = this.cut(source);

    // Paste operation
    const pasteChanges = this.paste(destinationLeftCorner);
    changes.push(...pasteChanges);

    return changes;
  }

  moveRows(
    sheetId: number,
    startRow: number,
    numberOfRows: number,
    targetRow: number
  ): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];

    const source: SimpleCellRange = {
      start: { sheet: sheetId, col: 0, row: startRow },
      end: { sheet: sheetId, col: 16383, row: startRow + numberOfRows - 1 }, // Excel max columns
    };

    const destination: SimpleCellAddress = {
      sheet: sheetId,
      col: 0,
      row: targetRow,
    };

    return this.moveCells(source, destination);
  }

  isItPossibleToMoveColumns(
    sheetId: number,
    startColumn: number,
    numberOfColumns: number,
    targetColumn: number
  ): boolean {
    return (
      this.sheets.has(sheetId) &&
      startColumn >= 0 &&
      numberOfColumns > 0 &&
      targetColumn >= 0
    );
  }

  moveColumns(
    sheetId: number,
    startColumn: number,
    numberOfColumns: number,
    targetColumn: number
  ): ExportedChange[] {
    if (
      !this.isItPossibleToMoveColumns(
        sheetId,
        startColumn,
        numberOfColumns,
        targetColumn
      )
    ) {
      return [];
    }

    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];

    const source: SimpleCellRange = {
      start: { sheet: sheetId, col: startColumn, row: 0 },
      end: {
        sheet: sheetId,
        col: startColumn + numberOfColumns - 1,
        row: 1048575,
      }, // Excel max rows
    };

    const destination: SimpleCellAddress = {
      sheet: sheetId,
      col: targetColumn,
      row: 0,
    };

    return this.moveCells(source, destination);
  }

  copy(source: SimpleCellRange): CellValue[][] {
    const values = this.getRangeValues(source);
    const serializedData = this.getRangeSerialized(source);

    this.clipboard = {
      range: source,
      data: serializedData,
      isValues: false,
    };

    return values;
  }

  cut(source: SimpleCellRange): CellValue[][] {
    const values = this.copy(source);

    // Clear the source range
    const sheet = this.sheets.get(source.start.sheet);
    if (sheet) {
      for (const address of iterateRange(source)) {
        removeCell(sheet, address);
      }
    }

    return values;
  }

  paste(targetLeftCorner: SimpleCellAddress): ExportedChange[] {
    if (!this.clipboard) return [];

    const changes: ExportedChange[] = [];
    const { range, data, isValues } = this.clipboard;

    for (let row = 0; row < data.length; row++) {
      const rowData = data[row];
      if (rowData) {
        for (let col = 0; col < rowData.length; col++) {
          const sourceAddress: SimpleCellAddress = {
            sheet: range.start.sheet,
            col: range.start.col + col,
            row: range.start.row + row,
          };

          const targetAddress: SimpleCellAddress = {
            sheet: targetLeftCorner.sheet,
            col: targetLeftCorner.col + col,
            row: targetLeftCorner.row + row,
          };

          let content = rowData[col];

          // If copying as values, get the actual value instead of formula
          if (
            isValues &&
            typeof content === "string" &&
            content.startsWith("=")
          ) {
            // Get the value from the source
            const sourceSheet = this.sheets.get(sourceAddress.sheet);
            if (sourceSheet) {
              const sourceCell = getCell(sourceSheet, sourceAddress);
              content = sourceCell?.value ?? undefined;
            }
          } else if (typeof content === "string" && content.startsWith("=")) {
            // If content is a formula and not copying as values, adjust relative references
            content = this.adjustFormula(content, sourceAddress, targetAddress);
          }

          const change = this.setCellValue(targetAddress, content);
          if (change) {
            changes.push(change);
          }
        }
      }
    }

    return changes;
  }

  isClipboardEmpty(): boolean {
    return this.clipboard === null;
  }

  clearClipboard(): void {
    this.clipboard = null;
  }

  // ===== Address Utilities =====

  simpleCellAddressFromString(
    cellAddress: string,
    contextSheetId: number
  ): SimpleCellAddress {
    const parsed = parseCellAddress(cellAddress, contextSheetId);
    if (!parsed) {
      throw new Error(`Invalid cell address: ${cellAddress}`);
    }
    return parsed;
  }

  simpleCellRangeFromString(
    cellRange: string,
    contextSheetId: number
  ): SimpleCellRange {
    const parsed = parseCellRange(cellRange, contextSheetId);
    if (!parsed) {
      throw new Error(`Invalid cell range: ${cellRange}`);
    }
    return parsed;
  }

  simpleCellRangeToString(
    cellRange: SimpleCellRange,
    optionsOrContextSheetId: { includeSheetName?: boolean } | number
  ): string {
    if (typeof optionsOrContextSheetId === "number") {
      // Legacy API - context sheet ID
      return rangeToA1(cellRange);
    }

    const sheet = this.sheets.get(cellRange.start.sheet);
    const sheetName = sheet?.name;

    return rangeToA1(cellRange, {
      includeSheetName: optionsOrContextSheetId.includeSheetName,
      sheetName,
    });
  }

  // ===== Dependency Analysis =====

  getCellDependents(
    address: SimpleCellAddress | SimpleCellRange
  ): (SimpleCellRange | SimpleCellAddress)[] {
    const result: (SimpleCellRange | SimpleCellAddress)[] = [];

    if ("start" in address && "end" in address) {
      // Range: Get dependents for each cell in the range
      const seenKeys = new Set<string>();

      for (const cellAddr of iterateRange(address)) {
        const key = addressToKey(cellAddr);
        const dependents = this.dependencyGraph.getDependents(key);

        for (const dep of dependents) {
          if (!seenKeys.has(dep)) {
            seenKeys.add(dep);
            const parsed = this.parseNodeKey(dep);
            if (parsed) {
              result.push(parsed);
            }
          }
        }
      }
    } else {
      // Single cell
      const key = addressToKey(address);
      const dependents = this.dependencyGraph.getDependents(key);

      for (const dep of dependents) {
        const parsed = this.parseNodeKey(dep);
        if (parsed) {
          result.push(parsed);
        }
      }
    }

    return result;
  }

  getCellPrecedents(
    address: SimpleCellAddress | SimpleCellRange
  ): (SimpleCellRange | SimpleCellAddress)[] {
    const result: (SimpleCellRange | SimpleCellAddress)[] = [];

    if ("start" in address && "end" in address) {
      // Range: Get precedents for each cell in the range
      const seenKeys = new Set<string>();

      for (const cellAddr of iterateRange(address)) {
        const key = addressToKey(cellAddr);
        // Use transitive precedents to include dependencies through named expressions
        const precedents = this.dependencyGraph.getTransitivePrecedents(key);

        for (const prec of precedents) {
          if (!seenKeys.has(prec)) {
            seenKeys.add(prec);
            const parsed = this.parseNodeKey(prec);
            if (parsed) {
              result.push(parsed);
            }
          }
        }
      }
    } else {
      // Single cell
      const key = addressToKey(address);
      // Use transitive precedents to include dependencies through named expressions
      const precedents = this.dependencyGraph.getTransitivePrecedents(key);

      // Group individual cells into ranges where possible
      const cellPrecedents: SimpleCellAddress[] = [];
      const rangePrecedents: SimpleCellRange[] = [];

      for (const prec of precedents) {
        const parsed = this.parseNodeKey(prec);
        if (parsed) {
          if ("start" in parsed && "end" in parsed) {
            rangePrecedents.push(parsed);
          } else {
            cellPrecedents.push(parsed);
          }
        }
      }

      // For now, prioritize range dependencies over individual cells
      // In a full implementation, we'd check if individual cells are covered by ranges
      result.push(...rangePrecedents);

      // Only add individual cells that aren't covered by ranges
      for (const cell of cellPrecedents) {
        const isCoveredByRange = rangePrecedents.some((range) =>
          this.isAddressInRange(cell, range)
        );
        if (!isCoveredByRange) {
          result.push(cell);
        }
      }
    }

    return result;
  }

  private parseNodeKey(
    key: string
  ): SimpleCellAddress | SimpleCellRange | null {
    const parts = key.split(":");

    if (parts.length === 3 && parts[0] && parts[1] && parts[2]) {
      // Cell: sheet:col:row
      return {
        sheet: parseInt(parts[0]),
        col: parseInt(parts[1]),
        row: parseInt(parts[2]),
      };
    } else if (
      parts.length === 5 &&
      parts[0] &&
      parts[1] &&
      parts[2] &&
      parts[3] &&
      parts[4]
    ) {
      // Range: sheet:startCol:startRow:endCol:endRow
      return {
        start: {
          sheet: parseInt(parts[0]),
          col: parseInt(parts[1]),
          row: parseInt(parts[2]),
        },
        end: {
          sheet: parseInt(parts[0]),
          col: parseInt(parts[3]),
          row: parseInt(parts[4]),
        },
      };
    } else if (key.startsWith("name:")) {
      // Named expression: name:name or name:scope:name
      // For now, we don't have a way to represent named expressions in our return type
      // This is a limitation of the current API design
      // In a full implementation, we'd extend the return type
      return null;
    }

    return null;
  }

  private isAddressInRange(
    address: SimpleCellAddress,
    range: SimpleCellRange
  ): boolean {
    return (
      address.sheet === range.start.sheet &&
      address.col >= range.start.col &&
      address.col <= range.end.col &&
      address.row >= range.start.row &&
      address.row <= range.end.row
    );
  }

  // ===== Cell Information =====

  getCellType(cellAddress: SimpleCellAddress): CellType {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return "EMPTY";

    const cell = getCell(sheet, cellAddress);
    return cell?.type || "EMPTY";
  }

  doesCellHaveSimpleValue(cellAddress: SimpleCellAddress): boolean {
    return this.getCellType(cellAddress) === "VALUE";
  }

  doesCellHaveFormula(cellAddress: SimpleCellAddress): boolean {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return false;

    const cell = getCell(sheet, cellAddress);
    return hasFormula(cell);
  }

  isCellEmpty(cellAddress: SimpleCellAddress): boolean {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return true;

    const cell = getCell(sheet, cellAddress);
    return isEmptyCell(cell);
  }

  isCellPartOfArray(cellAddress: SimpleCellAddress): boolean {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return false;

    const cell = getCell(sheet, cellAddress);
    return isArrayCell(cell);
  }

  getCellValueType(cellAddress: SimpleCellAddress): CellValueType {
    const value = this.getCellValue(cellAddress);
    return getCellValueType(value);
  }

  getCellValueDetailedType(
    cellAddress: SimpleCellAddress
  ): CellValueDetailedType {
    // For now, detailed type is the same as regular type
    return this.getCellValueType(cellAddress);
  }

  getCellValueFormat(cellAddress: SimpleCellAddress): FormatInfo {
    // Placeholder - formatting not yet implemented
    return {};
  }

  // ===== Evaluation Control =====

  suspendEvaluation(): void {
    this.evaluationSuspended = true;
  }

  resumeEvaluation(): ExportedChange[] {
    this.evaluationSuspended = false;
    const changes = [...this.pendingChanges];
    this.pendingChanges = [];

    // TODO: Trigger recalculation for pending changes

    return changes;
  }

  isEvaluationSuspended(): boolean {
    return this.evaluationSuspended;
  }

  // ===== Named Expressions =====

  isItPossibleToAddNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number
  ): boolean {
    if (!this.options.enableNamedExpressions) return false;
    if (!expressionName || typeof expressionName !== "string") return false;

    const key = this.getNamedExpressionKey(expressionName, scope);
    return !this.namedExpressions.has(key);
  }

  addNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: NamedExpressionOptions
  ): ExportedChange[] {
    if (
      !this.isItPossibleToAddNamedExpression(expressionName, expression, scope)
    ) {
      if (options?.overwrite) {
        return this.changeNamedExpression(
          expressionName,
          expression,
          scope,
          options
        );
      }
      return [];
    }

    const key = this.getNamedExpressionKey(expressionName, scope);
    const namedExpr: NamedExpression = {
      name: expressionName,
      expression: String(expression),
      scope,
    };

    this.namedExpressions.set(key, namedExpr);

    // Add the named expression to the dependency graph
    const namedExprKey = DependencyGraph.getNamedExpressionKey(
      expressionName,
      scope
    );
    this.dependencyGraph.addNamedExpression(expressionName, scope);

    // If the expression is a formula, parse it and build dependencies
    if (typeof expression === "string" && expression.startsWith("=")) {
      try {
        const formula = expression.substring(1);
        const ast = parseFormula(formula, scope || 0, (sheetName) =>
          this.getSheetId(sheetName)
        );

        const context: EvaluationContext = {
          currentSheet: scope || 0,
          namedExpressions: this.namedExpressions,
          getCellValue: (addr) => this.getCellValue(addr),
          getRangeValues: (range) => this.getRangeValues(range),
          getFunction: (name: string) => functionRegistry.get(name),
          errorHandler: this.errorHandler,
          evaluationStack: new Set<string>(),
          sheetResolver: (sheetName: string) => this.getSheetId(sheetName),
        };

        const result = this.evaluator.evaluate(ast, context);

        // Register dependencies for the named expression
        for (const dep of result.dependencies) {
          // Parse the dependency and add to graph if needed
          const parts = dep.split(":");

          if (parts.length === 3 && parts[0] && parts[1] && parts[2]) {
            // Cell dependency
            const depAddress: SimpleCellAddress = {
              sheet: parseInt(parts[0]),
              col: parseInt(parts[1]),
              row: parseInt(parts[2]),
            };
            this.dependencyGraph.addCell(depAddress);
            this.dependencyGraph.addDependency(namedExprKey, dep);
          } else if (
            parts.length === 5 &&
            parts[0] &&
            parts[1] &&
            parts[2] &&
            parts[3] &&
            parts[4]
          ) {
            // Range dependency
            const depRange: SimpleCellRange = {
              start: {
                sheet: parseInt(parts[0]),
                col: parseInt(parts[1]),
                row: parseInt(parts[2]),
              },
              end: {
                sheet: parseInt(parts[0]),
                col: parseInt(parts[3]),
                row: parseInt(parts[4]),
              },
            };
            this.dependencyGraph.addRange(depRange);
            this.dependencyGraph.addDependency(namedExprKey, dep);
          } else if (dep.startsWith("name:")) {
            // Named expression dependency
            const nameParts = dep.split(":");
            if (nameParts.length === 2 && nameParts[1]) {
              this.dependencyGraph.addNamedExpression(nameParts[1], undefined);
            } else if (nameParts.length === 3 && nameParts[1] && nameParts[2]) {
              this.dependencyGraph.addNamedExpression(
                nameParts[2],
                parseInt(nameParts[1])
              );
            }
            this.dependencyGraph.addDependency(namedExprKey, dep);
          }
        }
      } catch (error) {
        // Formula parsing failed, but we can still store the named expression
        console.warn(
          `Failed to parse named expression formula: ${expression}`,
          error
        );
      }
    }

    // TODO: Update dependent cells

    return [];
  }

  getNamedExpressionValue(expressionName: string, scope?: number): CellValue {
    const expr = this.getNamedExpression(expressionName, scope);
    if (!expr) return undefined;

    // TODO: Evaluate the expression
    return undefined;
  }

  getNamedExpressionFormula(expressionName: string, scope?: number): string {
    const expr = this.getNamedExpression(expressionName, scope);
    return expr?.expression || "";
  }

  getNamedExpression(
    expressionName: string,
    scope?: number
  ): NamedExpression | undefined {
    // Try scoped first
    if (scope !== undefined) {
      const scopedKey = this.getNamedExpressionKey(expressionName, scope);
      const scoped = this.namedExpressions.get(scopedKey);
      if (scoped) return scoped;
    }

    // Try global
    const globalKey = this.getNamedExpressionKey(expressionName);
    return this.namedExpressions.get(globalKey);
  }

  isItPossibleToChangeNamedExpression(
    expressionName: string,
    newExpression: RawCellContent,
    scope?: number
  ): boolean {
    const expr = this.getNamedExpression(expressionName, scope);
    return expr !== undefined;
  }

  changeNamedExpression(
    expressionName: string,
    newExpression: RawCellContent,
    scope?: number,
    options?: NamedExpressionOptions
  ): ExportedChange[] {
    if (
      !this.isItPossibleToChangeNamedExpression(
        expressionName,
        newExpression,
        scope
      )
    ) {
      return [];
    }

    const key = this.getNamedExpressionKey(expressionName, scope);
    const existing = this.namedExpressions.get(key);
    if (existing) {
      existing.expression = String(newExpression);
    }

    // TODO: Update dependent cells

    return [];
  }

  isItPossibleToRemoveNamedExpression(
    expressionName: string,
    scope?: number
  ): boolean {
    const expr = this.getNamedExpression(expressionName, scope);
    return expr !== undefined;
  }

  removeNamedExpression(
    expressionName: string,
    scope?: number
  ): ExportedChange[] {
    if (!this.isItPossibleToRemoveNamedExpression(expressionName, scope)) {
      return [];
    }

    const key = this.getNamedExpressionKey(expressionName, scope);
    this.namedExpressions.delete(key);

    // TODO: Update dependent cells

    return [];
  }

  listNamedExpressions(scope?: number): string[] {
    const result: string[] = [];

    for (const [key, expr] of this.namedExpressions) {
      if (scope === undefined && expr.scope === undefined) {
        // Global expressions
        result.push(expr.name);
      } else if (scope !== undefined && expr.scope === scope) {
        // Scoped expressions
        result.push(expr.name);
      }
    }

    return result;
  }

  getAllNamedExpressionsSerialized(): SerializedNamedExpression[] {
    const result: SerializedNamedExpression[] = [];

    for (const [key, expr] of this.namedExpressions) {
      result.push({
        ...expr,
        id: key,
      });
    }

    return result;
  }

  private getNamedExpressionKey(name: string, scope?: number): string {
    return scope === undefined ? name : `${scope}:${name}`;
  }

  // ===== Internal Helper Methods for Evaluation =====

  private getCellValueInternal(
    address: SimpleCellAddress,
    evaluationStack?: Set<string>
  ): CellValue {
    const sheet = this.sheets.get(address.sheet);
    if (!sheet) return undefined;

    const cell = getCell(sheet, address);
    if (!cell) return undefined;

    // If this cell has a formula and we have an evaluation stack, check for cycles
    if (cell.formula && evaluationStack) {
      const key = addressToKey(address);
      if (evaluationStack.has(key)) {
        return "#CYCLE!";
      }

      // If this cell needs to be re-evaluated, do it now with the current stack
      if (cell.value === undefined || isFormulaError(cell.value)) {
        try {
          const ast = parseFormula(cell.formula, address.sheet, (sheetName) =>
            this.getSheetId(sheetName)
          );

          // Create evaluation context with the existing stack
          const context: EvaluationContext = {
            currentSheet: address.sheet,
            currentCell: address,
            namedExpressions: this.namedExpressions,
            getCellValue: (addr: SimpleCellAddress) =>
              this.getCellValueInternal(addr, evaluationStack),
            getRangeValues: (range: SimpleCellRange) =>
              this.getRangeValuesInternal(range, evaluationStack),
            getFunction: (name: string) => functionRegistry.get(name),
            errorHandler: this.errorHandler,
            evaluationStack: new Set(evaluationStack),
            sheetResolver: (sheetName: string) => this.getSheetId(sheetName),
          };

          // Add current cell to stack
          context.evaluationStack.add(key);

          const result = this.evaluator.evaluate(ast, context);
          return result.type === "value" ? result.value : result.value[0]?.[0];
        } catch (error) {
          if (
            error instanceof ParseError &&
            error.message.includes("not found")
          ) {
            return "#REF!";
          }
          return "#ERROR!";
        }
      }
    }

    return cell?.value;
  }

  private getRangeValuesInternal(
    range: SimpleCellRange,
    evaluationStack?: Set<string>
  ): CellValue[][] {
    const sheet = this.sheets.get(range.start.sheet);
    if (!sheet) return [];

    // Check if this is an infinite range
    const isInfiniteColumn = range.end.row === Number.MAX_SAFE_INTEGER;
    const isInfiniteRow = range.end.col === Number.MAX_SAFE_INTEGER;

    if (isInfiniteColumn || isInfiniteRow) {
      // Handle infinite ranges using sparse iteration
      return this.getInfiniteRangeValues(
        sheet,
        range,
        isInfiniteColumn,
        isInfiniteRow,
        evaluationStack
      );
    }

    // Normal range handling
    const result: CellValue[][] = [];

    for (let row = range.start.row; row <= range.end.row; row++) {
      const rowValues: CellValue[] = [];
      for (let col = range.start.col; col <= range.end.col; col++) {
        const cellAddress = { sheet: range.start.sheet, row, col };

        // Use getCellValueInternal with evaluation stack for cycle detection
        const value = evaluationStack
          ? this.getCellValueInternal(cellAddress, evaluationStack)
          : this.getCellValue(cellAddress);

        rowValues.push(value);
      }
      result.push(rowValues);
    }

    return result;
  }

  private getInfiniteRangeValues(
    sheet: Sheet,
    range: SimpleCellRange,
    isInfiniteColumn: boolean,
    isInfiniteRow: boolean,
    evaluationStack?: Set<string>
  ): CellValue[][] {
    const result: CellValue[][] = [];

    if (isInfiniteColumn && !isInfiniteRow) {
      // Infinite column range (e.g., A:A, B:D)
      // For functions like INDEX to work, we need to return all rows up to the highest populated row
      // Find the highest populated row in the specified columns
      let maxRow = -1;

      for (const [key, cell] of sheet.cells) {
        const address = parseA1Key(key, sheet.id);
        if (
          address &&
          address.col >= range.start.col &&
          address.col <= range.end.col
        ) {
          maxRow = Math.max(maxRow, address.row);
        }
      }

      // If no cells found, return empty array
      if (maxRow === -1) {
        return [];
      }

      // Create array from row 0 to maxRow
      for (let row = 0; row <= maxRow; row++) {
        const rowValues: CellValue[] = [];
        for (let col = range.start.col; col <= range.end.col; col++) {
          const cellAddress = { sheet: sheet.id, row, col };
          const value = evaluationStack
            ? this.getCellValueInternal(cellAddress, evaluationStack)
            : (getCell(sheet, cellAddress)?.value ?? undefined);
          rowValues.push(value);
        }
        result.push(rowValues);
      }
    } else if (isInfiniteRow && !isInfiniteColumn) {
      // Infinite row range (e.g., 5:5, 1:10)
      // For functions like INDEX to work, we need to return all columns up to the highest populated column
      // Find the highest populated column in the specified rows
      let maxCol = -1;

      for (const [key, cell] of sheet.cells) {
        const address = parseA1Key(key, sheet.id);
        if (
          address &&
          address.row >= range.start.row &&
          address.row <= range.end.row
        ) {
          maxCol = Math.max(maxCol, address.col);
        }
      }

      // If no cells found, return empty array for each row
      if (maxCol === -1) {
        const rows = range.end.row - range.start.row + 1;
        for (let i = 0; i < rows; i++) {
          result.push([]);
        }
        return result;
      }

      // Create array with all columns from 0 to maxCol
      for (let row = range.start.row; row <= range.end.row; row++) {
        const rowValues: CellValue[] = [];
        for (let col = 0; col <= maxCol; col++) {
          const cellAddress = { sheet: sheet.id, row, col };
          const value = evaluationStack
            ? this.getCellValueInternal(cellAddress, evaluationStack)
            : (getCell(sheet, cellAddress)?.value ?? undefined);
          rowValues.push(value);
        }
        result.push(rowValues);
      }
    } else {
      // Both infinite (entire sheet) - not typically used but handle it
      // Return all populated cells
      const cellMap = new Map<string, CellValue>();
      let maxRow = 0;
      let maxCol = 0;

      for (const [key, cell] of sheet.cells) {
        const address = parseA1Key(key, sheet.id);
        if (address) {
          cellMap.set(`${address.row},${address.col}`, cell.value);
          maxRow = Math.max(maxRow, address.row);
          maxCol = Math.max(maxCol, address.col);
        }
      }

      // Create result array
      for (let row = 0; row <= maxRow; row++) {
        const rowValues: CellValue[] = [];
        for (let col = 0; col <= maxCol; col++) {
          const value = cellMap.get(`${row},${col}`);
          rowValues.push(value ?? undefined);
        }
        result.push(rowValues);
      }
    }

    return result;
  }

  private canSpillArray(
    spillRange: SimpleCellRange,
    originAddress: SimpleCellAddress
  ): boolean {
    const sheet = this.sheets.get(spillRange.start.sheet);
    if (!sheet) return false;

    // Check if any cells in the spill range (except origin) are non-empty
    for (let row = spillRange.start.row; row <= spillRange.end.row; row++) {
      for (let col = spillRange.start.col; col <= spillRange.end.col; col++) {
        // Skip the origin cell
        if (row === originAddress.row && col === originAddress.col) {
          continue;
        }

        const cell = getCell(sheet, {
          sheet: spillRange.start.sheet,
          row,
          col,
        });
        if (cell && !isEmptyCell(cell)) {
          return false;
        }
      }
    }

    return true;
  }

  private clearPreviousSpillRange(originAddress: SimpleCellAddress): void {
    const sheet = this.sheets.get(originAddress.sheet);
    if (!sheet) return;

    // Get the origin cell to find its previous spill range
    const originCell = getCell(sheet, originAddress);
    if (!originCell || !originCell.arrayFormula) return;

    const spillRange = originCell.arrayFormula.spillRange;

    // Clear all cells in the spill range except the origin
    for (let row = spillRange.start.row; row <= spillRange.end.row; row++) {
      for (let col = spillRange.start.col; col <= spillRange.end.col; col++) {
        if (row === originAddress.row && col === originAddress.col) {
          continue;
        }

        removeCell(sheet, { sheet: spillRange.start.sheet, row, col });
      }
    }
  }

  private spillArray(
    arrayValue: CellValue[][],
    spillRange: SimpleCellRange,
    arrayFormula: ArrayFormula
  ): void {
    const sheet = this.sheets.get(spillRange.start.sheet);
    if (!sheet) return;

    // Set each cell in the array
    for (let r = 0; r < arrayValue.length; r++) {
      for (let c = 0; c < (arrayValue[r]?.length || 0); c++) {
        const row = spillRange.start.row + r;
        const col = spillRange.start.col + c;

        // Skip the origin cell (already set)
        if (
          row === arrayFormula.originAddress.row &&
          col === arrayFormula.originAddress.col
        ) {
          continue;
        }

        const cellValue = arrayValue[r]?.[c];
        if (cellValue !== undefined) {
          const spilledCell = createArrayCell("", cellValue, arrayFormula);
          setCell(
            sheet,
            { sheet: spillRange.start.sheet, row, col },
            spilledCell
          );
        }
      }
    }
  }

  // ===== Formula Adjustment Methods =====

  private adjustFormula(
    formula: string,
    sourceAddress: SimpleCellAddress,
    targetAddress: SimpleCellAddress
  ): string {
    // Calculate offset
    const rowOffset = targetAddress.row - sourceAddress.row;
    const colOffset = targetAddress.col - sourceAddress.col;

    // Use regex to find cell references and ranges
    // Matches: A1, $A$1, A$1, $A1, A1:B2, A:A, 1:1, etc.
    // First try to match infinite column ranges (A:A)
    const infiniteColRegex = /(\$?)([A-Z]+):(\$?)([A-Z]+)/g;
    // Then infinite row ranges (1:1)
    const infiniteRowRegex = /(\$?)(\d+):(\$?)(\d+)/g;
    // Then normal cell references
    const cellRefRegex =
      /(\$?)([A-Z]+)(\$?)(\d+)(?::(\$?)([A-Z]+)(\$?)(\d+))?/g;

    // Handle infinite column ranges first (A:A)
    let result = formula.replace(
      infiniteColRegex,
      (match, startDollar, startCol, endDollar, endCol) => {
        // Helper to adjust column
        const adjustCol = (dollar: string, col: string) => {
          if (dollar) return dollar + col; // Absolute reference

          // Convert column letters to number
          let colNum = 0;
          for (let i = 0; i < col.length; i++) {
            colNum = colNum * 26 + (col.charCodeAt(i) - 64);
          }
          colNum--; // Zero-based

          // Adjust
          colNum += colOffset;
          if (colNum < 0) return "#REF!";

          // Convert back to letters
          let newCol = "";
          let tempCol = colNum + 1;
          while (tempCol > 0) {
            tempCol--;
            newCol = String.fromCharCode(65 + (tempCol % 26)) + newCol;
            tempCol = Math.floor(tempCol / 26);
          }

          return newCol;
        };

        const newStart = adjustCol(startDollar, startCol);
        const newEnd = adjustCol(endDollar, endCol);

        if (newStart === "#REF!" || newEnd === "#REF!") return "#REF!";
        return newStart + ":" + newEnd;
      }
    );

    // Handle infinite row ranges (1:1)
    result = result.replace(
      infiniteRowRegex,
      (match, startDollar, startRow, endDollar, endRow) => {
        // In Excel, row numbers in ranges like 1:1 are absolute by default
        // Only adjust if explicitly marked as relative with $ (which is rare)
        // Since our regex captures $ for absolute, no $ means it's already absolute
        if (!startDollar && !endDollar) {
          // No dollar signs means absolute in row ranges
          return match; // Keep as-is
        }

        // If there are dollar signs, they indicate relative references (opposite of normal cells)
        const adjustRow = (dollar: string, row: string) => {
          if (!dollar) return row; // No dollar = absolute for row ranges

          // Dollar sign present = relative (unusual but possible)
          let rowNum = parseInt(row) - 1; // Zero-based
          rowNum += rowOffset;

          if (rowNum < 0) return "#REF!";
          return dollar + String(rowNum + 1);
        };

        const newStart = adjustRow(startDollar, startRow);
        const newEnd = adjustRow(endDollar, endRow);

        if (newStart === "#REF!" || newEnd === "#REF!") return "#REF!";
        return newStart + ":" + newEnd;
      }
    );

    // Handle normal cell references
    return result.replace(
      cellRefRegex,
      (
        match,
        col1Dollar,
        col1,
        row1Dollar,
        row1,
        col2Dollar,
        col2,
        row2Dollar,
        row2
      ) => {
        // Helper function to adjust a single cell reference
        const adjustCell = (
          colDollar: string,
          col: string,
          rowDollar: string,
          row: string
        ) => {
          // Convert column letters to number
          let colNum = 0;
          for (let i = 0; i < col.length; i++) {
            colNum = colNum * 26 + (col.charCodeAt(i) - 64);
          }
          colNum--; // Zero-based

          // Adjust based on relative/absolute
          if (!colDollar) {
            colNum += colOffset;
          }

          let rowNum = parseInt(row) - 1; // Zero-based
          if (!rowDollar) {
            rowNum += rowOffset;
          }

          // Convert back to A1 notation
          let newCol = "";
          let tempCol = colNum + 1;
          while (tempCol > 0) {
            tempCol--;
            newCol = String.fromCharCode(65 + (tempCol % 26)) + newCol;
            tempCol = Math.floor(tempCol / 26);
          }

          return colDollar + newCol + rowDollar + (rowNum + 1);
        };

        // Adjust first cell reference
        const adjustedFirst = adjustCell(col1Dollar, col1, row1Dollar, row1);

        // If it's a range, adjust the second cell reference too
        if (col2) {
          const adjustedSecond = adjustCell(col2Dollar, col2, row2Dollar, row2);
          return adjustedFirst + ":" + adjustedSecond;
        }

        return adjustedFirst;
      }
    );
  }

  // ===== Recalculation Methods =====

  private recalculateDependents(address: SimpleCellAddress): void {
    // Use a set to track all cells that need recalculation (keys)
    const cellsToRecalculate = new Set<string>();
    const visitedForCollection = new Set<string>();

    // Collect dependents for a given cell and add to the set
    const collectDependents = (cellAddress: SimpleCellAddress) => {
      const addressKey = addressToKey(cellAddress);
      if (visitedForCollection.has(addressKey)) return;
      visitedForCollection.add(addressKey);

      // Direct dependents
      const direct = this.dependencyGraph.getDependents(addressKey);
      for (const dep of direct) cellsToRecalculate.add(dep);

      // Dependents from ranges containing this cell
      const rangeKeys =
        this.dependencyGraph.getRangesContainingCell(cellAddress);
      for (const rangeKey of rangeKeys) {
        const deps = this.dependencyGraph.getDependents(rangeKey);
        for (const dep of deps) cellsToRecalculate.add(dep);
      }
    };

    // Seed with initial address dependents
    collectDependents(address);

    // Track recalculated cells to avoid duplicate work
    const recalculated = new Set<string>();

    // Iteratively process until no new dependents remain
    while (true) {
      const pending = [...cellsToRecalculate].filter(
        (k) => !recalculated.has(k)
      );
      if (pending.length === 0) break;

      const ordered = this.sortDependentsTopologically(pending);
      for (const depKey of ordered) {
        if (recalculated.has(depKey)) continue;
        recalculated.add(depKey);
        this.recalculateCell(depKey);

        // After recalculation, collect further dependents transitively
        const parts = depKey.split(":");
        if (parts.length === 3 && parts[0] && parts[1] && parts[2]) {
          const depAddress: SimpleCellAddress = {
            sheet: parseInt(parts[0]),
            col: parseInt(parts[1]),
            row: parseInt(parts[2]),
          };
          collectDependents(depAddress);
        }
      }
    }
  }

  private sortDependentsTopologically(dependents: string[]): string[] {
    // Simple implementation - in a real system this would use a proper topological sort
    // For now, just return as-is since our dependency graph handles cycles
    return dependents;
  }

  private recalculateCell(cellKey: string): void {
    // Parse the cell key
    const parts = cellKey.split(":");
    if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) return;

    const address: SimpleCellAddress = {
      sheet: parseInt(parts[0]),
      col: parseInt(parts[1]),
      row: parseInt(parts[2]),
    };

    const sheet = this.sheets.get(address.sheet);
    if (!sheet) return;

    const cell = getCell(sheet, address);
    if (!cell || !cell.formula) return;

    const oldValue = cell.value;

    // Re-evaluate the formula
    try {
      const ast = parseFormula(cell.formula, address.sheet, (sheetName) =>
        this.getSheetId(sheetName)
      );

      const context: EvaluationContext = {
        currentSheet: address.sheet,
        currentCell: address,
        namedExpressions: this.namedExpressions,
        getCellValue: (addr: SimpleCellAddress) =>
          this.getCellValueInternal(addr, context.evaluationStack),
        getRangeValues: (range: SimpleCellRange) =>
          this.getRangeValuesInternal(range, context.evaluationStack),
        getFunction: (name: string) => functionRegistry.get(name),
        errorHandler: this.errorHandler,
        evaluationStack: new Set<string>(),
        sheetResolver: (sheetName: string) => this.getSheetId(sheetName),
      };

      // Add current cell to evaluation stack to detect cycles
      context.evaluationStack.add(cellKey);

      const result = this.evaluator.evaluate(ast, context);

      // Update the cell value
      if (result.type === "2d-array") {
        // Handle array spilling for recalculation
        const arrayValue = result.value;
        cell.value = arrayValue[0]?.[0] ?? undefined;

        // If this is an array formula origin, update the spilled values
        if (
          cell.arrayFormula &&
          cell.arrayFormula.originAddress.row === address.row &&
          cell.arrayFormula.originAddress.col === address.col
        ) {
          this.spillArray(
            arrayValue,
            cell.arrayFormula.spillRange,
            cell.arrayFormula
          );
        }
      } else {
        cell.value = result.value;
      }
    } catch (error) {
      if (error instanceof ParseError && error.message.includes("not found")) {
        cell.value = "#REF!";
      } else if (error instanceof Error && error.message.startsWith("#")) {
        cell.value = error.message as any;
      } else {
        cell.value = "#ERROR!";
      }
    }

    // Notify listeners if the value actually changed
    if (oldValue !== cell.value) {
      this.emitCellUpdate({
        address,
        oldValue: oldValue ?? undefined,
        newValue: cell.value,
      });
    }
  }

  // ===== Formula Utilities =====

  normalizeFormula(formulaString: string): string {
    // TODO: Implement formula normalization
    return formulaString.trim();
  }

  calculateFormula(formulaString: string, sheetId: number): CellValue {
    // TODO: Implement formula calculation
    return undefined;
  }

  getNamedExpressionsFromFormula(formulaString: string): string[] {
    // TODO: Parse formula and extract named expressions
    return [];
  }

  validateFormula(formulaString: string): boolean {
    // TODO: Implement formula validation
    return true;
  }

  // ===== Undo/Redo System =====

  undo(): ExportedChange[] {
    if (this.undoStack.length === 0) return [];

    const command = this.undoStack.pop()!;
    const changes = command.undo();
    this.redoStack.push(command);

    return changes;
  }

  redo(): ExportedChange[] {
    if (this.redoStack.length === 0) return [];

    const command = this.redoStack.pop()!;
    const changes = command.redo();
    this.undoStack.push(command);

    return changes;
  }

  isThereSomethingToUndo(): boolean {
    return this.undoStack.length > 0;
  }

  isThereSomethingToRedo(): boolean {
    return this.redoStack.length > 0;
  }

  clearRedoStack(): void {
    this.redoStack = [];
  }

  clearUndoStack(): void {
    this.undoStack = [];
  }

  // ===== Event System =====

  /**
   * Subscribe to FormulaEngine events
   * @param event The event name
   * @param listener The event listener function
   * @returns Unsubscribe function
   */
  on<K extends keyof FormulaEngineEvents>(
    event: K,
    listener: EventListener<FormulaEngineEvents[K]>
  ): () => void {
    return this.eventEmitter.on(event, listener);
  }

  /**
   * Subscribe to FormulaEngine events (alias for on)
   * @param event The event name
   * @param listener The event listener function
   * @returns Unsubscribe function
   */
  subscribe<K extends keyof FormulaEngineEvents>(
    event: K,
    listener: EventListener<FormulaEngineEvents[K]>
  ): () => void {
    return this.eventEmitter.on(event, listener);
  }

  /**
   * Remove all event listeners
   */
  removeAllListeners(): void {
    this.eventEmitter.removeAllListeners();
  }

  /**
   * Emit an event (internal use)
   */
  private emit<K extends keyof FormulaEngineEvents>(
    event: K,
    data: FormulaEngineEvents[K]
  ): void {
    this.eventEmitter.emit(event, data);
  }

  // ===== New Event API: Cell and Sheet Updates =====

  /**
   * Register listener for specific cell updates. Returns an unsubscribe function.
   */
  onCellUpdate(
    address: SimpleCellAddress,
    listener: (event: {
      address: SimpleCellAddress;
      oldValue: CellValue;
      newValue: CellValue;
    }) => void
  ): () => void {
    const key = addressToKey(address);
    if (!this.cellUpdateListeners.has(key)) {
      this.cellUpdateListeners.set(key, new Set());
    }
    const set = this.cellUpdateListeners.get(key)!;
    set.add(listener);
    return () => {
      const listeners = this.cellUpdateListeners.get(key);
      if (listeners) {
        listeners.delete(listener);
        if (listeners.size === 0) this.cellUpdateListeners.delete(key);
      }
    };
  }

  /**
   * Register listener for batched sheet updates. Returns an unsubscribe function.
   */
  onCellsUpdate(
    sheetId: number,
    listener: (
      events: {
        address: SimpleCellAddress;
        oldValue: CellValue;
        newValue: CellValue;
      }[]
    ) => void
  ): () => void {
    if (!this.cellsUpdateListeners.has(sheetId)) {
      this.cellsUpdateListeners.set(sheetId, new Set());
    }
    const set = this.cellsUpdateListeners.get(sheetId)!;
    set.add(listener);
    return () => {
      const listeners = this.cellsUpdateListeners.get(sheetId);
      if (listeners) {
        listeners.delete(listener);
        if (listeners.size === 0) this.cellsUpdateListeners.delete(sheetId);
      }
    };
  }

  private beginBatch(): void {
    this.batchDepth++;
  }

  private endBatch(): void {
    this.batchDepth = Math.max(0, this.batchDepth - 1);
    if (this.batchDepth === 0) {
      this.processDeferredEvaluations();
      this.flushSheetUpdates();
    }
  }

  private queueSheetUpdate(event: {
    address: SimpleCellAddress;
    oldValue: CellValue;
    newValue: CellValue;
  }): void {
    const SheetId = event.address.sheet;
    if (!this.pendingSheetUpdates.has(SheetId)) {
      this.pendingSheetUpdates.set(SheetId, []);
    }
    this.pendingSheetUpdates.get(SheetId)!.push(event);
  }

  private processDeferredEvaluations(): void {
    if (this.deferredEvaluations.size === 0) return;

    // Sort deferred evaluations by dependency order
    const sortedEvaluations = this.sortByDependencyOrder(
      Array.from(this.deferredEvaluations.values())
    );

    // Process evaluations in dependency order
    for (const { address, formula, ast } of sortedEvaluations) {
      const sheet = this.sheets.get(address.sheet);
      if (!sheet) continue;

      try {
        // Create evaluation context
        const context: EvaluationContext = {
          currentSheet: address.sheet,
          currentCell: address,
          namedExpressions: this.namedExpressions,
          getCellValue: (addr: SimpleCellAddress) =>
            this.getCellValueInternal(addr, context.evaluationStack),
          getRangeValues: (range: SimpleCellRange) =>
            this.getRangeValuesInternal(range, context.evaluationStack),
          getFunction: (name: string) => functionRegistry.get(name),
          errorHandler: this.errorHandler,
          evaluationStack: new Set<string>(),
          sheetResolver: (sheetName: string) => this.getSheetId(sheetName),
        };

        const addressKey = addressToKey(address);
        context.evaluationStack.add(addressKey);

        // Evaluate the formula
        const result = this.evaluator.evaluate(ast, context);

        // Register dependencies
        for (const dep of result.dependencies) {
          const parsed = parseDependencyKey(dep);

          switch (parsed.type) {
            case "cell":
              this.dependencyGraph.addCell(parsed.address);
              this.dependencyGraph.addDependency(addressKey, dep);
              break;

            case "range":
              this.dependencyGraph.addRange(parsed.range);
              this.dependencyGraph.addDependency(addressKey, dep);
              break;

            case "named":
              this.dependencyGraph.addNamedExpression(
                parsed.name,
                parsed.scope
              );
              this.dependencyGraph.addDependency(addressKey, dep);
              break;

            default:
              throw new Error(
                `Invalid dependency type: ${(parsed as any).type}`
              );
          }
        }

        // Update the cell with the evaluated value and handle array spilling
        const currentCell = getCell(sheet, address);
        if (currentCell && currentCell.type === "FORMULA") {
          const oldValue = currentCell.value;

          // Handle array spilling (copied from setCellValue logic)
          if (result.type === "2d-array") {
            // This is a 2D array - implement array spilling
            const arrayValue = result.value;
            const rows = arrayValue.length;
            const cols = arrayValue[0]?.length ?? 0;

            if (cols > 0) {
              // Calculate spill range
              const spillRange: SimpleCellRange = {
                start: address,
                end: {
                  sheet: address.sheet,
                  row: address.row + rows - 1,
                  col: address.col + cols - 1,
                },
              };

              // Check if spill range is available (no non-empty cells)
              if (this.canSpillArray(spillRange, address)) {
                // Clear any previous spill range for this formula
                this.clearPreviousSpillRange(address);

                // Create array formula info
                const arrayFormula: ArrayFormula = {
                  originAddress: address,
                  spillRange,
                  formula: formula, // Include the formula string
                };

                // Convert to array cell and set origin cell info
                currentCell.type = "ARRAY";
                currentCell.value = arrayValue[0]?.[0] ?? undefined;
                currentCell.arrayFormula = arrayFormula;

                // Spill the array values
                this.spillArray(arrayValue, spillRange, arrayFormula);
              } else {
                // Spill blocked - return #SPILL! error
                currentCell.value = "#SPILL!";
              }
            } else {
              // Empty array - set as error
              currentCell.value = "#N/A";
            }
          } else {
            // Regular scalar value
            currentCell.value = result.value as CellValue;
          }

          // Emit cell update
          this.emitCellUpdate({
            address,
            oldValue,
            newValue: currentCell.value,
          });
        }
      } catch (error) {
        // Handle evaluation error
        const currentCell = getCell(sheet, address);
        if (currentCell && currentCell.type === "FORMULA") {
          const oldValue = currentCell.value;
          currentCell.value = "#ERROR!";
          this.emitCellUpdate({
            address,
            oldValue,
            newValue: currentCell.value,
          });
        }
      }
    }

    // Clear deferred evaluations
    this.deferredEvaluations.clear();
  }

  private sortByDependencyOrder(
    evaluations: { address: SimpleCellAddress; formula: string; ast: any }[]
  ): { address: SimpleCellAddress; formula: string; ast: any }[] {
    // Get topological sort order from dependency graph
    const topologicalOrder = this.dependencyGraph.topologicalSort();

    if (!topologicalOrder) {
      // Cycle detected - log warning and fall back to simple heuristic
      console.warn(
        "Circular dependency detected during deferred evaluation. Using fallback ordering."
      );
      return evaluations.sort((a, b) => {
        if (a.address.row !== b.address.row) {
          return a.address.row - b.address.row;
        }
        return a.address.col - b.address.col;
      });
    }

    // Create a map from cell key to evaluation object for quick lookup
    const evaluationMap = new Map<
      string,
      { address: SimpleCellAddress; formula: string; ast: any }
    >();
    for (const evaluation of evaluations) {
      const key = addressToKey(evaluation.address);
      evaluationMap.set(key, evaluation);
    }

    // Sort evaluations based on topological order
    const sortedEvaluations: {
      address: SimpleCellAddress;
      formula: string;
      ast: any;
    }[] = [];

    // First, add evaluations that appear in the topological order
    for (const nodeKey of topologicalOrder) {
      const evaluation = evaluationMap.get(nodeKey);
      if (evaluation) {
        sortedEvaluations.push(evaluation);
        evaluationMap.delete(nodeKey);
      }
    }

    // Add any remaining evaluations that weren't in the dependency graph yet
    // (this can happen if dependencies weren't registered yet)
    for (const evaluation of evaluationMap.values()) {
      sortedEvaluations.push(evaluation);
    }

    return sortedEvaluations;
  }

  private flushSheetUpdates(): void {
    if (this.pendingSheetUpdates.size === 0) return;
    const updates = this.pendingSheetUpdates;
    this.pendingSheetUpdates = new Map();
    for (const [sheetId, events] of updates) {
      const listeners = this.cellsUpdateListeners.get(sheetId);
      if (listeners && events.length > 0) {
        // Deliver one batched callback per listener
        for (const listener of listeners) {
          listener(events);
        }
      }
    }
  }

  private emitCellUpdate(event: {
    address: SimpleCellAddress;
    oldValue: CellValue;
    newValue: CellValue;
  }): void {
    // Immediate cell listener callbacks
    const key = addressToKey(event.address);
    const listeners = this.cellUpdateListeners.get(key);
    if (listeners) {
      for (const listener of listeners) listener(event);
    }
    // Queue for batched sheet listeners
    this.queueSheetUpdate(event);
    // If not in a batch, flush immediately for sheet listeners
    if (this.batchDepth === 0) {
      this.flushSheetUpdates();
    }
  }
}

// Command interface for undo/redo
interface Command {
  execute(): ExportedChange[];
  undo(): ExportedChange[];
  redo(): ExportedChange[];
}
