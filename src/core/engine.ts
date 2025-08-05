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
  FormulaEngineOptions
} from './types';

import {
  addressToKey,
  getCellValueType,
  isFormulaError,
  isCellEmpty
} from './types';

import {
  parseCellAddress,
  parseCellRange,
  addressToA1,
  rangeToA1,
  isValidAddress,
  iterateRange
} from './address';

import {
  createValueCell,
  createFormulaCell,
  createEmptyCell,
  isFormula,
  extractFormula,
  parseInputValue,
  serializeCell,
  deserializeCell,
  hasFormula,
  isArrayCell,
  isEmptyCell,
  getCellDisplayValue
} from './cell';

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
  shiftCells
} from './sheet';

/**
 * Main FormulaEngine class
 */
export class FormulaEngine {
  private sheets: Map<number, Sheet> = new Map();
  private namedExpressions: Map<string, NamedExpression> = new Map();
  private nextSheetId: number = 0;
  private evaluationSuspended: boolean = false;
  private pendingChanges: ExportedChange[] = [];
  private clipboard: { range: SimpleCellRange; data: RawCellContent[][] } | null = null;
  private undoStack: Command[] = [];
  private redoStack: Command[] = [];
  private options: FormulaEngineOptions;

  constructor(options: FormulaEngineOptions = {}) {
    this.options = {
      evaluationMode: 'eager',
      maxIterations: 100,
      cacheSize: 1000,
      enableArrayFormulas: true,
      enableNamedExpressions: true,
      locale: 'en-US',
      ...options
    };
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
    if (!sheet) return '';
    
    const cell = getCell(sheet, cellAddress);
    return cell?.formula || '';
  }

  getCellSerialized(cellAddress: SimpleCellAddress): RawCellContent {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return undefined;
    
    const cell = getCell(sheet, cellAddress);
    return serializeCell(cell);
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
      const serialized = serializeCell(cell);
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

  setCellContents(
    topLeftCornerAddress: SimpleCellAddress,
    cellContents: RawCellContent[][] | RawCellContent
  ): ExportedChange[] {
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
            row: topLeftCornerAddress.row + row
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

  private setCellValue(address: SimpleCellAddress, content: RawCellContent): ExportedChange | null {
    const sheet = this.sheets.get(address.sheet);
    if (!sheet) return null;
    
    const oldCell = getCell(sheet, address);
    const oldValue = oldCell?.value;
    
    // Parse the content and create appropriate cell
    let newCell: Cell;
    
    if (content === undefined || content === null || content === '') {
      // Empty cell
      removeCell(sheet, address);
      
      if (oldValue !== undefined) {
        return {
          address,
          oldValue,
          newValue: undefined,
          type: 'cell-change'
        };
      }
      return null;
    } else if (typeof content === 'string' && isFormula(content)) {
      // Formula cell
      const formula = extractFormula(content);
      newCell = createFormulaCell(formula);
      
      // TODO: Parse formula and evaluate
      // For now, just store the formula
      newCell.value = content; // Temporary
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
        type: 'cell-change'
      };
      
      if (this.evaluationSuspended) {
        this.pendingChanges.push(change);
        return null;
      }
      
      return change;
    }
    
    return null;
  }

  setSheetContents(sheetId: number, contents: Map<string, RawCellContent>): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];
    
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
    
    // Remove cells not in the new contents
    const keysToRemove: string[] = [];
    for (const key of sheet.cells.keys()) {
      if (!processedKeys.has(key)) {
        keysToRemove.push(key);
      }
    }
    
    for (const key of keysToRemove) {
      const address = parseCellAddress(key, sheetId);
      if (address) {
        const oldCell = getCell(sheet, address);
        if (oldCell) {
          removeCell(sheet, address);
          changes.push({
            address,
            oldValue: oldCell.value,
            newValue: undefined,
            type: 'cell-change'
          });
        }
      }
    }
    
    return this.evaluationSuspended ? [] : changes;
  }

  getSheetContents(sheetId: number): Map<string, CellValue> {
    return this.getSheetValues(sheetId);
  }

  getRangeValues(source: SimpleCellRange): CellValue[][] {
    const sheet = this.sheets.get(source.start.sheet);
    if (!sheet) return [];
    
    return getSheetRangeValues(sheet, source);
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
          row: source.start.row + row
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
          row: source.start.row + row
        };
        const cell = getCell(sheet, address);
        rowData.push(serializeCell(cell));
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
    if (!sheetName || typeof sheetName !== 'string') return false;
    
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
    
    // Collect all cell changes
    const changes: ExportedChange[] = [];
    for (const [key, cell] of sheet.cells) {
      const address = parseCellAddress(key, sheetId);
      if (address && cell.value !== undefined) {
        changes.push({
          address,
          oldValue: cell.value,
          newValue: undefined,
          type: 'cell-change'
        });
      }
    }
    
    this.sheets.delete(sheetId);
    
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
          type: 'cell-change'
        });
      }
    }
    
    clearSheetCells(sheet);
    
    return changes;
  }

  isItPossibleToReplaceSheetContent(sheetId: number, values: RawCellContent[][]): boolean {
    return this.sheets.has(sheetId);
  }

  setSheetContent(sheetId: number, values: RawCellContent[][]): ExportedChange[] {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];
    
    // Clear existing content
    const clearChanges = this.clearSheet(sheetId);
    
    // Set new content
    const setChanges = this.setCellContents(
      { sheet: sheetId, col: 0, row: 0 },
      values
    );
    
    return [...clearChanges, ...setChanges];
  }

  getSheetName(sheetId: number): string {
    const sheet = this.sheets.get(sheetId);
    return sheet ? sheet.name : '';
  }

  getSheetNames(): string[] {
    return Array.from(this.sheets.values()).map(sheet => sheet.name);
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
    if (!newName || typeof newName !== 'string') return false;
    
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
      sheet.name = newName;
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
            type: 'cell-change'
          });
        }
      }
      
      // Shift cells down
      shiftCells(sheet, 'row', rowIndex + 1, -1);
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
            type: 'cell-change'
          });
        }
      }
      
      // Shift cells left
      shiftCells(sheet, 'col', colIndex + 1, -1);
    }
    
    return changes;
  }

  isItPossibleToMoveCells(source: SimpleCellRange, destinationLeftCorner: SimpleCellAddress): boolean {
    // Check if sheets match
    if (source.start.sheet !== destinationLeftCorner.sheet) return false;
    
    // Check if destination is valid
    return isValidAddress(destinationLeftCorner);
  }

  moveCells(source: SimpleCellRange, destinationLeftCorner: SimpleCellAddress): ExportedChange[] {
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
      end: { sheet: sheetId, col: 16383, row: startRow + numberOfRows - 1 } // Excel max columns
    };
    
    const destination: SimpleCellAddress = {
      sheet: sheetId,
      col: 0,
      row: targetRow
    };
    
    return this.moveCells(source, destination);
  }

  isItPossibleToMoveColumns(
    sheetId: number,
    startColumn: number,
    numberOfColumns: number,
    targetColumn: number
  ): boolean {
    return this.sheets.has(sheetId) && startColumn >= 0 && numberOfColumns > 0 && targetColumn >= 0;
  }

  moveColumns(
    sheetId: number,
    startColumn: number,
    numberOfColumns: number,
    targetColumn: number
  ): ExportedChange[] {
    if (!this.isItPossibleToMoveColumns(sheetId, startColumn, numberOfColumns, targetColumn)) {
      return [];
    }
    
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return [];
    
    const source: SimpleCellRange = {
      start: { sheet: sheetId, col: startColumn, row: 0 },
      end: { sheet: sheetId, col: startColumn + numberOfColumns - 1, row: 1048575 } // Excel max rows
    };
    
    const destination: SimpleCellAddress = {
      sheet: sheetId,
      col: targetColumn,
      row: 0
    };
    
    return this.moveCells(source, destination);
  }

  copy(source: SimpleCellRange): CellValue[][] {
    const values = this.getRangeValues(source);
    
    // Store in clipboard
    this.clipboard = {
      range: source,
      data: this.getRangeSerialized(source)
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
    const { range, data } = this.clipboard;
    
    for (let row = 0; row < data.length; row++) {
      const rowData = data[row];
      if (rowData) {
        for (let col = 0; col < rowData.length; col++) {
          const address: SimpleCellAddress = {
            sheet: targetLeftCorner.sheet,
            col: targetLeftCorner.col + col,
            row: targetLeftCorner.row + row
          };
          
          const change = this.setCellValue(address, rowData[col]);
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

  simpleCellAddressFromString(cellAddress: string, contextSheetId: number): SimpleCellAddress {
    const parsed = parseCellAddress(cellAddress, contextSheetId);
    if (!parsed) {
      throw new Error(`Invalid cell address: ${cellAddress}`);
    }
    return parsed;
  }

  simpleCellRangeFromString(cellRange: string, contextSheetId: number): SimpleCellRange {
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
    if (typeof optionsOrContextSheetId === 'number') {
      // Legacy API - context sheet ID
      return rangeToA1(cellRange);
    }
    
    const sheet = this.sheets.get(cellRange.start.sheet);
    const sheetName = sheet?.name;
    
    return rangeToA1(cellRange, {
      includeSheetName: optionsOrContextSheetId.includeSheetName,
      sheetName
    });
  }

  // ===== Dependency Analysis =====

  getCellDependents(address: SimpleCellAddress | SimpleCellRange): (SimpleCellRange | SimpleCellAddress)[] {
    // TODO: Implement dependency tracking
    return [];
  }

  getCellPrecedents(address: SimpleCellAddress | SimpleCellRange): (SimpleCellRange | SimpleCellAddress)[] {
    // TODO: Implement dependency tracking
    return [];
  }

  // ===== Cell Information =====

  getCellType(cellAddress: SimpleCellAddress): CellType {
    const sheet = this.sheets.get(cellAddress.sheet);
    if (!sheet) return 'EMPTY';
    
    const cell = getCell(sheet, cellAddress);
    return cell?.type || 'EMPTY';
  }

  doesCellHaveSimpleValue(cellAddress: SimpleCellAddress): boolean {
    return this.getCellType(cellAddress) === 'VALUE';
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

  getCellValueDetailedType(cellAddress: SimpleCellAddress): CellValueDetailedType {
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
    if (!expressionName || typeof expressionName !== 'string') return false;
    
    const key = this.getNamedExpressionKey(expressionName, scope);
    return !this.namedExpressions.has(key);
  }

  addNamedExpression(
    expressionName: string,
    expression: RawCellContent,
    scope?: number,
    options?: NamedExpressionOptions
  ): ExportedChange[] {
    if (!this.isItPossibleToAddNamedExpression(expressionName, expression, scope)) {
      if (options?.overwrite) {
        return this.changeNamedExpression(expressionName, expression, scope, options);
      }
      return [];
    }
    
    const key = this.getNamedExpressionKey(expressionName, scope);
    const namedExpr: NamedExpression = {
      name: expressionName,
      expression: String(expression),
      scope
    };
    
    this.namedExpressions.set(key, namedExpr);
    
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
    return expr?.expression || '';
  }

  getNamedExpression(expressionName: string, scope?: number): NamedExpression | undefined {
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
    if (!this.isItPossibleToChangeNamedExpression(expressionName, newExpression, scope)) {
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

  isItPossibleToRemoveNamedExpression(expressionName: string, scope?: number): boolean {
    const expr = this.getNamedExpression(expressionName, scope);
    return expr !== undefined;
  }

  removeNamedExpression(expressionName: string, scope?: number): ExportedChange[] {
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
        id: key
      });
    }
    
    return result;
  }

  private getNamedExpressionKey(name: string, scope?: number): string {
    return scope === undefined ? `global:${name}` : `sheet:${scope}:${name}`;
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
}

// Command interface for undo/redo
interface Command {
  execute(): ExportedChange[];
  undo(): ExportedChange[];
  redo(): ExportedChange[];
}
