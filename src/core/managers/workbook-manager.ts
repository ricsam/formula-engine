import {
  FormulaError,
  type CellAddress,
  type FiniteSpreadsheetRange,
  type LocalCellAddress,
  type SerializedCellValue,
  type Sheet,
  type SpreadsheetRange,
  type Workbook,
} from "../types";
import type { WorkbookManagerSnapshot } from "../engine-snapshot";
import { getCellReference, parseCellReference } from "../utils";

import type { RangeAddress } from "../types";
import { buildRangeEvalOrder } from "./range-eval-order-builder";
import {
  EvaluationError,
  SheetNotFoundError,
  WorkbookNotFoundError,
} from "../../evaluator/evaluation-error";
import { normalizeSerializedCellValue } from "../../parser/formatter";

interface IndexEntry {
  number: number;
  key: string;
}

export interface SheetIndexes {
  // lookup maps - cells grouped by row/column
  rowGroups: Map<number, IndexEntry[]>; // row number -> cells in that row (sorted by col)
  colGroups: Map<number, IndexEntry[]>; // col number -> cells in that col (sorted by row)

  // Sorted flat indexes - for finding cells before a given row/col
  cellsSortedByRow: IndexEntry[];
  cellsSortedByCol: IndexEntry[];
}

/**
 * Utility class for binary search operations on IndexEntry arrays
 */
export class IndexEntryBinarySearch {
  /**
   * Find the insertion point for a number in a sorted IndexEntry array
   * Returns the index where the number should be inserted to maintain sort order
   */
  static findInsertionPoint(entries: IndexEntry[], target: number): number {
    let left = 0;
    let right = entries.length;

    while (left < right) {
      const mid = Math.floor((left + right) / 2);
      const midEntry = entries[mid];
      if (midEntry && midEntry.number < target) {
        left = mid + 1;
      } else {
        right = mid;
      }
    }

    return left;
  }

  /**
   * Find the first element >= target
   * Returns the index of the first element, or -1 if not found
   */
  static findFirstGreaterOrEqual(
    entries: IndexEntry[],
    target: number
  ): number {
    if (entries.length === 0) return -1;

    let left = 0;
    let right = entries.length - 1;
    let result = -1;

    while (left <= right) {
      const mid = Math.floor((left + right) / 2);
      const midEntry = entries[mid];
      if (midEntry && midEntry.number >= target) {
        result = mid;
        right = mid - 1;
      } else {
        left = mid + 1;
      }
    }

    return result;
  }

  /**
   * Find the rightmost position where we could insert a target value
   * Useful for finding elements that come before a target
   */
  static findRightmostInsertionPoint(
    entries: IndexEntry[],
    target: number
  ): number {
    return IndexEntryBinarySearch.findInsertionPoint(entries, target);
  }
}

export class WorkbookManager {
  private workbooks: Map<string, Workbook> = new Map();

  // Map from "workbookName|sheetName" to indexes
  private sheetIndexes: Map<string, SheetIndexes> = new Map();

  /**
   * Generate a key for the sheet indexes map
   */
  private getSheetIndexKey(workbookName: string, sheetName: string): string {
    return `${workbookName}|${sheetName}`;
  }

  /**
   * Get or create indexes for a sheet
   */
  public getSheetIndexes(opts: {
    workbookName: string;
    sheetName: string;
  }): SheetIndexes {
    const key = this.getSheetIndexKey(opts.workbookName, opts.sheetName);
    let indexes = this.sheetIndexes.get(key);

    if (!indexes) {
      indexes = {
        rowGroups: new Map(),
        colGroups: new Map(),
        cellsSortedByRow: [],
        cellsSortedByCol: [],
      };
      this.sheetIndexes.set(key, indexes);
    }

    return indexes;
  }

  getSheets(workbookName: string): Map<string, Sheet> {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }
    return workbook.sheets;
  }

  getWorkbooks(): Map<string, Workbook> {
    return this.workbooks;
  }

  addWorkbook(workbookName: string): void {
    if (this.workbooks.has(workbookName)) {
      throw new Error("Workbook already exists");
    }
    this.workbooks.set(workbookName, {
      name: workbookName,
      sheets: new Map(),
      workbookMetadata: undefined,
    });
  }

  removeWorkbook(workbookName: string): void {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }

    // Clean up indexes for all sheets in this workbook
    for (const sheetName of workbook.sheets.keys()) {
      const key = this.getSheetIndexKey(workbookName, sheetName);
      this.sheetIndexes.delete(key);
    }

    this.workbooks.delete(workbookName);
  }

  isContentEmpty(content: SerializedCellValue): boolean {
    return content === "" || content === undefined;
  }

  renameWorkbook(opts: {
    workbookName: string;
    newWorkbookName: string;
  }): void {
    const workbook = this.workbooks.get(opts.workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
    }

    // Update indexes for all sheets in this workbook
    for (const sheetName of workbook.sheets.keys()) {
      const oldKey = this.getSheetIndexKey(opts.workbookName, sheetName);
      const newKey = this.getSheetIndexKey(opts.newWorkbookName, sheetName);
      const indexes = this.sheetIndexes.get(oldKey);
      if (indexes) {
        this.sheetIndexes.set(newKey, indexes);
        this.sheetIndexes.delete(oldKey);
      }
    }

    this.workbooks.set(opts.newWorkbookName, workbook);
    this.workbooks.delete(opts.workbookName);
    workbook.name = opts.newWorkbookName;
  }

  resetWorkbooks(workbooks: Map<string, Workbook>): void {
    this.workbooks.clear();
    this.sheetIndexes.clear();

    workbooks.forEach((workbook, workbookName) => {
      this.workbooks.set(workbookName, workbook);
      workbook.sheets.forEach((sheet) => {
        // Initialize indexes for this sheet
        const indexes = this.getSheetIndexes({
          workbookName,
          sheetName: sheet.name,
        });
        indexes.rowGroups.clear();
        indexes.colGroups.clear();
        indexes.cellsSortedByRow = [];
        indexes.cellsSortedByCol = [];

        sheet.content.forEach((value, key) => {
          this.setCellContent(
            {
              workbookName,
              sheetName: sheet.name,
              colIndex: parseCellReference(key).colIndex,
              rowIndex: parseCellReference(key).rowIndex,
            },
            value,
            {
              sheet,
              buildingFromScratch: true,
            }
          );
        });
      });
    });
  }

  toSnapshot(): WorkbookManagerSnapshot {
    return this.getWorkbooks();
  }

  restoreFromSnapshot(snapshot: WorkbookManagerSnapshot): void {
    this.resetWorkbooks(snapshot);
  }

  getSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Sheet | undefined {
    const workbook = this.workbooks.get(workbookName);
    const sheet = workbook?.sheets.get(sheetName);
    return sheet;
  }

  addSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Sheet {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }
    const sheet = {
      name: sheetName,
      index: workbook.sheets.size,
      content: new Map(),
      metadata: new Map(),
      sheetMetadata: undefined,
    };

    if (workbook.sheets.has(sheet.name)) {
      throw new Error("Sheet already exists");
    }

    workbook.sheets.set(sheetName, sheet);

    // Initialize empty indexes for this sheet
    this.getSheetIndexes({ workbookName, sheetName });

    return sheet;
  }

  removeSheet({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Sheet {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Remove the sheet
    workbook.sheets.delete(sheetName);

    // Clean up indexes for this sheet
    const key = this.getSheetIndexKey(workbookName, sheetName);
    this.sheetIndexes.delete(key);

    return sheet;
  }

  renameSheet({
    workbookName,
    sheetName,
    newSheetName,
  }: {
    workbookName: string;
    sheetName: string;
    newSheetName: string;
  }): Sheet {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new SheetNotFoundError(sheetName);
    }

    if (workbook.sheets.has(newSheetName)) {
      throw new Error("Sheet with new name already exists");
    }

    // Update sheet name
    sheet.name = newSheetName;

    // Update sheets map
    workbook.sheets.set(newSheetName, sheet);
    workbook.sheets.delete(sheetName);

    // Move indexes to new key
    const oldKey = this.getSheetIndexKey(workbookName, sheetName);
    const newKey = this.getSheetIndexKey(workbookName, newSheetName);
    const indexes = this.sheetIndexes.get(oldKey);
    if (indexes) {
      this.sheetIndexes.set(newKey, indexes);
      this.sheetIndexes.delete(oldKey);
    }

    return sheet;
  }

  updateAllFormulas(updateCallback: (formula: string) => string): CellAddress[] {
    const changed: CellAddress[] = [];

    const update = (workbookName: string, map: Map<string, Sheet>) => {
      map.forEach((sheet, sheetName) => {
        sheet.content.forEach((cell, key) => {
          if (typeof cell === "string" && cell.startsWith("=")) {
            const formula = cell.slice(1);
            const updatedFormula = updateCallback(formula);

            // Only update if the formula actually changed
            if (updatedFormula !== formula) {
              sheet.content.set(key, `=${updatedFormula}`);
              const { colIndex, rowIndex } = parseCellReference(key);
              changed.push({
                workbookName,
                sheetName,
                colIndex,
                rowIndex,
              });
            }
          }
        });
      });
    };

    this.workbooks.forEach((workbook, workbookName) => {
      update(workbookName, workbook.sheets);
    });

    return changed;
  }

  updateFormulasExcluding(
    excludeCellsSet: Set<string>,
    updateCallback: (formula: string) => string
  ): void {
    this.workbooks.forEach((workbook, workbookName) => {
      workbook.sheets.forEach((sheet, sheetName) => {
        sheet.content.forEach((cell, key) => {
          if (typeof cell === "string" && cell.startsWith("=")) {
            const { colIndex, rowIndex } = parseCellReference(key);
            const cellKey = `${workbookName}:${sheetName}:${colIndex}:${rowIndex}`;
            
            // Skip if this cell is in the exclude set
            if (excludeCellsSet.has(cellKey)) {
              return;
            }

            const formula = cell.slice(1);
            const updatedFormula = updateCallback(formula);

            // Only update if the formula actually changed
            if (updatedFormula !== formula) {
              sheet.content.set(key, `=${updatedFormula}`);
            }
          }
        });
      });
    });
  }

  updateFormulasForWorkbook(
    workbookName: string,
    updateCallback: (formula: string) => string
  ): void {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }

    workbook.sheets.forEach((sheet) => {
      sheet.content.forEach((cell, key) => {
        if (typeof cell === "string" && cell.startsWith("=")) {
          const formula = cell.slice(1);
          const updatedFormula = updateCallback(formula);

          // Only update if the formula actually changed
          if (updatedFormula !== formula) {
            sheet.content.set(key, `=${updatedFormula}`);
          }
        }
      });
    });
  }

  getSheetSerialized({
    workbookName,
    sheetName,
  }: {
    workbookName: string;
    sheetName: string;
  }): Map<string, SerializedCellValue> {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new SheetNotFoundError(sheetName);
    }

    return sheet.content;
  }

  /**
   * Add a cell to the grouped indexes
   */
  private addCellToGroups(
    indexes: SheetIndexes,
    rowIndex: number,
    colIndex: number,
    key: string
  ): void {
    // Add to row group (cells in this row, sorted by column)
    let rowGroup = indexes.rowGroups.get(rowIndex);
    if (!rowGroup) {
      rowGroup = [];
      indexes.rowGroups.set(rowIndex, rowGroup);
    }
    const colEntry: IndexEntry = { number: colIndex, key };
    const colInsertIdx = this.findInsertIndex(rowGroup, colIndex);
    rowGroup.splice(colInsertIdx, 0, colEntry);

    // Add to column group (cells in this column, sorted by row)
    let colGroup = indexes.colGroups.get(colIndex);
    if (!colGroup) {
      colGroup = [];
      indexes.colGroups.set(colIndex, colGroup);
    }
    const rowEntry: IndexEntry = { number: rowIndex, key };
    const rowInsertIdx = this.findInsertIndex(colGroup, rowIndex);
    colGroup.splice(rowInsertIdx, 0, rowEntry);

    // Add to sorted flat indexes
    this.insertSorted(indexes.cellsSortedByRow, { number: rowIndex, key });
    this.insertSorted(indexes.cellsSortedByCol, { number: colIndex, key });
  }

  /**
   * Remove a cell from the grouped indexes
   */
  private removeCellFromGroups(
    indexes: SheetIndexes,
    rowIndex: number,
    colIndex: number,
    key: string
  ): void {
    // Remove from row group
    const rowGroup = indexes.rowGroups.get(rowIndex);
    if (rowGroup) {
      const filteredGroup = rowGroup.filter((e) => e.key !== key);
      if (filteredGroup.length === 0) {
        indexes.rowGroups.delete(rowIndex);
      } else {
        indexes.rowGroups.set(rowIndex, filteredGroup);
      }
    }

    // Remove from column group
    const colGroup = indexes.colGroups.get(colIndex);
    if (colGroup) {
      const filteredGroup = colGroup.filter((e) => e.key !== key);
      if (filteredGroup.length === 0) {
        indexes.colGroups.delete(colIndex);
      } else {
        indexes.colGroups.set(colIndex, filteredGroup);
      }
    }

    // Remove from sorted flat indexes
    indexes.cellsSortedByRow = indexes.cellsSortedByRow.filter(
      (item) => item.key !== key
    );
    indexes.cellsSortedByCol = indexes.cellsSortedByCol.filter(
      (item) => item.key !== key
    );
  }

  /**
   * Find insertion index in sorted array
   */
  private findInsertIndex(entries: IndexEntry[], n: number): number {
    return IndexEntryBinarySearch.findInsertionPoint(entries, n);
  }

  /**
   * Inserts an item into a sorted array by number, maintaining sort order.
   * If an item with the same number and key already exists, it won't be added again.
   */
  private insertSorted(array: IndexEntry[], item: IndexEntry): void {
    // Check if item already exists (same number and key)
    const existingIndex = array.findIndex(
      (existing) => existing.number === item.number && existing.key === item.key
    );

    if (existingIndex !== -1) {
      // Item already exists, no need to add it again
      return;
    }

    // Find the insertion point using binary search for efficiency
    const insertionPoint = IndexEntryBinarySearch.findInsertionPoint(
      array,
      item.number
    );

    // Insert at the found position
    array.splice(insertionPoint, 0, item);
  }

  setCellContent(
    address: CellAddress,
    content: SerializedCellValue,
    options?: {
      /**
       * for extra performance, if the sheet is already known, it can be passed in
       */
      sheet?: Sheet;
      /**
       * if the sheet is being built from scratch, we can skip some checks
       */
      buildingFromScratch?: boolean;
    }
  ): void {
    const sheet =
      options?.sheet ||
      this.getSheet({
        sheetName: address.sheetName,
        workbookName: address.workbookName,
      });

    if (!sheet) {
      throw new SheetNotFoundError(address.sheetName);
    }

    const indexes = this.getSheetIndexes({
      workbookName: address.workbookName,
      sheetName: address.sheetName,
    });
    const adr = getCellReference(address);

    if (this.isContentEmpty(content)) {
      if (!options?.buildingFromScratch) {
        sheet.content.delete(adr);
        // Remove from all indexes
        this.removeCellFromGroups(
          indexes,
          address.rowIndex,
          address.colIndex,
          adr
        );
      }
    } else {
      sheet.content.set(adr, content);
      // Add to all indexes
      this.addCellToGroups(indexes, address.rowIndex, address.colIndex, adr);
    }
  }

  /**
   * Set metadata for a cell
   */
  setCellMetadata<TMetadata = unknown>(address: CellAddress, metadata: TMetadata | undefined): void {
    const sheet = this.getSheet({
      workbookName: address.workbookName,
      sheetName: address.sheetName,
    });
    if (!sheet) {
      throw new SheetNotFoundError(address.sheetName);
    }

    const key = getCellReference(address);
    if (metadata === undefined) {
      sheet.metadata.delete(key);
    } else {
      sheet.metadata.set(key, metadata);
    }
  }

  /**
   * Get metadata for a cell
   */
  getCellMetadata<TMetadata = unknown>(address: CellAddress): TMetadata | undefined {
    const sheet = this.getSheet({
      workbookName: address.workbookName,
      sheetName: address.sheetName,
    });
    if (!sheet) {
      return undefined;
    }

    const key = getCellReference(address);
    return sheet.metadata.get(key) as TMetadata | undefined;
  }

  /**
   * Get all metadata for a sheet
   */
  getSheetMetadataSerialized<TMetadata = unknown>(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, TMetadata> {
    const sheet = this.getSheet(opts);
    return sheet?.metadata || new Map();
  }

  /**
   * Set metadata for a sheet
   */
  setSheetMetadata<TSheetMetadata = unknown>(
    opts: { workbookName: string; sheetName: string },
    metadata: TSheetMetadata
  ): void {
    const sheet = this.getSheet(opts);
    if (!sheet) {
      throw new SheetNotFoundError(opts.sheetName);
    }
    sheet.sheetMetadata = metadata;
  }

  /**
   * Get metadata for a sheet
   */
  getSheetMetadata<TSheetMetadata = unknown>(
    opts: { workbookName: string; sheetName: string }
  ): TSheetMetadata | undefined {
    const sheet = this.getSheet(opts);
    if (!sheet) {
      return undefined;
    }
    return sheet.sheetMetadata as TSheetMetadata | undefined;
  }

  /**
   * Set metadata for a workbook
   */
  setWorkbookMetadata<TWorkbookMetadata = unknown>(
    workbookName: string,
    metadata: TWorkbookMetadata
  ): void {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error(`Workbook "${workbookName}" not found`);
    }
    workbook.workbookMetadata = metadata;
  }

  /**
   * Get metadata for a workbook
   */
  getWorkbookMetadata<TWorkbookMetadata = unknown>(
    workbookName: string
  ): TWorkbookMetadata | undefined {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      return undefined;
    }
    return workbook.workbookMetadata as TWorkbookMetadata | undefined;
  }

  /**
   * Replace all content for a sheet (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setSheetContent(
    opts: { sheetName: string; workbookName: string },
    newContent: Map<string, SerializedCellValue>
  ): void {
    const sheet = this.getSheet(opts);
    if (!sheet) {
      throw new SheetNotFoundError(opts.sheetName);
    }

    // Clear existing content without breaking the Map reference
    sheet.content.clear();

    // Clean up indexes for this sheet
    const key = this.getSheetIndexKey(opts.workbookName, opts.sheetName);
    this.sheetIndexes.delete(key);

    // Repopulate with new content
    newContent.forEach((value, key) => {
      this.setCellContent(
        {
          workbookName: opts.workbookName,
          sheetName: opts.sheetName,
          colIndex: parseCellReference(key).colIndex,
          rowIndex: parseCellReference(key).rowIndex,
        },
        value,
        {
          sheet,
          buildingFromScratch: true,
        }
      );
    });
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   * OPTIMIZED: Uses indexes to only process cells that actually exist.
   * ENHANCED: Now supports infinite ranges.
   */
  clearSpreadsheetRange(address: RangeAddress) {
    const sheet = this.getSheet(address);

    if (!sheet) {
      throw new SheetNotFoundError(address.sheetName);
    }

    // Get current sheet content and prepare new content with cleared cells
    const newContent = new Map(sheet.content);
    const newMetadata = new Map(sheet.metadata);

    // Use iterateCellsInRange to only process cells that actually exist
    // This handles both finite and infinite ranges efficiently
    for (const cellAddress of this.iterateCellsInRange(address)) {
      const cellRef = getCellReference(cellAddress);

      // Remove from content and metadata
      newContent.delete(cellRef);
      newMetadata.delete(cellRef);
    }

    // Update content
    this.setSheetContent(address, newContent);
    
    // Update metadata
    sheet.metadata = newMetadata;
  }

  /**
   * Optimized generator to iterate over cells defined in the content within a range
   * Uses indexes to efficiently find and yield only cells that exist within the range
   */
  *iterateCellsInRange(address: RangeAddress): Generator<CellAddress> {
    // First check if the sheet exists
    const sheet = this.getSheet(address);
    if (!sheet) {
      throw new SheetNotFoundError(address.sheetName);
    }

    const indexes = this.getSheetIndexes(address);

    const range = address.range;

    // Use the sorted index to find only rows that actually contain cells
    // This avoids iterating through empty rows regardless of finite/infinite bounds

    if (range.end.row.type === "number") {
      // Finite bounds: Use binary search to find the range of cells to check
      const startIndex = IndexEntryBinarySearch.findFirstGreaterOrEqual(
        indexes.cellsSortedByRow,
        range.start.row
      );

      if (startIndex === -1) return; // No cells at or after start row

      // Process cells from startIndex until we exceed the end row
      for (let i = startIndex; i < indexes.cellsSortedByRow.length; i++) {
        const cellEntry = indexes.cellsSortedByRow[i];
        if (!cellEntry) continue;

        const parsed = parseCellReference(cellEntry.key);

        // Stop if we've gone beyond the row range
        if (parsed.rowIndex > range.end.row.value) break;

        // Check if cell is within column bounds
        if (parsed.colIndex < range.start.col) continue;

        if (
          range.end.col.type === "number" &&
          parsed.colIndex > range.end.col.value
        ) {
          continue; // Skip this cell but keep checking others in different rows
        }

        yield {
          rowIndex: parsed.rowIndex,
          colIndex: parsed.colIndex,
          sheetName: address.sheetName,
          workbookName: address.workbookName,
        };
      }
    } else {
      // Infinite row bounds: Use binary search to find starting point
      const startIndex = IndexEntryBinarySearch.findFirstGreaterOrEqual(
        indexes.cellsSortedByRow,
        range.start.row
      );

      if (startIndex === -1) return; // No cells at or after start row

      // Process all cells from startIndex to end
      for (let i = startIndex; i < indexes.cellsSortedByRow.length; i++) {
        const cellEntry = indexes.cellsSortedByRow[i];
        if (!cellEntry) continue;

        const parsed = parseCellReference(cellEntry.key);

        // Check if cell is within column bounds
        if (parsed.colIndex < range.start.col) continue;

        if (
          range.end.col.type === "number" &&
          parsed.colIndex > range.end.col.value
        ) {
          continue; // Skip this cell but keep checking others in different rows
        }

        yield {
          rowIndex: parsed.rowIndex,
          colIndex: parsed.colIndex,
          sheetName: address.sheetName,
          workbookName: address.workbookName,
        };
      }
    }
  }

  getCellsInRange(address: RangeAddress): CellAddress[] {
    return Array.from(this.iterateCellsInRange(address));
  }

  public getCellContent(cellAddress: CellAddress): SerializedCellValue {
    const sheet = this.getSheet(cellAddress);
    if (!sheet) {
      throw new SheetNotFoundError(cellAddress.sheetName);
    }
    return sheet.content.get(getCellReference(cellAddress));
  }

  public getSerializedCellValue(cellAddress: CellAddress): SerializedCellValue {
    const sheet = this.getSheet(cellAddress);
    if (!sheet) {
      throw new SheetNotFoundError(cellAddress.sheetName);
    }
    return normalizeSerializedCellValue(
      sheet.content.get(getCellReference(cellAddress))
    );
  }

  public isCellEmpty(cellAddress: CellAddress): boolean {
    const content = this.getCellContent(cellAddress);
    return (
      content === undefined || (typeof content === "string" && content === "")
    );
  }
  public isFormulaCell(cellAddress: CellAddress): boolean {
    const content = this.getCellContent(cellAddress);
    return typeof content === "string" && content.startsWith("=");
  }

  /**
   * Build evaluation order for a range
   * Delegates to the buildRangeEvalOrder function
   */
  public buildRangeEvalOrder(
    lookupOrder: "row-major" | "col-major",
    lookupRange: RangeAddress
  ) {
    // Import and call the function
    return buildRangeEvalOrder.call(this, lookupOrder, lookupRange);
  }
}
