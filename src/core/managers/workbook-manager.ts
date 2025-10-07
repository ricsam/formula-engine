import type {
  CellAddress,
  FiniteSpreadsheetRange,
  LocalCellAddress,
  SerializedCellValue,
  Sheet,
  SpreadsheetRange,
  Workbook,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";

import type { RangeAddress } from "src/core/types";
import { buildRangeEvalOrder } from "./range-eval-order-builder";

interface IndexEntry {
  number: number;
  key: string;
}

interface SheetIndexes {
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
class IndexEntryBinarySearch {
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
  private getSheetIndexes(opts: {
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
      throw new Error("Workbook not found");
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
    });
  }

  removeWorkbook(workbookName: string): void {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new Error("Workbook not found");
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
      throw new Error("Workbook not found");
    }
    const sheet = {
      name: sheetName,
      index: workbook.sheets.size,
      content: new Map(),
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
      throw new Error("Workbook not found");
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
      throw new Error("Workbook not found");
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
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

  updateAllFormulas(updateCallback: (formula: string) => string): void {
    const update = (map: Map<string, Sheet>) => {
      map.forEach((sheet) => {
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
    };

    this.workbooks.forEach((workbook) => {
      update(workbook.sheets);
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
      throw new Error("Workbook not found");
    }
    const sheet = workbook.sheets.get(sheetName);
    if (!sheet) {
      throw new Error("Sheet not found");
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
      throw new Error("Sheet not found");
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
   * Replace all content for a sheet (safely, without breaking references)
   * This method clears the existing Map and repopulates it rather than replacing the Map reference
   */
  setSheetContent(
    opts: { sheetName: string; workbookName: string },
    newContent: Map<string, SerializedCellValue>
  ): void {
    const sheet = this.getSheet(opts);
    if (!sheet) {
      throw new Error("Sheet not found");
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
   * Converts a SpreadsheetRange to FiniteSpreadsheetRange, throwing an error if infinite
   */
  private toFiniteRange(range: SpreadsheetRange): FiniteSpreadsheetRange {
    if (
      range.end.col.type === "infinity" ||
      range.end.row.type === "infinity"
    ) {
      throw new Error("Clearing infinite ranges is not supported");
    }

    return {
      start: range.start,
      end: {
        col: range.end.col.value,
        row: range.end.row.value,
      },
    };
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   * OPTIMIZED: Uses indexes to only process cells that actually exist.
   * ENHANCED: Now supports infinite ranges.
   */
  clearSpreadsheetRange(address: RangeAddress) {
    const sheet = this.getSheet(address);

    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Get current sheet content and prepare new content with cleared cells
    const newContent = new Map(sheet.content);

    // Use iterateCellsInRange to only process cells that actually exist
    // This handles both finite and infinite ranges efficiently
    for (const cellAddress of this.iterateCellsInRange(address)) {
      const cellRef = getCellReference(cellAddress);

      // Remove from content
      newContent.delete(cellRef);
    }

    // setSheetContent will rebuild indexes from scratch, so no need to manually update them
    this.setSheetContent(address, newContent);
  }

  /**
   * Get all cells in a specific row (pre-computed from grouped index)
   */
  private getCellsInRow(
    opts: { workbookName: string; sheetName: string },
    row: number
  ): IndexEntry[] {
    const indexes = this.getSheetIndexes(opts);
    // Direct O(1) lookup from Map
    return indexes.rowGroups.get(row) ?? [];
  }

  /**
   * Optimized generator to iterate over cells defined in the content within a range
   * Uses indexes to efficiently find and yield only cells that exist within the range
   */
  *iterateCellsInRange(address: RangeAddress): Generator<CellAddress> {
    // First check if the sheet exists
    const sheet = this.getSheet(address);
    if (!sheet) {
      throw new Error("Sheet not found");
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

  /**
   * Get all cells in a specific column (pre-computed from grouped index)
   */
  private getCellsInColumn(
    opts: { workbookName: string; sheetName: string },
    column: number
  ): IndexEntry[] {
    const indexes = this.getSheetIndexes(opts);
    // Direct O(1) lookup from Map
    return indexes.colGroups.get(column) ?? [];
  }

  /**
   * Generator that yields frontier candidates that might spill into the range
   * A frontier candidate is a formula cell that could potentially spill into the range.
   * This includes cells in three regions:
   * 1. Above the range (same columns, rows before) - nearest formula cell per column
   * 2. To the left of the range (same rows, columns before) - nearest formula cell per row
   * 3. Top-left quadrant (both above AND to the left of the range) - all formula cells
   *    with clear paths to the range
   *
   * For regions 1 & 2, we only yield the nearest formula cell in each direction.
   * For region 3, we yield all formula cells that have no blocking formula cells
   * between them and the range (both horizontally and vertically).
   */
  *iterateFrontierCandidates(range: RangeAddress): Generator<CellAddress> {
    const candidateKeys = new Set<string>();

    const sheet = this.getSheet({
      sheetName: range.sheetName,
      workbookName: range.workbookName,
    });

    if (!sheet) {
      return;
    }

    // 1. Get columns that intersect with the range (cells above)
    const colsToCheck = this.getColumnsInRange(range);

    // For each column, find the nearest formula cell above the range
    for (const col of colsToCheck) {
      const cellsInCol = this.getCellsInColumn(range, col);
      const nearestAbove = this.findNearestFormulaAbove(
        cellsInCol,
        range.range.start.row,
        sheet
      );
      if (nearestAbove && !candidateKeys.has(nearestAbove)) {
        candidateKeys.add(nearestAbove);
        yield {
          ...parseCellReference(nearestAbove),
          sheetName: range.sheetName,
          workbookName: range.workbookName,
        };
      }
    }

    // 2. Get rows that intersect with the range (cells to the left)
    const rowsToCheck = this.getRowsInRange(range);

    // For each row, find the nearest formula cell to the left of the range
    for (const row of rowsToCheck) {
      const cellsInRow = this.getCellsInRow(range, row);
      const nearestLeft = this.findNearestFormulaLeft(
        cellsInRow,
        range.range.start.col,
        sheet
      );
      if (nearestLeft && !candidateKeys.has(nearestLeft)) {
        candidateKeys.add(nearestLeft);
        yield {
          ...parseCellReference(nearestLeft),
          sheetName: range.sheetName,
          workbookName: range.workbookName,
        };
      }
    }

    // 3. Check the top-left quadrant (cells both above AND to the left)
    // These are cells that might spill diagonally into the range
    for (const candidate of this.findTopLeftQuadrantCandidates(
      range,
      sheet,
      candidateKeys
    )) {
      candidateKeys.add(getCellReference(candidate));
      yield candidate;
    }
  }

  /**
   * Get frontier candidates that might spill into the range
   * A frontier candidate is a formula cell that could potentially spill into the range.
   * This includes cells in three regions:
   * 1. Above the range (same columns, rows before)
   * 2. To the left of the range (same rows, columns before)
   * 3. Top-left quadrant (both above AND to the left of the range)
   */
  getFrontierCandidates(range: RangeAddress): CellAddress[] {
    return Array.from(this.iterateFrontierCandidates(range));
  }

  /**
   * Find formula cells in the top-left quadrant that could spill into the range
   * This checks cells that are both above AND to the left of the range
   */
  *findTopLeftQuadrantCandidates(
    address: RangeAddress,
    sheet: Sheet,
    existingCandidates: Set<string>
  ): Generator<CellAddress> {
    // Only process if there's actually a top-left quadrant
    if (address.range.start.row === 0 || address.range.start.col === 0) {
      return;
    }

    const indexes = this.getSheetIndexes(address);

    // Use sorted index to find only cells that exist above and to the left of the range
    // This avoids iterating through empty rows/columns
    const candidateCells = indexes.cellsSortedByRow.filter((cell) => {
      const parsed = parseCellReference(cell.key);
      return (
        parsed.rowIndex < address.range.start.row &&
        parsed.colIndex < address.range.start.col
      );
    });

    // Process each candidate cell
    for (const cell of candidateCells) {
      // Skip if already found
      if (existingCandidates.has(cell.key)) continue;

      const content = sheet.content.get(cell.key);
      if (typeof content === "string" && content.startsWith("=")) {
        const parsed = parseCellReference(cell.key);
        const row = parsed.rowIndex;
        const col = parsed.colIndex;

        // Check if there's a clear path to the range
        // A cell at (row, col) can spill to the range if:
        // 1. No formulas exist between (row, col+1) and (row, range.start.col-1)
        // 2. No formulas exist between (row+1, col) and (range.start.row-1, col)

        let hasBlockingCell = false;

        // Check horizontal path (same row, columns between this cell and range)
        const cellsInRow = this.getCellsInRow(address, row);
        for (const blockingCandidate of cellsInRow) {
          if (
            blockingCandidate.number > col &&
            blockingCandidate.number < address.range.start.col
          ) {
            const blockingContent = sheet.content.get(blockingCandidate.key);
            if (
              typeof blockingContent === "string" &&
              blockingContent.startsWith("=")
            ) {
              hasBlockingCell = true;
              break;
            }
          }
        }

        if (!hasBlockingCell) {
          // Check vertical path (same column, rows between this cell and range)
          const colCells = this.getCellsInColumn(address, col);
          for (const blockingCandidate of colCells) {
            if (
              blockingCandidate.number > row &&
              blockingCandidate.number < address.range.start.row
            ) {
              const blockingContent = sheet.content.get(blockingCandidate.key);
              if (
                typeof blockingContent === "string" &&
                blockingContent.startsWith("=")
              ) {
                hasBlockingCell = true;
                break;
              }
            }
          }
        }

        if (!hasBlockingCell) {
          yield {
            ...parsed,
            sheetName: address.sheetName,
            workbookName: address.workbookName,
          };
        }
      }
    }
  }

  /**
   * Find the nearest formula cell above the given row in a column
   */
  private findNearestFormulaAbove(
    cellsInCol: IndexEntry[],
    beforeRow: number,
    sheet: Sheet
  ): string | null {
    if (cellsInCol.length === 0) return null;

    // Binary search to find the rightmost position where we could insert beforeRow
    const insertionPoint = IndexEntryBinarySearch.findRightmostInsertionPoint(
      cellsInCol,
      beforeRow
    );

    // Now search backwards from the insertion point to find the nearest formula
    for (let i = insertionPoint - 1; i >= 0; i--) {
      const cell = cellsInCol[i];
      if (!cell) continue;

      const content = sheet.content.get(cell.key);
      if (typeof content === "string" && content.startsWith("=")) {
        return cell.key;
      }
    }

    return null;
  }

  /**
   * Find the nearest formula cell to the left of the given column in a row
   */
  private findNearestFormulaLeft(
    cellsInRow: IndexEntry[],
    beforeCol: number,
    sheet: Sheet
  ): string | null {
    if (cellsInRow.length === 0) return null;

    // Binary search to find the rightmost position where we could insert beforeCol
    const insertionPoint = IndexEntryBinarySearch.findRightmostInsertionPoint(
      cellsInRow,
      beforeCol
    );

    // Now search backwards from the insertion point to find the nearest formula
    for (let i = insertionPoint - 1; i >= 0; i--) {
      const cell = cellsInRow[i];
      if (!cell) continue;

      const content = sheet.content.get(cell.key);
      if (typeof content === "string" && content.startsWith("=")) {
        return cell.key;
      }
    }

    return null;
  }

  /**
   * Get unique dimensions (rows or columns) that actually contain cells and intersect with a range
   * OPTIMIZED: Only returns dimensions that have actual cell data, not all numbers in the range
   */
  private getActualDimensionsInRange(
    list: IndexEntry[],
    startNum: number,
    endDimension: { type: "number"; value: number } | { type: "infinity" }
  ): number[] {
    const dimensions = new Set<number>();

    // Use binary search to find the first cell >= startNum
    const startIdx = IndexEntryBinarySearch.findFirstGreaterOrEqual(
      list,
      startNum
    );
    if (startIdx === -1) {
      // No cells at or after startNum
      return [];
    }

    // Only iterate through cells that actually exist and are within bounds
    for (let i = startIdx; i < list.length; i++) {
      const entry = list[i];
      if (!entry) continue;

      const num = entry.number;

      // Check if we've gone past the end of a finite range
      if (endDimension.type === "number" && num > endDimension.value) {
        break;
      }

      dimensions.add(num);
    }

    return Array.from(dimensions).sort((a, b) => a - b);
  }

  /**
   * Binary search to find the first element >= target
   */
  private findFirstGreaterOrEqual(list: IndexEntry[], target: number): number {
    return IndexEntryBinarySearch.findFirstGreaterOrEqual(list, target);
  }

  /**
   * Get columns that actually contain cells and intersect with the range
   * OPTIMIZED: Only returns columns that have actual cell data within the range
   */
  private getColumnsInRange(range: RangeAddress): number[] {
    const indexes = this.getSheetIndexes(range);
    return this.getActualDimensionsInRange(
      indexes.cellsSortedByCol,
      range.range.start.col,
      range.range.end.col
    );
  }

  /**
   * Get rows that actually contain cells and intersect with the range
   * OPTIMIZED: Only returns rows that have actual cell data within the range
   */
  private getRowsInRange(range: RangeAddress): number[] {
    const indexes = this.getSheetIndexes(range);
    return this.getActualDimensionsInRange(
      indexes.cellsSortedByRow,
      range.range.start.row,
      range.range.end.row
    );
  }

  public getCellContent(cellAddress: CellAddress): SerializedCellValue {
    const sheet = this.getSheet(cellAddress);
    if (!sheet) {
      throw new Error("Sheet not found");
    }
    return sheet.content.get(getCellReference(cellAddress));
  }

  public isCellEmpty(cellAddress: CellAddress): boolean {
    const content = this.getCellContent(cellAddress);
    return (
      content === undefined || (typeof content === "string" && content === "")
    );
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
