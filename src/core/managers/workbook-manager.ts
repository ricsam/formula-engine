import type {
  CellAddress,
  FiniteSpreadsheetRange,
  SerializedCellValue,
  Sheet,
  SpreadsheetRange,
  Workbook,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";

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
  private getSheetIndexes(
    workbookName: string,
    sheetName: string
  ): SheetIndexes {
    const key = this.getSheetIndexKey(workbookName, sheetName);
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
        const indexes = this.getSheetIndexes(workbookName, sheet.name);
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
    this.getSheetIndexes(workbookName, sheetName);

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
    let left = 0;
    let right = entries.length;

    while (left < right) {
      const mid = Math.floor((left + right) / 2);
      if (entries[mid]!.number < n) {
        left = mid + 1;
      } else {
        right = mid;
      }
    }

    return left;
  }

  /**
   * Inserts an item into a sorted array by number, maintaining sort order.
   * If an item with the same number and key already exists, it won't be added again.
   */
  private insertSorted(
    array: { number: number; key: string }[],
    item: { number: number; key: string }
  ): void {
    // Check if item already exists (same number and key)
    const existingIndex = array.findIndex(
      (existing) => existing.number === item.number && existing.key === item.key
    );

    if (existingIndex !== -1) {
      // Item already exists, no need to add it again
      return;
    }

    // Find the insertion point using binary search for efficiency
    let left = 0;
    let right = array.length;

    while (left < right) {
      const mid = Math.floor((left + right) / 2);
      if (array[mid]!.number < item.number) {
        left = mid + 1;
      } else {
        right = mid;
      }
    }

    // Insert at the found position
    array.splice(left, 0, item);
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

    const indexes = this.getSheetIndexes(
      address.workbookName,
      address.sheetName
    );
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
   */
  clearSpreadsheetRange(
    opts: { sheetName: string; workbookName: string },
    range: SpreadsheetRange
  ) {
    // Check if range has infinite ends - not supported for now
    if (
      range.end.col.type === "infinity" ||
      range.end.row.type === "infinity"
    ) {
      throw new Error("Clearing infinite ranges is not supported");
    }

    const sheet = this.getSheet(opts);

    if (!sheet) {
      throw new Error("Sheet not found");
    }

    // Get current sheet content and prepare new content with cleared cells
    const newContent = new Map(sheet.content);

    // Convert to finite range (throws error if infinite)
    const finiteRange = this.toFiniteRange(range);

    // Iterate through all cells in the range and clear them
    const startCol = finiteRange.start.col;
    const startRow = finiteRange.start.row;
    const endCol = finiteRange.end.col;
    const endRow = finiteRange.end.row;

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cellRef = getCellReference({
          colIndex: col,
          rowIndex: row,
        });
        newContent.delete(cellRef);
      }
    }

    this.setSheetContent(opts, newContent);
  }

  /**
   * Get all cells in a specific row (pre-computed from grouped index)
   */
  private getCellsInRow(
    workbookName: string,
    sheetName: string,
    row: number
  ): IndexEntry[] {
    const indexes = this.getSheetIndexes(workbookName, sheetName);
    // Direct O(1) lookup from Map
    return indexes.rowGroups.get(row) ?? [];
  }

  /**
   * Get all cells in a specific column (pre-computed from grouped index)
   */
  private getCellsInColumn(
    workbookName: string,
    sheetName: string,
    column: number
  ): IndexEntry[] {
    const indexes = this.getSheetIndexes(workbookName, sheetName);
    // Direct O(1) lookup from Map
    return indexes.colGroups.get(column) ?? [];
  }

  /**
   * Get frontier candidates that might spill into the range
   * A frontier candidate is a formula cell located above or to the left of the range
   * with no blocking cells between it and the range
   */
  getFrontierCandidates(
    range: SpreadsheetRange,
    opts: {
      sheetName: string;
      workbookName: string;
    }
  ): CellAddress[] {
    const candidates = new Set<string>();

    const sheet = this.getSheet({
      sheetName: opts.sheetName,
      workbookName: opts.workbookName,
    });

    if (!sheet) {
      return [];
    }

    // Get columns that intersect with the range
    const colsToCheck = this.getColumnsInRange(
      range,
      opts.workbookName,
      opts.sheetName
    );

    // For each column, find the nearest formula cell above the range
    for (const col of colsToCheck) {
      const cellsInCol = this.getCellsInColumn(
        opts.workbookName,
        opts.sheetName,
        col
      );
      const nearestAbove = this.findNearestFormulaAbove(
        cellsInCol,
        range.start.row,
        sheet
      );
      if (nearestAbove) {
        candidates.add(nearestAbove);
      }
    }

    // Get rows that intersect with the range
    const rowsToCheck = this.getRowsInRange(
      range,
      opts.workbookName,
      opts.sheetName
    );

    // For each row, find the nearest formula cell to the left of the range
    for (const row of rowsToCheck) {
      const cellsInRow = this.getCellsInRow(
        opts.workbookName,
        opts.sheetName,
        row
      );
      const nearestLeft = this.findNearestFormulaLeft(
        cellsInRow,
        range.start.col,
        sheet
      );
      if (nearestLeft) {
        candidates.add(nearestLeft);
      }
    }

    return Array.from(candidates).map((key) => ({
      ...parseCellReference(key),
      sheetName: opts.sheetName,
      workbookName: opts.workbookName,
    }));
  }

  /**
   * Find the nearest formula cell above the given row in a column
   */
  private findNearestFormulaAbove(
    cellsInCol: IndexEntry[],
    beforeRow: number,
    sheet: Sheet
  ): string | null {
    // Search from end to beginning (cells are sorted by row)
    for (let i = cellsInCol.length - 1; i >= 0; i--) {
      const cell = cellsInCol[i];
      if (!cell || cell.number >= beforeRow) continue;

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
    // Search from end to beginning (cells are sorted by column)
    for (let i = cellsInRow.length - 1; i >= 0; i--) {
      const cell = cellsInRow[i];
      if (!cell || cell.number >= beforeCol) continue;

      const content = sheet.content.get(cell.key);
      if (typeof content === "string" && content.startsWith("=")) {
        return cell.key;
      }
    }

    return null;
  }

  /**
   * Get unique numbers (rows or columns) that intersect with a range dimension
   */
  private getNumbersInRangeDimension(
    list: { number: number; key: string }[],
    startNum: number,
    endDimension: { type: "number"; value: number } | { type: "infinity" }
  ): number[] {
    const numbers = new Set<number>();

    // Always include the starting number
    numbers.add(startNum);

    // Since the list is sorted, we can use binary search to find the start position
    const startIdx = this.findFirstGreaterOrEqual(list, startNum);
    if (startIdx === -1) return [startNum];

    // Iterate from the start position
    for (let i = startIdx; i < list.length; i++) {
      const num = list[i]!.number;

      // Check if we've gone past the end of a finite range
      if (endDimension.type === "number" && num > endDimension.value) {
        break;
      }

      numbers.add(num);
    }

    return Array.from(numbers);
  }

  /**
   * Binary search to find the first element >= target
   */
  private findFirstGreaterOrEqual(
    list: { number: number; key: string }[],
    target: number
  ): number {
    if (list.length === 0) return -1;

    let left = 0;
    let right = list.length - 1;
    let result = -1;

    while (left <= right) {
      const mid = Math.floor((left + right) / 2);
      if (list[mid]!.number >= target) {
        result = mid;
        right = mid - 1;
      } else {
        left = mid + 1;
      }
    }

    return result;
  }

  /**
   * Get columns that intersect with the range
   */
  private getColumnsInRange(
    range: SpreadsheetRange,
    workbookName: string,
    sheetName: string
  ): number[] {
    const indexes = this.getSheetIndexes(workbookName, sheetName);
    return this.getNumbersInRangeDimension(
      indexes.cellsSortedByCol,
      range.start.col,
      range.end.col
    );
  }

  /**
   * Get rows that intersect with the range
   */
  private getRowsInRange(
    range: SpreadsheetRange,
    workbookName: string,
    sheetName: string
  ): number[] {
    const indexes = this.getSheetIndexes(workbookName, sheetName);
    return this.getNumbersInRangeDimension(
      indexes.cellsSortedByRow,
      range.start.row,
      range.end.row
    );
  }
}
