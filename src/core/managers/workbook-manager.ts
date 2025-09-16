import type {
  CellAddress,
  FiniteSpreadsheetRange,
  SerializedCellValue,
  Sheet,
  SpreadsheetRange,
  Workbook,
} from "../types";
import { getCellReference, parseCellReference } from "../utils";

export class WorkbookManager {
  private workbooks: Map<string, Workbook> = new Map();

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
    if (!this.workbooks.has(workbookName)) {
      throw new Error("Workbook not found");
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
    this.workbooks.set(opts.newWorkbookName, workbook);
    this.workbooks.delete(opts.workbookName);
    workbook.name = opts.newWorkbookName;
  }

  resetWorkbooks(workbooks: Map<string, Workbook>): void {
    this.workbooks.clear();
    workbooks.forEach((workbook, workbookName) => {
      this.workbooks.set(workbookName, workbook);
      workbook.sheets.forEach((sheet) => {
        sheet.rows = [];
        sheet.cols = [];
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
      rows: [],
      cols: [],
    };

    if (workbook.sheets.has(sheet.name)) {
      throw new Error("Sheet already exists");
    }

    workbook.sheets.set(sheetName, sheet);

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

    const adr = getCellReference(address);
    if (this.isContentEmpty(content)) {
      if (!options?.buildingFromScratch) {
        sheet.content.delete(adr);
        // Remove from rows and cols arrays when content is deleted
        sheet.rows = sheet.rows.filter((row) => row.key !== adr);
        sheet.cols = sheet.cols.filter((col) => col.key !== adr);
      }
    } else {
      sheet.content.set(adr, content);
      this.insertSorted(sheet.rows, { number: address.rowIndex, key: adr });
      this.insertSorted(sheet.cols, { number: address.colIndex, key: adr });
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
   * Get all cells in specific columns within the sheet
   */
  private getCellsInColumns(
    sheet: Sheet,
    columns: number[]
  ): { col: number; cells: { row: number; key: string }[] }[] {
    const result: { col: number; cells: { row: number; key: string }[] }[] = [];

    for (const col of columns) {
      const cells: { row: number; key: string }[] = [];

      // Use binary search to find the range of cells in this column
      const startIdx = this.findFirstInList(sheet.cols, col);
      if (startIdx !== -1) {
        let idx = startIdx;
        while (idx < sheet.cols.length && sheet.cols[idx]!.number === col) {
          const cellRef = sheet.cols[idx]!.key;
          const { rowIndex } = parseCellReference(cellRef);
          cells.push({ row: rowIndex, key: cellRef });
          idx++;
        }
      }

      if (cells.length > 0) {
        // Sort cells by row number
        cells.sort((a, b) => a.row - b.row);
        result.push({ col, cells });
      }
    }

    return result;
  }

  /**
   * Get all cells in specific rows within the sheet
   */
  private getCellsInRows(
    sheet: Sheet,
    rows: number[]
  ): { row: number; cells: { col: number; key: string }[] }[] {
    const result: { row: number; cells: { col: number; key: string }[] }[] = [];

    for (const row of rows) {
      const cells: { col: number; key: string }[] = [];

      // Use binary search to find the range of cells in this row
      const startIdx = this.findFirstInList(sheet.rows, row);
      if (startIdx !== -1) {
        let idx = startIdx;
        while (idx < sheet.rows.length && sheet.rows[idx]!.number === row) {
          const cellRef = sheet.rows[idx]!.key;
          const { colIndex } = parseCellReference(cellRef);
          cells.push({ col: colIndex, key: cellRef });
          idx++;
        }
      }

      if (cells.length > 0) {
        // Sort cells by column number
        cells.sort((a, b) => a.col - b.col);
        result.push({ row, cells });
      }
    }

    return result;
  }

  /**
   * Binary search to find the first cell in a specific row (or column)
   */
  private findFirstInList(
    list: { number: number; key: string }[],
    targetNum: number
  ): number {
    let left = 0;
    let right = list.length - 1;
    let result = -1;

    while (left <= right) {
      const mid = Math.floor((left + right) / 2);
      if (list[mid]!.number === targetNum) {
        result = mid;
        // Continue searching to the left to find the first occurrence
        right = mid - 1;
      } else if (list[mid]!.number < targetNum) {
        left = mid + 1;
      } else {
        right = mid - 1;
      }
    }

    return result;
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
    const colsToCheck = this.getColumnsInRange(range, sheet);

    // For each column, find the nearest formula cell above the range
    for (const col of colsToCheck) {
      const cellsInCol = this.getCellsInColumns(sheet, [col])[0];
      if (cellsInCol) {
        const nearestAbove = this.findNearestFormulaAbove(
          cellsInCol.cells,
          range.start.row,
          sheet
        );
        if (nearestAbove) {
          candidates.add(nearestAbove);
        }
      }
    }

    // Get rows that intersect with the range
    const rowsToCheck = this.getRowsInRange(range, sheet);

    // For each row, find the nearest formula cell to the left of the range
    for (const row of rowsToCheck) {
      const cellsInRow = this.getCellsInRows(sheet, [row])[0];
      if (cellsInRow) {
        const nearestLeft = this.findNearestFormulaLeft(
          cellsInRow.cells,
          range.start.col,
          sheet
        );
        if (nearestLeft) {
          candidates.add(nearestLeft);
        }
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
    cellsInCol: { row: number; key: string }[],
    beforeRow: number,
    sheet: Sheet
  ): string | null {
    // Search from bottom to top (reverse order since cells are sorted by row)
    for (let i = cellsInCol.length - 1; i >= 0; i--) {
      const cell = cellsInCol[i];
      if (!cell || cell.row >= beforeRow) continue;

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
    cellsInRow: { col: number; key: string }[],
    beforeCol: number,
    sheet: Sheet
  ): string | null {
    // Search from right to left (reverse order since cells are sorted by col)
    for (let i = cellsInRow.length - 1; i >= 0; i--) {
      const cell = cellsInRow[i];
      if (!cell || cell.col >= beforeCol) continue;

      const content = sheet.content.get(cell.key);
      if (typeof content === "string" && content.startsWith("=")) {
        return cell.key;
      }
    }

    return null;
  }

  /**
   * Get columns that intersect with the range
   */
  private getColumnsInRange(range: SpreadsheetRange, sheet: Sheet): number[] {
    const cols = new Set<number>();

    // Always include the starting column
    cols.add(range.start.col);

    // Find all unique columns in the sheet that are within the range
    for (const { number: colIndex } of sheet.cols) {
      if (colIndex >= range.start.col) {
        if (
          range.end.col.type === "number" &&
          colIndex <= range.end.col.value
        ) {
          cols.add(colIndex);
        } else if (range.end.col.type === "infinity") {
          cols.add(colIndex);
        }
      }
    }

    return Array.from(cols).sort((a, b) => a - b);
  }

  /**
   * Get rows that intersect with the range
   */
  private getRowsInRange(range: SpreadsheetRange, sheet: Sheet): number[] {
    const rows = new Set<number>();

    // Always include the starting row
    rows.add(range.start.row);

    // Find all unique rows in the sheet that are within the range
    for (const { number: rowIndex } of sheet.rows) {
      if (rowIndex >= range.start.row) {
        if (
          range.end.row.type === "number" &&
          rowIndex <= range.end.row.value
        ) {
          rows.add(rowIndex);
        } else if (range.end.row.type === "infinity") {
          rows.add(rowIndex);
        }
      }
    }

    return Array.from(rows).sort((a, b) => a - b);
  }
}
