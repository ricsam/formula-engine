import {
  DEFAULT_SEARCH_MAX_RESULTS,
  FormulaError,
  type CellAddress,
  type FiniteSpreadsheetRange,
  type LocalCellAddress,
  type ReplaceChange,
  type ReplaceTarget,
  type SearchMatch,
  type SearchOptions,
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

type SearchScopeSheet = {
  workbookName: string;
  sheet: Sheet;
};

type SearchScopeStringCell = {
  workbookName: string;
  sheetName: string;
  cellReference: string;
  cellContent: string;
  rowIndex: number;
  colIndex: number;
};

type StringCellMatch = Pick<
  SearchMatch,
  "occurrenceIndex" | "startIndex" | "endIndexExclusive" | "matchedText"
>;

type PreparedReplace = {
  address: CellAddress;
  beforeContent: string;
  afterContent: string;
  change: ReplaceChange;
};

type PreparedCellReplaceAll = {
  address: CellAddress;
  beforeContent: string;
  afterContent: string;
  changes: ReplaceChange[];
};

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

  getOrderedSheets(workbookName: string): Sheet[] {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }

    return Array.from(workbook.sheets.entries())
      .map(([name, sheet], insertionOrder) => ({
        name,
        sheet,
        insertionOrder,
      }))
      .sort((left, right) => {
        if (left.sheet.index !== right.sheet.index) {
          return left.sheet.index - right.sheet.index;
        }
        return left.insertionOrder - right.insertionOrder;
      })
      .map(({ sheet }) => sheet);
  }

  getOrderedSheetNames(workbookName: string): string[] {
    return this.getOrderedSheets(workbookName).map((sheet) => sheet.name);
  }

  getNextAvailableSheetName(
    workbookName: string,
    baseName: string = "Sheet"
  ): string {
    const workbook = this.workbooks.get(workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(workbookName);
    }

    let suffix = 1;
    while (workbook.sheets.has(`${baseName}${suffix}`)) {
      suffix++;
    }

    return `${baseName}${suffix}`;
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

    let nextSheetIndex = -1;
    for (const existingSheet of workbook.sheets.values()) {
      nextSheetIndex = Math.max(nextSheetIndex, existingSheet.index);
    }

    const sheet = {
      name: sheetName,
      index: nextSheetIndex + 1,
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

    // Rebuild the map so the renamed sheet keeps its original position
    const renamedSheets = new Map<string, Sheet>();
    for (const [existingSheetName, existingSheet] of workbook.sheets.entries()) {
      if (existingSheetName === sheetName) {
        renamedSheets.set(newSheetName, sheet);
      } else {
        renamedSheets.set(existingSheetName, existingSheet);
      }
    }
    workbook.sheets.clear();
    for (const [existingSheetName, existingSheet] of renamedSheets.entries()) {
      workbook.sheets.set(existingSheetName, existingSheet);
    }

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

  updateAllFormulas(
    updateCallback: (formula: string, address: CellAddress) => string
  ): CellAddress[] {
    const changed: CellAddress[] = [];

    const update = (workbookName: string, map: Map<string, Sheet>) => {
      map.forEach((sheet, sheetName) => {
        sheet.content.forEach((cell, key) => {
          if (typeof cell === "string" && cell.startsWith("=")) {
            const formula = cell.slice(1);
            const { colIndex, rowIndex } = parseCellReference(key);
            const address = {
              workbookName,
              sheetName,
              colIndex,
              rowIndex,
            };
            const updatedFormula = updateCallback(formula, address);

            // Only update if the formula actually changed
            if (updatedFormula !== formula) {
              sheet.content.set(key, `=${updatedFormula}`);
              changed.push(address);
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

  private resolveSearchScope(
    options?: SearchOptions
  ): SearchScopeSheet[] {
    if (options?.sheetName && !options.workbookName) {
      throw new Error("workbookName is required when sheetName is provided");
    }

    if (!options?.workbookName) {
      const scopedSheets: SearchScopeSheet[] = [];
      for (const workbookName of this.workbooks.keys()) {
        for (const sheet of this.getOrderedSheets(workbookName)) {
          scopedSheets.push({ workbookName, sheet });
        }
      }
      return scopedSheets;
    }

    const workbookName = options.workbookName;
    if (!this.workbooks.has(workbookName)) {
      throw new WorkbookNotFoundError(workbookName);
    }

    if (!options.sheetName) {
      return this.getOrderedSheets(workbookName).map((sheet) => ({
        workbookName,
        sheet,
      }));
    }

    const sheet = this.getSheet({
      workbookName,
      sheetName: options.sheetName,
    });

    if (!sheet) {
      throw new SheetNotFoundError(options.sheetName);
    }

    return [{ workbookName, sheet }];
  }

  private getStringContentKind(cellContent: string): SearchMatch["contentKind"] {
    return cellContent.startsWith("=") ? "formula" : "text";
  }

  private findMatchesInString(
    cellContent: string,
    query: string,
    caseSensitive: boolean,
    maxMatches: number = Number.POSITIVE_INFINITY
  ): StringCellMatch[] {
    if (query.length === 0 || maxMatches <= 0) {
      return [];
    }

    const normalizedContent = caseSensitive
      ? cellContent
      : cellContent.toLowerCase();
    const normalizedQuery = caseSensitive ? query : query.toLowerCase();
    const matches: StringCellMatch[] = [];
    let searchFromIndex = 0;

    while (
      matches.length < maxMatches &&
      searchFromIndex <= normalizedContent.length - normalizedQuery.length
    ) {
      const startIndex = normalizedContent.indexOf(normalizedQuery, searchFromIndex);
      if (startIndex === -1) {
        break;
      }

      const endIndexExclusive = startIndex + normalizedQuery.length;
      matches.push({
        occurrenceIndex: matches.length,
        startIndex,
        endIndexExclusive,
        matchedText: cellContent.slice(startIndex, endIndexExclusive),
      });
      searchFromIndex = endIndexExclusive;
    }

    return matches;
  }

  private normalizeSearchMaxResults(maxResults: number | undefined): number {
    if (maxResults === undefined) {
      return DEFAULT_SEARCH_MAX_RESULTS;
    }

    if (maxResults === Number.POSITIVE_INFINITY) {
      return Number.POSITIVE_INFINITY;
    }

    if (!Number.isFinite(maxResults) || maxResults <= 0) {
      return 0;
    }

    return Math.floor(maxResults);
  }

  private *iterateStringCellsInSearchOrder(
    scopedSheets: SearchScopeSheet[]
  ): Generator<SearchScopeStringCell> {
    for (const { workbookName, sheet } of scopedSheets) {
      const indexes = this.getSheetIndexes({
        workbookName,
        sheetName: sheet.name,
      });
      const rowIndexes = Array.from(indexes.rowGroups.keys()).sort(
        (left, right) => left - right
      );

      for (const rowIndex of rowIndexes) {
        const rowGroup = indexes.rowGroups.get(rowIndex);
        if (!rowGroup) {
          continue;
        }

        for (const { number: colIndex, key: cellReference } of rowGroup) {
          const cellContent = sheet.content.get(cellReference);
          if (typeof cellContent !== "string") {
            continue;
          }

          yield {
            workbookName,
            sheetName: sheet.name,
            cellReference,
            cellContent,
            rowIndex,
            colIndex,
          };
        }
      }
    }
  }

  private buildSearchMatchesInScope(
    query: string,
    scopedSheets: SearchScopeSheet[],
    caseSensitive: boolean,
    maxResults: number
  ): SearchMatch[] {
    const results: SearchMatch[] = [];

    if (maxResults <= 0) {
      return results;
    }

    for (const cell of this.iterateStringCellsInSearchOrder(scopedSheets)) {
      const matches = this.findMatchesInString(
        cell.cellContent,
        query,
        caseSensitive,
        maxResults - results.length
      );
      if (matches.length === 0) {
        continue;
      }

      const contentKind = this.getStringContentKind(cell.cellContent);
      for (const match of matches) {
        results.push({
          workbookName: cell.workbookName,
          sheetName: cell.sheetName,
          cellReference: cell.cellReference,
          cellContent: cell.cellContent,
          contentKind,
          occurrenceIndex: match.occurrenceIndex,
          startIndex: match.startIndex,
          endIndexExclusive: match.endIndexExclusive,
          matchedText: match.matchedText,
        });
      }

      if (results.length >= maxResults) {
        return results;
      }
    }

    return results;
  }

  search(
    query: string,
    options?: SearchOptions
  ): SearchMatch[] {
    const scopedSheets = this.resolveSearchScope(options);
    const maxResults = this.normalizeSearchMaxResults(options?.maxResults);

    if (query.length === 0) {
      return [];
    }

    return this.buildSearchMatchesInScope(
      query,
      scopedSheets,
      options?.caseSensitive === true,
      maxResults
    );
  }

  private buildReplacedContent(
    beforeContent: string,
    matches: Array<Pick<SearchMatch, "startIndex" | "endIndexExclusive">>,
    replacement: string
  ): string {
    let cursor = 0;
    let replacedContent = "";

    for (const match of matches) {
      replacedContent += beforeContent.slice(cursor, match.startIndex);
      replacedContent += replacement;
      cursor = match.endIndexExclusive;
    }

    replacedContent += beforeContent.slice(cursor);
    return replacedContent;
  }

  prepareReplace(
    query: string,
    replacement: string,
    target: ReplaceTarget,
    options?: { caseSensitive?: boolean }
  ): PreparedReplace {
    if (query.length === 0) {
      throw new Error("replace requires a non-empty query");
    }

    const address: CellAddress = {
      workbookName: target.workbookName,
      sheetName: target.sheetName,
      ...parseCellReference(target.cellReference),
    };
    const beforeContent = this.getCellContent(address);

    if (typeof beforeContent !== "string") {
      throw new Error(
        `replace requires target cell ${target.cellReference} to contain a string`
      );
    }

    const matches = this.findMatchesInString(
      beforeContent,
      query,
      options?.caseSensitive === true
    );
    const match = matches[target.occurrenceIndex];

    if (!match) {
      throw new Error(
        `Occurrence ${target.occurrenceIndex} not found in cell ${target.cellReference}`
      );
    }

    const afterContent = this.buildReplacedContent(beforeContent, [match], replacement);
    const contentKind = this.getStringContentKind(beforeContent);

    return {
      address,
      beforeContent,
      afterContent,
      change: {
        workbookName: target.workbookName,
        sheetName: target.sheetName,
        cellReference: target.cellReference,
        contentKind,
        occurrenceIndex: match.occurrenceIndex,
        startIndex: match.startIndex,
        endIndexExclusive: match.endIndexExclusive,
        matchedText: match.matchedText,
        replacementText: replacement,
        beforeContent,
        afterContent,
      },
    };
  }

  prepareReplaceAll(
    query: string,
    replacement: string,
    options?: SearchOptions
  ): PreparedCellReplaceAll[] {
    const scopedSheets = this.resolveSearchScope(options);

    if (query.length === 0) {
      throw new Error("replaceAll requires a non-empty query");
    }

    const replacements: PreparedCellReplaceAll[] = [];
    const caseSensitive = options?.caseSensitive === true;

    for (const cell of this.iterateStringCellsInSearchOrder(scopedSheets)) {
      const matches = this.findMatchesInString(
        cell.cellContent,
        query,
        caseSensitive
      );
      if (matches.length === 0) {
        continue;
      }

      const address: CellAddress = {
        workbookName: cell.workbookName,
        sheetName: cell.sheetName,
        rowIndex: cell.rowIndex,
        colIndex: cell.colIndex,
      };
      const afterContent = this.buildReplacedContent(
        cell.cellContent,
        matches,
        replacement
      );
      const contentKind = this.getStringContentKind(cell.cellContent);

      replacements.push({
        address,
        beforeContent: cell.cellContent,
        afterContent,
        changes: matches.map((match) => ({
          workbookName: cell.workbookName,
          sheetName: cell.sheetName,
          cellReference: cell.cellReference,
          contentKind,
          occurrenceIndex: match.occurrenceIndex,
          startIndex: match.startIndex,
          endIndexExclusive: match.endIndexExclusive,
          matchedText: match.matchedText,
          replacementText: replacement,
          beforeContent: cell.cellContent,
          afterContent,
        })),
      });
    }

    return replacements;
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
