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

const DATA_CHANGE_OBSERVER_CHUNK_SIZE = 1024;

type IndexedMapInsertion<TValue> = {
  index: number;
  entry: [string, TValue];
  insertionOrder: number;
};

/**
 * Merges entries whose indexes are expressed in the final Map coordinate
 * system. Rebuilding once avoids the O(n*m) cost of repeated Array#splice.
 */
function mergeIndexedMapEntries<TValue>(
  baseEntries: Array<[string, TValue]>,
  rawInsertions: Array<Omit<IndexedMapInsertion<TValue>, "insertionOrder">>
): Array<[string, TValue]> {
  const insertions = rawInsertions
    .map((insertion, insertionOrder) => ({
      ...insertion,
      insertionOrder,
    }))
    .sort(
      (left, right) =>
        left.index - right.index || left.insertionOrder - right.insertionOrder
    );
  const result: Array<[string, TValue]> = [];
  let baseIndex = 0;
  let insertionIndex = 0;
  const finalLength = baseEntries.length + insertions.length;

  while (result.length < finalLength) {
    const insertion = insertions[insertionIndex];
    if (
      insertion &&
      Math.max(0, Math.min(insertion.index, finalLength - 1)) <= result.length
    ) {
      result.push(insertion.entry);
      insertionIndex++;
      continue;
    }

    const baseEntry = baseEntries[baseIndex];
    if (baseEntry) {
      result.push(baseEntry);
      baseIndex++;
      continue;
    }

    // Invalid externally supplied indexes are clamped to the tail while
    // preserving the observer's insertion order.
    if (insertion) {
      result.push(insertion.entry);
      insertionIndex++;
    }
  }

  return result;
}

/** Tracks available positions and selects by current rank in O(log n). */
class AvailablePositionIndex {
  private readonly tree: Int32Array;

  constructor(private readonly size: number) {
    this.tree = new Int32Array(size + 1);
    // Fenwick representation for an array initially filled with ones.
    for (let index = 1; index <= size; index++) {
      this.tree[index] = index & -index;
    }
  }

  take(rank: number): number {
    const available = this.sum(this.size);
    if (!Number.isSafeInteger(rank) || rank < 0 || rank >= available) {
      throw new Error(`Invalid ordered-map history index ${rank}`);
    }

    let target = rank + 1;
    let index = 0;
    let bit = 1;
    while (bit * 2 <= this.size) {
      bit *= 2;
    }
    for (; bit > 0; bit = Math.floor(bit / 2)) {
      const next = index + bit;
      if (next <= this.size && this.tree[next]! < target) {
        index = next;
        target -= this.tree[next]!;
      }
    }

    this.add(index, -1);
    return index;
  }

  private add(zeroBasedIndex: number, delta: number): void {
    for (
      let index = zeroBasedIndex + 1;
      index <= this.size;
      index += index & -index
    ) {
      this.tree[index]! += delta;
    }
  }

  private sum(count: number): number {
    let total = 0;
    for (let index = count; index > 0; index -= index & -index) {
      total += this.tree[index]!;
    }
    return total;
  }
}

export type CellContentDataChange = {
  kind: "cell-content";
  address: CellAddress;
  before: SerializedCellValue;
  after: SerializedCellValue;
  beforeIndex?: number;
  afterIndex?: number;
};

export type CellMetadataDataChange = {
  kind: "cell-metadata";
  address: CellAddress;
  before: unknown;
  after: unknown;
  beforeIndex?: number;
  afterIndex?: number;
};

export type SheetMetadataDataChange = {
  kind: "sheet-metadata";
  workbookName: string;
  sheetName: string;
  before: unknown;
  after: unknown;
};

export type WorkbookMetadataDataChange = {
  kind: "workbook-metadata";
  workbookName: string;
  before: unknown;
  after: unknown;
};

export type WorkbookDataChange =
  | CellContentDataChange
  | CellMetadataDataChange
  | SheetMetadataDataChange
  | WorkbookMetadataDataChange;

export type WorkbookDataChangePatch = {
  readonly changes: readonly WorkbookDataChange[];
  /** Fragments with the same id share one before/after index coordinate. */
  readonly atomicGroupId?: number;
  /** Sent after the manager mutation represented by the group is applied. */
  readonly committed?: true;
};

export type WorkbookMutationObserver = (
  patches: readonly WorkbookDataChangePatch[]
) => void;

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

  private dataChangeBatchDepth = 0;
  private pendingDataChangePatches: WorkbookDataChangePatch[] = [];
  private pendingDataChangeCount = 0;
  private nextAtomicDataChangeGroupId = 1;
  private mapTailKeys = new WeakMap<object, unknown>();

  // Map from "workbookName|sheetName" to indexes
  private sheetIndexes: Map<string, SheetIndexes> = new Map();

  constructor(
    private mutationObserver?: WorkbookMutationObserver,
    private readonly shouldObserve: () => boolean = () => true,
    private readonly shouldBatchDataChanges: () => boolean = () => true
  ) {}

  private get observingMutations(): boolean {
    return this.mutationObserver !== undefined && this.shouldObserve();
  }

  /**
   * Groups data-change notifications without delaying the underlying writes.
   * Copy/fill operations use this to keep observer overhead proportional to
   * the operation count rather than the number of affected cells.
   */
  batchDataChanges<T>(callback: () => T): T {
    if (!this.shouldBatchDataChanges()) {
      return callback();
    }

    this.dataChangeBatchDepth++;
    try {
      return callback();
    } finally {
      this.dataChangeBatchDepth--;
      if (this.dataChangeBatchDepth === 0) {
        const patches = this.pendingDataChangePatches;
        this.pendingDataChangePatches = [];
        this.pendingDataChangeCount = 0;
        if (patches.length > 0 && this.observingMutations) {
          this.mutationObserver!(patches);
        }
      }
    }
  }

  private reportDataChanges(
    changes: readonly WorkbookDataChange[],
    atomicGroupId?: number
  ): void {
    if (!this.observingMutations || changes.length === 0) {
      return;
    }

    this.enqueueDataChangePatch({
      changes,
      ...(atomicGroupId === undefined ? {} : { atomicGroupId }),
    });
  }

  private commitAtomicDataChangeGroup(atomicGroupId: number | undefined): void {
    if (!this.observingMutations || atomicGroupId === undefined) {
      return;
    }
    this.enqueueDataChangePatch({
      changes: [],
      atomicGroupId,
      committed: true,
    });
  }

  private enqueueDataChangePatch(patch: WorkbookDataChangePatch): void {
    if (this.dataChangeBatchDepth > 0 && this.shouldBatchDataChanges()) {
      this.pendingDataChangePatches.push(patch);
      this.pendingDataChangeCount += patch.changes.length;
      if (this.pendingDataChangeCount >= DATA_CHANGE_OBSERVER_CHUNK_SIZE) {
        const patches = this.pendingDataChangePatches;
        this.pendingDataChangePatches = [];
        this.pendingDataChangeCount = 0;
        this.mutationObserver!(patches);
      }
      return;
    }

    this.mutationObserver!([patch]);
  }

  private normalizeCellContent(
    content: SerializedCellValue
  ): SerializedCellValue {
    return this.isContentEmpty(content) ? undefined : content;
  }

  private getMapKeyIndex<TKey, TValue>(
    map: Map<TKey, TValue>,
    key: TKey
  ): number {
    const first = map.keys().next();
    if (!first.done && Object.is(first.value, key)) {
      return 0;
    }
    if (Object.is(this.mapTailKeys.get(map), key)) {
      return map.size - 1;
    }
    let index = 0;
    for (const candidate of map.keys()) {
      if (Object.is(candidate, key)) {
        return index;
      }
      index++;
    }
    return -1;
  }

  private trackMapTail<TKey, TValue>(map: Map<TKey, TValue>): void {
    let lastKey: TKey | undefined;
    for (const key of map.keys()) {
      lastKey = key;
    }
    if (lastKey === undefined) {
      this.mapTailKeys.delete(map);
    } else {
      this.mapTailKeys.set(map, lastKey);
    }
  }

  /**
   * Generate a key for the sheet indexes map
   */
  private getSheetIndexKey(workbookName: string, sheetName: string): string {
    return `${workbookName}|${sheetName}`;
  }

  /** Builds all four sheet indexes in linear collection + sort time. */
  private rebuildSheetIndexes(workbookName: string, sheet: Sheet): void {
    const indexKey = this.getSheetIndexKey(workbookName, sheet.name);
    this.sheetIndexes.delete(indexKey);
    const indexes = this.getSheetIndexes({
      workbookName,
      sheetName: sheet.name,
    });
    const byRow: Array<IndexEntry & { insertionOrder: number }> = [];
    const byCol: Array<IndexEntry & { insertionOrder: number }> = [];

    let insertionOrder = 0;
    let lastContentKey: string | undefined;
    for (const [key, storedValue] of sheet.content) {
      const value = this.normalizeCellContent(storedValue);
      if (value === undefined) {
        sheet.content.delete(key);
        continue;
      }
      if (!Object.is(value, storedValue)) {
        sheet.content.set(key, value);
      }

      const { colIndex, rowIndex } = parseCellReference(key);
      let rowGroup = indexes.rowGroups.get(rowIndex);
      if (!rowGroup) {
        rowGroup = [];
        indexes.rowGroups.set(rowIndex, rowGroup);
      }
      rowGroup.push({ number: colIndex, key });

      let colGroup = indexes.colGroups.get(colIndex);
      if (!colGroup) {
        colGroup = [];
        indexes.colGroups.set(colIndex, colGroup);
      }
      colGroup.push({ number: rowIndex, key });

      byRow.push({ number: rowIndex, key, insertionOrder });
      byCol.push({ number: colIndex, key, insertionOrder });
      lastContentKey = key;
      insertionOrder++;
    }
    if (lastContentKey === undefined) {
      this.mapTailKeys.delete(sheet.content);
    } else {
      this.mapTailKeys.set(sheet.content, lastContentKey);
    }
    this.trackMapTail(sheet.metadata);

    for (const rowGroup of indexes.rowGroups.values()) {
      rowGroup.sort((left, right) => left.number - right.number);
    }
    for (const colGroup of indexes.colGroups.values()) {
      colGroup.sort((left, right) => left.number - right.number);
    }

    // The previous binary insertion placed equal row/column entries before
    // earlier ones. Preserve that ordering while replacing O(n²) splices with
    // one O(n log n) sort.
    const compare = (
      left: IndexEntry & { insertionOrder: number },
      right: IndexEntry & { insertionOrder: number }
    ) =>
      left.number - right.number || right.insertionOrder - left.insertionOrder;
    byRow.sort(compare);
    byCol.sort(compare);
    indexes.cellsSortedByRow = byRow.map(({ number, key }) => ({
      number,
      key,
    }));
    indexes.cellsSortedByCol = byCol.map(({ number, key }) => ({
      number,
      key,
    }));
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
    if (this.workbooks.has(opts.newWorkbookName)) {
      throw new Error("Workbook with new name already exists");
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

    workbook.name = opts.newWorkbookName;

    const renamedWorkbooks = new Map<string, Workbook>();
    for (const [workbookName, existingWorkbook] of this.workbooks) {
      renamedWorkbooks.set(
        workbookName === opts.workbookName
          ? opts.newWorkbookName
          : workbookName,
        existingWorkbook
      );
    }
    this.workbooks = renamedWorkbooks;
  }

  resetWorkbooks(workbooks: Map<string, Workbook>): void {
    this.workbooks.clear();
    this.sheetIndexes.clear();

    workbooks.forEach((workbook, workbookName) => {
      this.workbooks.set(workbookName, workbook);
      workbook.sheets.forEach((sheet) => {
        this.rebuildSheetIndexes(workbookName, sheet);
      });
    });
  }

  /**
   * Restores one workbook and its map position for undo/redo without rebuilding
   * indexes for unrelated workbooks.
   */
  restoreWorkbookForHistory(opts: {
    workbookName: string;
    workbookOrder: readonly string[];
    workbook?: Workbook;
  }): void {
    const current = this.workbooks.get(opts.workbookName);
    if (current) {
      for (const sheetName of current.sheets.keys()) {
        this.sheetIndexes.delete(
          this.getSheetIndexKey(opts.workbookName, sheetName)
        );
      }
    }

    this.workbooks.delete(opts.workbookName);
    if (opts.workbook) {
      this.workbooks.set(opts.workbookName, opts.workbook);
      for (const sheet of opts.workbook.sheets.values()) {
        this.rebuildSheetIndexes(opts.workbookName, sheet);
      }
    }

    this.reorderWorkbooks(opts.workbookOrder);
  }

  /** Restores one sheet and its map position for undo/redo. */
  restoreSheetForHistory(opts: {
    workbookName: string;
    sheetName: string;
    sheetOrder: readonly string[];
    sheet?: Sheet;
  }): void {
    const workbook = this.workbooks.get(opts.workbookName);
    if (!workbook) {
      throw new WorkbookNotFoundError(opts.workbookName);
    }

    this.sheetIndexes.delete(
      this.getSheetIndexKey(opts.workbookName, opts.sheetName)
    );
    workbook.sheets.delete(opts.sheetName);

    if (opts.sheet) {
      workbook.sheets.set(opts.sheetName, opts.sheet);
      this.rebuildSheetIndexes(opts.workbookName, opts.sheet);
    }

    const reordered = new Map<string, Sheet>();
    for (const sheetName of opts.sheetOrder) {
      const sheet = workbook.sheets.get(sheetName);
      if (sheet) {
        reordered.set(sheetName, sheet);
      }
    }
    for (const [sheetName, sheet] of workbook.sheets) {
      if (!reordered.has(sheetName)) {
        reordered.set(sheetName, sheet);
      }
    }
    workbook.sheets = reordered;
  }

  private reorderWorkbooks(order: readonly string[]): void {
    const reordered = new Map<string, Workbook>();
    for (const workbookName of order) {
      const workbook = this.workbooks.get(workbookName);
      if (workbook) {
        reordered.set(workbookName, workbook);
      }
    }
    for (const [workbookName, workbook] of this.workbooks) {
      if (!reordered.has(workbookName)) {
        reordered.set(workbookName, workbook);
      }
    }
    this.workbooks = reordered;
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
    for (const [
      existingSheetName,
      existingSheet,
    ] of workbook.sheets.entries()) {
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
              this.setCellContent(address, `=${updatedFormula}`, { sheet });
              changed.push(address);
            }
          }
        });
      });
    };

    this.batchDataChanges(() => {
      this.workbooks.forEach((workbook, workbookName) => {
        update(workbookName, workbook.sheets);
      });
    });

    return changed;
  }

  updateFormulasExcluding(
    excludeCellsSet: Set<string>,
    updateCallback: (formula: string) => string
  ): void {
    this.batchDataChanges(() => {
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
                this.setCellContent(
                  { workbookName, sheetName, colIndex, rowIndex },
                  `=${updatedFormula}`,
                  { sheet }
                );
              }
            }
          });
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

    this.batchDataChanges(() => {
      workbook.sheets.forEach((sheet, sheetName) => {
        sheet.content.forEach((cell, key) => {
          if (typeof cell === "string" && cell.startsWith("=")) {
            const formula = cell.slice(1);
            const updatedFormula = updateCallback(formula);

            // Only update if the formula actually changed
            if (updatedFormula !== formula) {
              const { colIndex, rowIndex } = parseCellReference(key);
              this.setCellContent(
                { workbookName, sheetName, colIndex, rowIndex },
                `=${updatedFormula}`,
                { sheet }
              );
            }
          }
        });
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

  private resolveSearchScope(options?: SearchOptions): SearchScopeSheet[] {
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

  private getStringContentKind(
    cellContent: string
  ): SearchMatch["contentKind"] {
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
      const startIndex = normalizedContent.indexOf(
        normalizedQuery,
        searchFromIndex
      );
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

  search(query: string, options?: SearchOptions): SearchMatch[] {
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

    const afterContent = this.buildReplacedContent(
      beforeContent,
      [match],
      replacement
    );
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

  private setCellContentInternal(
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
    },
    reportChange = true
  ): boolean {
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

    const storedBefore = sheet.content.get(adr);
    const before = this.normalizeCellContent(storedBefore);
    const after = this.normalizeCellContent(content);
    const changed = !Object.is(before, after);
    const shouldReport = changed && reportChange && this.observingMutations;
    const beforeIndex =
      shouldReport && before !== undefined && after === undefined
        ? this.getMapKeyIndex(sheet.content, adr)
        : undefined;
    const afterIndex =
      shouldReport && before === undefined && after !== undefined
        ? sheet.content.size
        : undefined;

    if (after === undefined) {
      // Delete even when the normalized values are equal so legacy/directly
      // inserted empty-string entries are cleaned up.
      if (sheet.content.has(adr)) {
        sheet.content.delete(adr);
        if (!options?.buildingFromScratch && before !== undefined) {
          this.removeCellFromGroups(
            indexes,
            address.rowIndex,
            address.colIndex,
            adr
          );
        }
      }
      if (Object.is(this.mapTailKeys.get(sheet.content), adr)) {
        this.mapTailKeys.delete(sheet.content);
      }
    } else {
      sheet.content.set(adr, after);
      // Updating one non-empty value to another does not change membership in
      // any index. Avoid the previous duplicate scans on this hot path.
      if (options?.buildingFromScratch || before === undefined) {
        this.addCellToGroups(indexes, address.rowIndex, address.colIndex, adr);
        this.mapTailKeys.set(sheet.content, adr);
      }
    }

    if (shouldReport) {
      this.reportDataChanges([
        {
          kind: "cell-content",
          address: { ...address },
          before,
          after,
          ...(beforeIndex === undefined ? {} : { beforeIndex }),
          ...(afterIndex === undefined ? {} : { afterIndex }),
        },
      ]);
    }

    return changed;
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
    this.setCellContentInternal(address, content, options);
  }

  /**
   * Applies a retained cell-content patch and rebuilds each affected sheet's
   * indexes once. History replay can contain thousands of cells; routing those
   * writes through the single-cell index path would otherwise be quadratic.
   */
  applyCellContentChangesForHistory(
    changes: Iterable<CellContentDataChange>,
    direction: "undo" | "redo"
  ): void {
    const changesBySheet = new Map<
      string,
      {
        workbookName: string;
        sheet: Sheet;
        changes: CellContentDataChange[];
      }
    >();

    for (const change of changes) {
      const sheet = this.getSheet(change.address);
      if (!sheet) {
        throw new SheetNotFoundError(change.address.sheetName);
      }

      const sheetKey = this.getSheetIndexKey(
        change.address.workbookName,
        change.address.sheetName
      );
      let group = changesBySheet.get(sheetKey);
      if (!group) {
        group = {
          workbookName: change.address.workbookName,
          sheet,
          changes: [],
        };
        changesBySheet.set(sheetKey, group);
      }
      group.changes.push(change);
    }

    for (const {
      workbookName,
      sheet,
      changes: sheetChanges,
    } of changesBySheet.values()) {
      if (
        sheetChanges.length === 1 &&
        this.tryApplySingleCellContentMembershipChangeForHistory(
          sheet,
          sheetChanges[0]!,
          direction
        )
      ) {
        continue;
      }

      const indexedChanges = new Map<
        string,
        { value: SerializedCellValue; index?: number }
      >();
      const inPlaceChanges = new Map<string, SerializedCellValue>();
      for (const change of sheetChanges) {
        const key = getCellReference(change.address);
        const value = this.normalizeCellContent(
          direction === "redo" ? change.after : change.before
        );
        const sourceIndex =
          direction === "redo" ? change.beforeIndex : change.afterIndex;
        const targetIndex =
          direction === "redo" ? change.afterIndex : change.beforeIndex;
        if (sourceIndex !== undefined || targetIndex !== undefined) {
          indexedChanges.set(key, { value, index: targetIndex });
        } else {
          inPlaceChanges.set(key, value);
        }
      }

      // Replacing values does not affect Map order or any sheet index. Keep
      // the common single-cell undo/redo path proportional to the delta rather
      // than rebuilding an otherwise unrelated large sheet.
      let canApplyInPlace = indexedChanges.size === 0;
      if (canApplyInPlace) {
        for (const [key, value] of inPlaceChanges) {
          if (value === undefined || !sheet.content.has(key)) {
            canApplyInPlace = false;
            break;
          }
        }
      }
      if (canApplyInPlace) {
        for (const [key, value] of inPlaceChanges) {
          sheet.content.set(key, value);
        }
        continue;
      }

      const consumedInPlaceKeys = new Set<string>();
      const entries: Array<[string, SerializedCellValue]> = [];
      for (const [key, currentValue] of sheet.content) {
        if (indexedChanges.has(key)) {
          continue;
        }
        if (inPlaceChanges.has(key)) {
          consumedInPlaceKeys.add(key);
          const value = inPlaceChanges.get(key);
          if (value !== undefined) {
            entries.push([key, value]);
          }
        } else {
          entries.push([key, currentValue]);
        }
      }
      for (const [key, value] of inPlaceChanges) {
        if (!consumedInPlaceKeys.has(key) && value !== undefined) {
          entries.push([key, value]);
        }
      }

      const insertions = Array.from(indexedChanges, ([key, target]) => ({
        key,
        ...target,
      })).filter(
        (
          insertion
        ): insertion is {
          key: string;
          value: Exclude<SerializedCellValue, undefined>;
          index: number;
        } => insertion.value !== undefined && insertion.index !== undefined
      );

      const orderedEntries = mergeIndexedMapEntries(
        entries,
        insertions.map(({ key, value, index }) => ({
          index,
          entry: [key, value],
        }))
      );

      sheet.content.clear();
      for (const [key, value] of orderedEntries) {
        sheet.content.set(key, value);
      }
      this.rebuildSheetIndexes(workbookName, sheet);
    }
  }

  /**
   * Replays a run of single-cell deletions whose recorded indexes belong to
   * sequentially shrinking Maps. Redo is a stable filter. Undo maps each
   * historical rank back into the original coordinate space with a Fenwick
   * tree, then rebuilds each affected sheet once.
   */
  applySequentialCellContentDeletionsForHistory(
    patches: readonly (readonly WorkbookDataChange[])[],
    direction: "undo" | "redo"
  ): void {
    const changesBySheet = new Map<
      string,
      {
        workbookName: string;
        sheet: Sheet;
        changes: CellContentDataChange[];
      }
    >();

    for (const patch of patches) {
      if (patch.length !== 1 || patch[0]?.kind !== "cell-content") {
        throw new Error("Invalid sequential cell-content deletion history");
      }
      const change = patch[0];
      if (
        change.before === undefined ||
        change.after !== undefined ||
        change.beforeIndex === undefined ||
        change.afterIndex !== undefined
      ) {
        throw new Error("Invalid sequential cell-content deletion history");
      }
      const sheet = this.getSheet(change.address);
      if (!sheet) {
        throw new SheetNotFoundError(change.address.sheetName);
      }
      const sheetKey = this.getSheetIndexKey(
        change.address.workbookName,
        change.address.sheetName
      );
      let group = changesBySheet.get(sheetKey);
      if (!group) {
        group = {
          workbookName: change.address.workbookName,
          sheet,
          changes: [],
        };
        changesBySheet.set(sheetKey, group);
      }
      group.changes.push(change);
    }

    for (const { workbookName, sheet, changes } of changesBySheet.values()) {
      if (direction === "redo") {
        const deletedKeys = new Set(
          changes.map((change) => getCellReference(change.address))
        );
        const entries = Array.from(sheet.content).filter(
          ([key]) => !deletedKeys.has(key)
        );
        if (sheet.content.size - entries.length !== deletedKeys.size) {
          throw new Error("Cell-content deletion history does not match state");
        }
        sheet.content.clear();
        for (const [key, value] of entries) {
          sheet.content.set(key, value);
        }
        this.rebuildSheetIndexes(workbookName, sheet);
        continue;
      }

      const finalEntries = Array.from(sheet.content);
      const originalLength = finalEntries.length + changes.length;
      const availablePositions = new AvailablePositionIndex(originalLength);
      const restored: Array<
        [string, Exclude<SerializedCellValue, undefined>] | undefined
      > = new Array(originalLength);

      for (const change of changes) {
        const position = availablePositions.take(change.beforeIndex!);
        const value = this.normalizeCellContent(change.before);
        if (value === undefined || restored[position] !== undefined) {
          throw new Error("Invalid sequential cell-content deletion history");
        }
        restored[position] = [getCellReference(change.address), value];
      }

      let finalIndex = 0;
      for (let index = 0; index < restored.length; index++) {
        if (restored[index] === undefined) {
          const entry = finalEntries[finalIndex++];
          if (!entry) {
            throw new Error("Cell-content deletion history is incomplete");
          }
          restored[index] = entry as [
            string,
            Exclude<SerializedCellValue, undefined>
          ];
        }
      }
      if (finalIndex !== finalEntries.length) {
        throw new Error("Cell-content deletion history does not match state");
      }

      sheet.content.clear();
      for (const [key, value] of restored as Array<
        [string, Exclude<SerializedCellValue, undefined>]
      >) {
        sheet.content.set(key, value);
      }
      this.rebuildSheetIndexes(workbookName, sheet);
    }
  }

  private tryApplySingleCellContentMembershipChangeForHistory(
    sheet: Sheet,
    change: CellContentDataChange,
    direction: "undo" | "redo"
  ): boolean {
    const key = getCellReference(change.address);
    const value = this.normalizeCellContent(
      direction === "redo" ? change.after : change.before
    );
    const sourceIndex =
      direction === "redo" ? change.beforeIndex : change.afterIndex;
    const targetIndex =
      direction === "redo" ? change.afterIndex : change.beforeIndex;

    if (
      value === undefined &&
      sourceIndex === sheet.content.size - 1 &&
      sheet.content.has(key)
    ) {
      this.setCellContentInternal(change.address, undefined, { sheet }, false);
      return true;
    }
    if (
      value !== undefined &&
      targetIndex === sheet.content.size &&
      !sheet.content.has(key)
    ) {
      this.setCellContentInternal(change.address, value, { sheet }, false);
      return true;
    }
    return false;
  }

  /**
   * Set metadata for a cell
   */
  setCellMetadata<TMetadata = unknown>(
    address: CellAddress,
    metadata: TMetadata | undefined
  ): void {
    const sheet = this.getSheet({
      workbookName: address.workbookName,
      sheetName: address.sheetName,
    });
    if (!sheet) {
      throw new SheetNotFoundError(address.sheetName);
    }

    const key = getCellReference(address);
    const before = sheet.metadata.get(key);
    if (Object.is(before, metadata)) {
      return;
    }
    const shouldReport = this.observingMutations;
    const beforeIndex =
      shouldReport && before !== undefined && metadata === undefined
        ? this.getMapKeyIndex(sheet.metadata, key)
        : undefined;
    const afterIndex =
      shouldReport && before === undefined && metadata !== undefined
        ? sheet.metadata.size
        : undefined;

    if (metadata === undefined) {
      sheet.metadata.delete(key);
      if (Object.is(this.mapTailKeys.get(sheet.metadata), key)) {
        this.mapTailKeys.delete(sheet.metadata);
      }
    } else {
      sheet.metadata.set(key, metadata);
      if (before === undefined) {
        this.mapTailKeys.set(sheet.metadata, key);
      }
    }

    if (shouldReport) {
      this.reportDataChanges([
        {
          kind: "cell-metadata",
          address: { ...address },
          before,
          after: metadata,
          ...(beforeIndex === undefined ? {} : { beforeIndex }),
          ...(afterIndex === undefined ? {} : { afterIndex }),
        },
      ]);
    }
  }

  /** Restores ordered cell metadata in one rebuild per affected sheet. */
  applyCellMetadataChangesForHistory(
    changes: Iterable<CellMetadataDataChange>,
    direction: "undo" | "redo",
    cloneValue: (value: unknown) => unknown
  ): void {
    const changesBySheet = new Map<
      string,
      { sheet: Sheet; changes: CellMetadataDataChange[] }
    >();
    for (const change of changes) {
      const sheet = this.getSheet(change.address);
      if (!sheet) {
        throw new SheetNotFoundError(change.address.sheetName);
      }
      const sheetKey = this.getSheetIndexKey(
        change.address.workbookName,
        change.address.sheetName
      );
      let group = changesBySheet.get(sheetKey);
      if (!group) {
        group = { sheet, changes: [] };
        changesBySheet.set(sheetKey, group);
      }
      group.changes.push(change);
    }

    for (const { sheet, changes: sheetChanges } of changesBySheet.values()) {
      if (sheetChanges.length === 1) {
        const change = sheetChanges[0]!;
        const key = getCellReference(change.address);
        const value = cloneValue(
          direction === "redo" ? change.after : change.before
        );
        const sourceIndex =
          direction === "redo" ? change.beforeIndex : change.afterIndex;
        const targetIndex =
          direction === "redo" ? change.afterIndex : change.beforeIndex;
        if (
          value === undefined &&
          sourceIndex === sheet.metadata.size - 1 &&
          sheet.metadata.has(key)
        ) {
          sheet.metadata.delete(key);
          if (Object.is(this.mapTailKeys.get(sheet.metadata), key)) {
            this.mapTailKeys.delete(sheet.metadata);
          }
          continue;
        }
        if (
          value !== undefined &&
          targetIndex === sheet.metadata.size &&
          !sheet.metadata.has(key)
        ) {
          sheet.metadata.set(key, value);
          this.mapTailKeys.set(sheet.metadata, key);
          continue;
        }
      }

      const indexedChanges = new Map<
        string,
        { value: unknown; index?: number }
      >();
      const inPlaceChanges = new Map<string, unknown>();
      for (const change of sheetChanges) {
        const key = getCellReference(change.address);
        const value = cloneValue(
          direction === "redo" ? change.after : change.before
        );
        const sourceIndex =
          direction === "redo" ? change.beforeIndex : change.afterIndex;
        const targetIndex =
          direction === "redo" ? change.afterIndex : change.beforeIndex;
        if (sourceIndex !== undefined || targetIndex !== undefined) {
          indexedChanges.set(key, { value, index: targetIndex });
        } else {
          inPlaceChanges.set(key, value);
        }
      }

      // Existing-key metadata replacements preserve insertion order. Avoid a
      // full Map copy for the overwhelmingly common sparse replay case.
      let canApplyInPlace = indexedChanges.size === 0;
      if (canApplyInPlace) {
        for (const [key, value] of inPlaceChanges) {
          if (value === undefined || !sheet.metadata.has(key)) {
            canApplyInPlace = false;
            break;
          }
        }
      }
      if (canApplyInPlace) {
        for (const [key, value] of inPlaceChanges) {
          sheet.metadata.set(key, value);
        }
        continue;
      }

      const consumedInPlaceKeys = new Set<string>();
      const entries: Array<[string, unknown]> = [];
      for (const [key, currentValue] of sheet.metadata) {
        if (indexedChanges.has(key)) {
          continue;
        }
        if (inPlaceChanges.has(key)) {
          consumedInPlaceKeys.add(key);
          const value = inPlaceChanges.get(key);
          if (value !== undefined) {
            entries.push([key, value]);
          }
        } else {
          entries.push([key, currentValue]);
        }
      }
      for (const [key, value] of inPlaceChanges) {
        if (!consumedInPlaceKeys.has(key) && value !== undefined) {
          entries.push([key, value]);
        }
      }

      const insertions = Array.from(indexedChanges, ([key, target]) => ({
        key,
        ...target,
      })).filter(
        (
          insertion
        ): insertion is {
          key: string;
          value: unknown;
          index: number;
        } => insertion.value !== undefined && insertion.index !== undefined
      );

      sheet.metadata = new Map(
        mergeIndexedMapEntries(
          entries,
          insertions.map(({ key, value, index }) => ({
            index,
            entry: [key, value],
          }))
        )
      );
      this.trackMapTail(sheet.metadata);
    }
  }

  /**
   * Get metadata for a cell
   */
  getCellMetadata<TMetadata = unknown>(
    address: CellAddress
  ): TMetadata | undefined {
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
    const before = sheet.sheetMetadata;
    if (Object.is(before, metadata)) {
      return;
    }

    sheet.sheetMetadata = metadata;
    if (this.observingMutations) {
      this.reportDataChanges([
        {
          kind: "sheet-metadata",
          ...opts,
          before,
          after: metadata,
        },
      ]);
    }
  }

  /**
   * Get metadata for a sheet
   */
  getSheetMetadata<TSheetMetadata = unknown>(opts: {
    workbookName: string;
    sheetName: string;
  }): TSheetMetadata | undefined {
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

    const before = workbook.workbookMetadata;
    if (Object.is(before, metadata)) {
      return;
    }

    workbook.workbookMetadata = metadata;
    if (this.observingMutations) {
      this.reportDataChanges([
        {
          kind: "workbook-metadata",
          workbookName,
          before,
          after: metadata,
        },
      ]);
    }
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
    const replacementContent =
      newContent === sheet.content ? new Map(newContent) : newContent;

    // Only pay the diffing cost when a consumer needs mutation data. The
    // replacement itself remains a clear-and-rebuild operation so index
    // construction stays linear and does not degrade into repeated removals.
    let changes: CellContentDataChange[] = [];
    const atomicGroupId = this.observingMutations
      ? this.nextAtomicDataChangeGroupId++
      : undefined;
    const reportChange = (change: CellContentDataChange) => {
      changes.push(change);
      if (changes.length >= DATA_CHANGE_OBSERVER_CHUNK_SIZE) {
        this.reportDataChanges(changes, atomicGroupId);
        changes = [];
      }
    };
    if (this.observingMutations) {
      // Reorder detection only needs target indexes for keys already present
      // in the sheet. New keys can use the running target index directly.
      // This avoids a sheet-sized history-only Map for empty/disjoint imports.
      const existingAfterIndexes = new Map<string, number>();
      let nextAfterIndex = 0;
      for (const [cellReference, storedAfter] of replacementContent) {
        if (this.normalizeCellContent(storedAfter) !== undefined) {
          if (sheet.content.has(cellReference)) {
            existingAfterIndexes.set(cellReference, nextAfterIndex);
          }
          nextAfterIndex++;
        }
      }

      let nextBeforeIndex = 0;
      for (const [cellReference, storedBefore] of sheet.content) {
        const before = this.normalizeCellContent(storedBefore);
        const after = this.normalizeCellContent(
          replacementContent.get(cellReference)
        );
        const beforeIndex =
          before === undefined ? undefined : nextBeforeIndex++;
        const afterIndex = existingAfterIndexes.get(cellReference);
        if (Object.is(before, after) && Object.is(beforeIndex, afterIndex)) {
          continue;
        }

        reportChange({
          kind: "cell-content",
          address: {
            ...opts,
            ...parseCellReference(cellReference),
          },
          before,
          after,
          ...(beforeIndex === undefined ? {} : { beforeIndex }),
          ...(afterIndex === undefined ? {} : { afterIndex }),
        });
      }

      let currentAfterIndex = 0;
      for (const [cellReference, storedAfter] of replacementContent) {
        const after = this.normalizeCellContent(storedAfter);
        if (after === undefined) {
          continue;
        }

        const afterIndex = currentAfterIndex++;
        if (sheet.content.has(cellReference)) {
          continue;
        }

        reportChange({
          kind: "cell-content",
          address: {
            ...opts,
            ...parseCellReference(cellReference),
          },
          before: undefined,
          after,
          afterIndex,
        });
      }
    }

    // Emit the final fragment before applying the replacement. If an explicit
    // transaction rejects the history payload, no part of this atomic group
    // has reached workbook state yet.
    this.reportDataChanges(changes, atomicGroupId);

    // Clear existing content without breaking the Map reference
    sheet.content.clear();

    // Repopulate first, then build all indexes with one collection/sort pass.
    // This avoids repeated flat-array scans and splices for large imports.
    replacementContent.forEach((value, key) => {
      const normalized = this.normalizeCellContent(value);
      if (normalized !== undefined) {
        sheet.content.set(key, normalized);
      }
    });
    this.rebuildSheetIndexes(opts.workbookName, sheet);

    this.commitAtomicDataChangeGroup(atomicGroupId);
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

    let metadataChanges: CellMetadataDataChange[] = [];
    const metadataIndexes = this.observingMutations
      ? new Map(Array.from(sheet.metadata.keys(), (key, index) => [key, index]))
      : undefined;
    const metadataAtomicGroupId = this.observingMutations
      ? this.nextAtomicDataChangeGroupId++
      : undefined;

    this.batchDataChanges(() => {
      // Use iterateCellsInRange to only process cells that actually exist.
      // Stream metadata deltas so a large clear never constructs one giant
      // observer payload before the history budget can reject it.
      for (const cellAddress of this.iterateCellsInRange(address)) {
        const cellRef = getCellReference(cellAddress);
        newContent.delete(cellRef);
        const beforeMetadata = sheet.metadata.get(cellRef);
        newMetadata.delete(cellRef);
        if (this.observingMutations && beforeMetadata !== undefined) {
          metadataChanges.push({
            kind: "cell-metadata",
            address: { ...cellAddress },
            before: beforeMetadata,
            after: undefined,
            beforeIndex: metadataIndexes!.get(cellRef),
          });
          if (metadataChanges.length >= DATA_CHANGE_OBSERVER_CHUNK_SIZE) {
            this.reportDataChanges(metadataChanges, metadataAtomicGroupId);
            metadataChanges = [];
          }
        }
      }
      this.reportDataChanges(metadataChanges, metadataAtomicGroupId);

      // Update content
      this.setSheetContent(address, newContent);

      // Update metadata
      sheet.metadata = newMetadata;
      this.trackMapTail(newMetadata);
      this.commitAtomicDataChangeGroup(metadataAtomicGroupId);
    });
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
