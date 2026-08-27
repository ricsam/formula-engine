import type {
  CellAddress,
  SerializedCellValue,
  SpreadsheetRange,
  SpreadsheetRangeEnd,
  TableDefinition,
} from "../types";
import type { TableManagerSnapshot } from "../engine-snapshot";
import {
  checkRangeIntersection,
  getCellReference,
  parseCellReference,
} from "../utils";
import type { WorkbookManager } from "./workbook-manager";
import type { MutationDirection } from "./mutation-observer";

export type TableEntryState = {
  workbookName: string;
  tableName: string;
  table: TableDefinition;
  index: number;
};

export type TableWorkbookBucketState = {
  workbookName: string;
  index: number;
};

export type TableMutation =
  | {
      kind: "table";
      before?: TableEntryState;
      after?: TableEntryState;
    }
  | {
      kind: "table-workbook-bucket";
      before?: TableWorkbookBucketState;
      after?: TableWorkbookBucketState;
    };

export type TableMutationObserver = (changes: readonly TableMutation[]) => void;

type TableManagerState = {
  entries: TableEntryState[];
  buckets: TableWorkbookBucketState[];
};

function cloneTableDefinition(table: TableDefinition): TableDefinition {
  return {
    ...table,
    start: { ...table.start },
    endRow: { ...table.endRow },
    headers: new Map(
      Array.from(table.headers, ([name, header]) => [name, { ...header }])
    ),
  };
}

function cloneTableEntryState(
  state: TableEntryState | undefined
): TableEntryState | undefined {
  return state
    ? {
        workbookName: state.workbookName,
        tableName: state.tableName,
        table: cloneTableDefinition(state.table),
        index: state.index,
      }
    : undefined;
}

function cloneTableBucketState(
  state: TableWorkbookBucketState | undefined
): TableWorkbookBucketState | undefined {
  return state ? { ...state } : undefined;
}

function cloneTableMutation(change: TableMutation): TableMutation {
  if (change.kind === "table") {
    return {
      kind: change.kind,
      before: cloneTableEntryState(change.before),
      after: cloneTableEntryState(change.after),
    };
  }
  return {
    kind: change.kind,
    before: cloneTableBucketState(change.before),
    after: cloneTableBucketState(change.after),
  };
}

function tableEntryKey(
  state: Pick<TableEntryState, "workbookName" | "tableName">
): string {
  return JSON.stringify([state.workbookName, state.tableName]);
}

function tableDefinitionsEqual(
  left: TableDefinition,
  right: TableDefinition
): boolean {
  if (
    left.name !== right.name ||
    left.workbookName !== right.workbookName ||
    left.sheetName !== right.sheetName ||
    left.start.rowIndex !== right.start.rowIndex ||
    left.start.colIndex !== right.start.colIndex ||
    left.endRow.type !== right.endRow.type ||
    (left.endRow.type === "number" &&
      (right.endRow.type !== "number" ||
        left.endRow.value !== right.endRow.value)) ||
    left.headers.size !== right.headers.size
  ) {
    return false;
  }
  const leftHeaders = Array.from(left.headers.entries());
  const rightHeaders = Array.from(right.headers.entries());
  return leftHeaders.every(([leftKey, leftHeader], index) => {
    const rightEntry = rightHeaders[index];
    return (
      rightEntry !== undefined &&
      leftKey === rightEntry[0] &&
      leftHeader.name === rightEntry[1].name &&
      leftHeader.index === rightEntry[1].index
    );
  });
}

function tableEntryStatesEqual(
  left: TableEntryState,
  right: TableEntryState
): boolean {
  return (
    left.workbookName === right.workbookName &&
    left.tableName === right.tableName &&
    left.index === right.index &&
    tableDefinitionsEqual(left.table, right.table)
  );
}

type IndexedMapInsertion<TKey, TValue> = {
  key: TKey;
  value: TValue;
  index: number;
  order: number;
};

/** Rebuild an ordered Map once from absolute target indices. */
function applyOrderedMapPatch<TKey, TValue>(
  map: Map<TKey, TValue>,
  removedKeys: ReadonlySet<TKey>,
  rawInsertions: readonly IndexedMapInsertion<TKey, TValue>[]
): void {
  if (removedKeys.size === 0 && rawInsertions.length === 0) {
    return;
  }
  const insertionKeys = new Set(rawInsertions.map(({ key }) => key));
  const remaining = Array.from(map.entries()).filter(
    ([key]) => !removedKeys.has(key) && !insertionKeys.has(key)
  );
  const insertions = [...rawInsertions].sort(
    (left, right) => left.index - right.index || right.order - left.order
  );
  const rebuilt: Array<[TKey, TValue]> = [];
  let remainingIndex = 0;

  for (const insertion of insertions) {
    const targetIndex = Math.max(0, insertion.index);
    while (rebuilt.length < targetIndex && remainingIndex < remaining.length) {
      rebuilt.push(remaining[remainingIndex++]!);
    }
    rebuilt.push([insertion.key, insertion.value]);
  }
  while (remainingIndex < remaining.length) {
    rebuilt.push(remaining[remainingIndex++]!);
  }

  map.clear();
  for (const [key, value] of rebuilt) {
    map.set(key, value);
  }
}

export type TableHeaderUpdate = {
  table: TableDefinition;
  index: number;
  oldName: string;
  newName: string;
};

export type GeneratedTableHeader = {
  address: CellAddress;
  name: string;
};

export class TableManager {
  tables: Map<
    /**
     * workbook name -> table name -> table definition
     */
    string,
    Map<string, TableDefinition>
  > = new Map();
  private workbookManager: WorkbookManager;

  private mutationBatchDepth = 0;
  private pendingMutationPatches: TableMutation[][] = [];
  private mutationReportingSuppressionDepth = 0;

  constructor(
    workbookManager: WorkbookManager,
    private mutationObserver?: TableMutationObserver,
    private readonly shouldObserve: () => boolean = () => true,
    private readonly detachMutationValues = true,
    private readonly shouldBatchMutations: () => boolean = () => true
  ) {
    this.workbookManager = workbookManager;
  }

  private get observingMutations(): boolean {
    return (
      this.mutationObserver !== undefined &&
      this.mutationReportingSuppressionDepth === 0 &&
      this.shouldObserve()
    );
  }

  batchMutations<T>(callback: () => T): T {
    if (!this.shouldBatchMutations()) {
      return callback();
    }

    this.mutationBatchDepth++;
    try {
      return callback();
    } finally {
      this.mutationBatchDepth--;
      if (this.mutationBatchDepth === 0) {
        const patches = this.pendingMutationPatches;
        this.pendingMutationPatches = [];
        for (const patch of patches) {
          this.emitMutationPatch(patch);
        }
      }
    }
  }

  private reportMutations(changes: readonly TableMutation[]): void {
    if (!this.observingMutations || changes.length === 0) {
      return;
    }
    const detached = this.detachMutationValues
      ? changes.map(cloneTableMutation)
      : [...changes];
    if (this.mutationBatchDepth > 0) {
      this.pendingMutationPatches.push(detached);
      return;
    }
    this.emitMutationPatch(detached);
  }

  private emitMutationPatch(changes: readonly TableMutation[]): void {
    if (!this.observingMutations) {
      return;
    }
    this.mutationObserver!(changes);
  }

  private getMapIndex<TKey, TValue>(map: Map<TKey, TValue>, key: TKey): number {
    let index = 0;
    for (const candidate of map.keys()) {
      if (Object.is(candidate, key)) {
        return index;
      }
      index++;
    }
    return -1;
  }

  private captureEntryState(
    workbookName: string,
    tableName: string
  ): TableEntryState | undefined {
    const tables = this.tables.get(workbookName);
    const table = tables?.get(tableName);
    if (!tables || !table) {
      return undefined;
    }
    return {
      workbookName,
      tableName,
      table: cloneTableDefinition(table),
      index: this.getMapIndex(tables, tableName),
    };
  }

  private captureTableObjectStates(
    targets: ReadonlySet<TableDefinition>
  ): Map<TableDefinition, TableEntryState> {
    const states = new Map<TableDefinition, TableEntryState>();
    if (targets.size === 0) {
      return states;
    }
    for (const [workbookName, tables] of this.tables) {
      let index = 0;
      for (const [tableName, table] of tables) {
        if (targets.has(table)) {
          states.set(table, {
            workbookName,
            tableName,
            table: cloneTableDefinition(table),
            index,
          });
        }
        index++;
      }
    }
    return states;
  }

  private captureWorkbookState(
    workbookNames?: ReadonlySet<string>
  ): TableManagerState {
    const entries: TableEntryState[] = [];
    const buckets: TableWorkbookBucketState[] = [];
    let workbookIndex = 0;
    for (const [workbookName, tables] of this.tables) {
      if (!workbookNames || workbookNames.has(workbookName)) {
        buckets.push({ workbookName, index: workbookIndex });
        let tableIndex = 0;
        for (const [tableName, table] of tables) {
          entries.push({
            workbookName,
            tableName,
            table: cloneTableDefinition(table),
            index: tableIndex++,
          });
        }
      }
      workbookIndex++;
    }
    return { entries, buckets };
  }

  private buildStateMutations(
    before: TableManagerState,
    after: TableManagerState
  ): TableMutation[] {
    const changes: TableMutation[] = [];
    const beforeBuckets = new Map(
      before.buckets.map((state) => [state.workbookName, state])
    );
    const afterBuckets = new Map(
      after.buckets.map((state) => [state.workbookName, state])
    );
    for (const workbookName of new Set([
      ...beforeBuckets.keys(),
      ...afterBuckets.keys(),
    ])) {
      const beforeState = beforeBuckets.get(workbookName);
      const afterState = afterBuckets.get(workbookName);
      if (beforeState && afterState && beforeState.index === afterState.index) {
        continue;
      }
      changes.push({
        kind: "table-workbook-bucket",
        before: cloneTableBucketState(beforeState),
        after: cloneTableBucketState(afterState),
      });
    }

    const beforeEntries = new Map(
      before.entries.map((state) => [tableEntryKey(state), state])
    );
    const afterEntries = new Map(
      after.entries.map((state) => [tableEntryKey(state), state])
    );
    for (const key of new Set([
      ...beforeEntries.keys(),
      ...afterEntries.keys(),
    ])) {
      const beforeState = beforeEntries.get(key);
      const afterState = afterEntries.get(key);
      if (
        beforeState &&
        afterState &&
        tableEntryStatesEqual(beforeState, afterState)
      ) {
        continue;
      }
      changes.push({
        kind: "table",
        before: cloneTableEntryState(beforeState),
        after: cloneTableEntryState(afterState),
      });
    }
    return changes;
  }

  private mutateWithStateDiff<T>(callback: () => T): T {
    if (!this.observingMutations) {
      return callback();
    }
    const before = this.captureWorkbookState();
    const result = callback();
    const after = this.captureWorkbookState();
    this.reportMutations(this.buildStateMutations(before, after));
    return result;
  }

  /** Applies detached observer deltas exactly, including map order and buckets. */
  private applyMutations(
    changes: readonly TableMutation[],
    direction: "before" | "after"
  ): void {
    this.batchMutations(() => {
      const entryChanges = changes.filter(
        (change): change is Extract<TableMutation, { kind: "table" }> =>
          change.kind === "table"
      );
      const bucketChanges = changes.filter(
        (
          change
        ): change is Extract<
          TableMutation,
          { kind: "table-workbook-bucket" }
        > => change.kind === "table-workbook-bucket"
      );

      for (const change of entryChanges) {
        for (const state of [change.before, change.after]) {
          if (state) {
            this.tables.get(state.workbookName)?.delete(state.tableName);
          }
        }
      }

      const bucketRemovals = new Set<string>();
      const bucketInsertions: IndexedMapInsertion<
        string,
        Map<string, TableDefinition>
      >[] = [];
      bucketChanges.forEach((change, order) => {
        if (change.before) {
          bucketRemovals.add(change.before.workbookName);
        }
        if (change.after) {
          bucketRemovals.add(change.after.workbookName);
        }
        const target = change[direction];
        if (!target) {
          return;
        }
        const other = direction === "before" ? change.after : change.before;
        bucketInsertions.push({
          key: target.workbookName,
          value:
            this.tables.get(target.workbookName) ??
            (other === undefined
              ? undefined
              : this.tables.get(other.workbookName)) ??
            new Map(),
          index: target.index,
          order,
        });
      });
      applyOrderedMapPatch(this.tables, bucketRemovals, bucketInsertions);

      const targetEntriesByWorkbook = new Map<
        string,
        IndexedMapInsertion<string, TableDefinition>[]
      >();
      entryChanges.forEach((change, order) => {
        const state = change[direction];
        if (!state) {
          return;
        }
        let insertions = targetEntriesByWorkbook.get(state.workbookName);
        if (!insertions) {
          insertions = [];
          targetEntriesByWorkbook.set(state.workbookName, insertions);
        }
        insertions.push({
          key: state.tableName,
          value: cloneTableDefinition(state.table),
          index: state.index,
          order,
        });
      });
      for (const [workbookName, insertions] of targetEntriesByWorkbook) {
        let workbookTables = this.tables.get(workbookName);
        if (!workbookTables) {
          workbookTables = new Map();
          this.tables.set(workbookName, workbookTables);
        }
        applyOrderedMapPatch(workbookTables, new Set(), insertions);
      }

      const reported =
        direction === "after"
          ? changes
          : changes.map(
              (change): TableMutation =>
                change.kind === "table"
                  ? {
                      kind: change.kind,
                      before: change.after,
                      after: change.before,
                    }
                  : {
                      kind: change.kind,
                      before: change.after,
                      after: change.before,
                    }
            );
      this.reportMutations(reported);
    });
  }

  /** Applies retained deltas directly without notifying the observer. */
  applyHistoryChanges(
    changes: readonly TableMutation[],
    direction: MutationDirection
  ): void {
    this.mutationReportingSuppressionDepth++;
    try {
      this.applyMutations(changes, direction === "undo" ? "before" : "after");
    } finally {
      this.mutationReportingSuppressionDepth--;
    }
  }

  getTables(workbookName: string): Map<string, TableDefinition> {
    return this.tables.get(workbookName) ?? new Map();
  }

  getTable(opts: {
    workbookName: string;
    name: string;
  }): TableDefinition | undefined {
    return this.tables.get(opts.workbookName)?.get(opts.name);
  }

  private getDefaultHeaderName(index: number, usedNames: Set<string>): string {
    let number = index + 1;
    let name = `Column ${number}`;
    while (usedNames.has(name)) {
      number++;
      name = `Column ${number}`;
    }
    return name;
  }

  private getHeaderName(
    value: SerializedCellValue,
    index: number,
    usedNames: Set<string>
  ): string {
    if (value === undefined || value === "") {
      return this.getDefaultHeaderName(index, usedNames);
    }
    return String(value);
  }

  private replaceHeader(
    table: TableDefinition,
    index: number,
    newName: string
  ): void {
    const headers = new Map<string, { name: string; index: number }>();
    for (const header of Array.from(table.headers.values()).sort(
      (a, b) => a.index - b.index
    )) {
      const name = header.index === index ? newName : header.name;
      headers.set(name, { name, index: header.index });
    }
    table.headers = headers;
  }

  prepareHeaderUpdate(
    address: CellAddress,
    value: SerializedCellValue
  ): { content: SerializedCellValue; updates: TableHeaderUpdate[] } {
    const updates: TableHeaderUpdate[] = [];

    for (const table of this.getTables(address.workbookName).values()) {
      if (
        table.sheetName !== address.sheetName ||
        address.rowIndex !== table.start.rowIndex
      ) {
        continue;
      }

      const index = address.colIndex - table.start.colIndex;
      if (index < 0 || index >= table.headers.size) {
        continue;
      }

      const oldHeader = Array.from(table.headers.values()).find(
        (header) => header.index === index
      );
      if (!oldHeader) {
        continue;
      }

      const usedNames = new Set(
        Array.from(table.headers.values())
          .filter((header) => header.index !== index)
          .map((header) => header.name)
      );
      const newName = this.getHeaderName(value, index, usedNames);
      if (usedNames.has(newName)) {
        throw new Error(`Duplicate table header "${newName}"`);
      }

      updates.push({
        table,
        index,
        oldName: oldHeader.name,
        newName,
      });
    }

    const generatedNames = new Set(
      updates
        .filter(() => value === undefined || value === "")
        .map((update) => update.newName)
    );
    if (generatedNames.size > 1) {
      throw new Error("Overlapping tables require the same header name");
    }

    return {
      content:
        value === undefined || value === ""
          ? updates[0]?.newName ?? value
          : value,
      updates,
    };
  }

  prepareHeaderUpdatesForSheet(options: {
    workbookName: string;
    sheetName: string;
    getCellContent: (address: CellAddress) => SerializedCellValue;
  }): {
    updates: TableHeaderUpdate[];
    generatedHeaders: GeneratedTableHeader[];
  } {
    const updates: TableHeaderUpdate[] = [];
    const generatedHeaders = new Map<string, GeneratedTableHeader>();

    for (const table of this.getTables(options.workbookName).values()) {
      if (table.sheetName !== options.sheetName) {
        continue;
      }

      const oldHeadersByIndex = new Map(
        Array.from(table.headers.values()).map((header) => [
          header.index,
          header,
        ])
      );
      const usedNames = new Set<string>();
      for (let index = 0; index < table.headers.size; index++) {
        const address = {
          workbookName: table.workbookName,
          sheetName: table.sheetName,
          rowIndex: table.start.rowIndex,
          colIndex: table.start.colIndex + index,
        };
        const value = options.getCellContent(address);
        const newName = this.getHeaderName(value, index, usedNames);
        if (usedNames.has(newName)) {
          throw new Error(`Duplicate table header "${newName}"`);
        }
        usedNames.add(newName);

        const oldHeader = oldHeadersByIndex.get(index);
        if (!oldHeader) {
          continue;
        }
        updates.push({
          table,
          index,
          oldName: oldHeader.name,
          newName,
        });

        if (value === undefined || value === "") {
          const key = `${address.workbookName}:${address.sheetName}:${address.rowIndex}:${address.colIndex}`;
          const existing = generatedHeaders.get(key);
          if (existing && existing.name !== newName) {
            throw new Error("Overlapping tables require the same header name");
          }
          generatedHeaders.set(key, { address, name: newName });
        }
      }
    }

    return {
      updates,
      generatedHeaders: Array.from(generatedHeaders.values()),
    };
  }

  applyHeaderUpdates(updates: TableHeaderUpdate[]): void {
    return this.batchMutations(() => {
      const affectedTables = new Set(updates.map((update) => update.table));
      const before = this.observingMutations
        ? this.captureTableObjectStates(affectedTables)
        : undefined;

      for (const update of updates) {
        this.replaceHeader(update.table, update.index, update.newName);
      }

      if (before) {
        const after = this.captureTableObjectStates(affectedTables);
        const changes: TableMutation[] = [];
        for (const table of affectedTables) {
          const beforeState = before.get(table);
          const afterState = after.get(table);
          if (
            beforeState &&
            afterState &&
            tableEntryStatesEqual(beforeState, afterState)
          ) {
            continue;
          }
          changes.push({
            kind: "table",
            before: beforeState,
            after: afterState,
          });
        }
        this.reportMutations(changes);
      }
    });
  }

  makeTable({
    tableName,
    sheetName,
    workbookName,
    start,
    numRows,
    numCols,
    getCellValue,
  }: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    workbookName: string;
    getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  }): TableDefinition {
    const { rowIndex, colIndex } = parseCellReference(start);

    const headers = new Map<string, { name: string; index: number }>();
    const usedNames = new Set<string>();
    for (let i = 0; i < numCols; i++) {
      const header = getCellValue({
        rowIndex,
        colIndex: colIndex + i,
        sheetName,
        workbookName,
      });

      const name = this.getHeaderName(header, i, usedNames);
      if (usedNames.has(name)) {
        throw new Error(`Duplicate table header "${name}"`);
      }
      usedNames.add(name);
      headers.set(name, { name, index: i });
    }

    const endRow: SpreadsheetRangeEnd =
      numRows.type === "number"
        ? { type: "number", value: rowIndex + numRows.value }
        : numRows;

    const table: TableDefinition = {
      name: tableName,
      sheetName,
      workbookName,
      start: {
        rowIndex,
        colIndex,
      },
      headers,
      endRow,
    };

    return table;
  }

  copyTable(
    from: {
      workbookName: string;
      tableName: string;
    },
    to: {
      workbookName: string;
      tableName: string;
    }
  ): void {
    const fromTable = this.getTable({
      workbookName: from.workbookName,
      name: from.tableName,
    });
    if (!fromTable) {
      throw new Error("Table not found");
    }
    const wb = this.tables.get(to.workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }
    const before = this.observingMutations
      ? this.captureEntryState(to.workbookName, to.tableName)
      : undefined;
    const newTable: TableDefinition = {
      ...fromTable,
      workbookName: to.workbookName,
    };
    wb.set(to.tableName, newTable);
    if (this.observingMutations) {
      const after: TableEntryState = {
        workbookName: to.workbookName,
        tableName: to.tableName,
        table: cloneTableDefinition(newTable),
        index: before?.index ?? wb.size - 1,
      };
      if (!before || !tableEntryStatesEqual(before, after)) {
        this.reportMutations([{ kind: "table", before, after }]);
      }
    }
  }

  addTable(props: {
    tableName: string;
    sheetName: string;
    start: string;
    numRows: SpreadsheetRangeEnd;
    numCols: number;
    workbookName: string;
    getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  }): TableDefinition {
    const tableName = props.tableName;
    const table = this.makeTable(props);

    let wb = this.tables.get(props.workbookName);
    const bucketExisted = wb !== undefined;
    const before = this.observingMutations
      ? this.captureEntryState(props.workbookName, tableName)
      : undefined;
    if (!wb) {
      wb = new Map();
      this.tables.set(props.workbookName, wb);
    }

    wb.set(tableName, table);

    if (this.observingMutations) {
      const changes: TableMutation[] = [];
      if (!bucketExisted) {
        changes.push({
          kind: "table-workbook-bucket",
          before: undefined,
          after: {
            workbookName: props.workbookName,
            index: this.tables.size - 1,
          },
        });
      }
      const after: TableEntryState = {
        workbookName: props.workbookName,
        tableName,
        table: cloneTableDefinition(table),
        index: before?.index ?? wb.size - 1,
      };
      if (!before || !tableEntryStatesEqual(before, after)) {
        changes.push({ kind: "table", before, after });
      }
      this.reportMutations(changes);
    }

    return table;
  }

  renameTable(
    workbookName: string,
    names: { oldName: string; newName: string }
  ): void {
    const wb = this.tables.get(workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }
    const table = wb.get(names.oldName);
    if (!table) {
      throw new Error("Table not found");
    }
    const before = this.observingMutations
      ? this.captureEntryState(workbookName, names.oldName)
      : undefined;
    const overwritten =
      this.observingMutations && names.newName !== names.oldName
        ? this.captureEntryState(workbookName, names.newName)
        : undefined;
    table.name = names.newName;
    wb.set(names.newName, table);
    wb.delete(names.oldName);
    if (before) {
      const changes: TableMutation[] = [];
      if (names.newName === names.oldName) {
        changes.push({ kind: "table", before, after: undefined });
      } else {
        const after = this.captureEntryState(workbookName, names.newName);
        changes.push({
          kind: "table",
          before,
          after,
        });
        if (overwritten) {
          changes.push({
            kind: "table",
            before: overwritten,
            after: undefined,
          });
        }
      }
      this.reportMutations(changes);
    }
  }

  updateTable({
    tableName,
    sheetName,
    start,
    numRows,
    numCols,
    workbookName,
    getCellValue,
  }: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    workbookName: string;
    numCols?: number;
    getCellValue: (cellAddress: CellAddress) => SerializedCellValue;
  }): void {
    const wb = this.tables.get(workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }

    const table = wb.get(tableName);
    if (!table) {
      throw new Error("Table not found");
    }
    const before = this.observingMutations
      ? this.captureEntryState(workbookName, tableName)
      : undefined;

    const newStart = start ? parseCellReference(start) : table.start;

    let newNumRows: SpreadsheetRangeEnd;
    if (numRows) {
      newNumRows = numRows;
    } else {
      if (table.endRow.type === "infinity") {
        newNumRows = table.endRow;
      } else {
        newNumRows = {
          type: "number",
          value: table.endRow.value - newStart.rowIndex,
        };
      }
    }

    const newTable = this.makeTable({
      tableName,
      sheetName: sheetName ?? table.sheetName,
      workbookName: workbookName ?? table.workbookName,
      start: getCellReference(newStart),
      numRows: newNumRows,
      numCols: numCols ?? table.headers.size,
      getCellValue,
    });

    wb.set(tableName, newTable);
    if (before) {
      const after: TableEntryState = {
        workbookName,
        tableName,
        table: cloneTableDefinition(newTable),
        index: before.index,
      };
      if (!tableEntryStatesEqual(before, after)) {
        this.reportMutations([{ kind: "table", before, after }]);
      }
    }
  }

  removeTable({
    tableName,
    workbookName,
  }: {
    tableName: string;
    workbookName: string;
  }): boolean {
    const wb = this.tables.get(workbookName);
    if (!wb) {
      return false;
    }
    const before = this.observingMutations
      ? this.captureEntryState(workbookName, tableName)
      : undefined;
    const found = wb.delete(tableName);

    if (found && before) {
      this.reportMutations([{ kind: "table", before, after: undefined }]);
    }

    return found;
  }

  updateTablesForSheetRename(options: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    // Update tables that belong to the renamed sheet
    const wb = this.tables.get(options.workbookName);
    if (!wb) {
      // No tables exist for this workbook yet — nothing to update
      return;
    }
    this.batchMutations(() => {
      const changes: TableMutation[] = [];
      let index = 0;
      wb.forEach((table, tableName) => {
        if (table.sheetName === options.sheetName) {
          const before = this.observingMutations
            ? {
                workbookName: options.workbookName,
                tableName,
                table: cloneTableDefinition(table),
                index,
              }
            : undefined;
          table.sheetName = options.newSheetName;
          if (before) {
            changes.push({
              kind: "table",
              before,
              after: {
                workbookName: options.workbookName,
                tableName,
                table: cloneTableDefinition(table),
                index,
              },
            });
          }
        }
        index++;
      });
      this.reportMutations(changes);
    });
  }

  updateTablesForWorkbookRename(options: {
    workbookName: string;
    newWorkbookName: string;
  }): void {
    const wb = this.tables.get(options.workbookName);
    if (!wb) {
      // No tables exist for this workbook yet — nothing to update
      return;
    }
    const workbookNames = new Set([
      options.workbookName,
      options.newWorkbookName,
    ]);
    const before = this.observingMutations
      ? this.captureWorkbookState(workbookNames)
      : undefined;
    this.batchMutations(() => {
      this.tables.set(options.newWorkbookName, wb);
      this.tables.delete(options.workbookName);
      wb.forEach((table) => {
        if (table.workbookName === options.workbookName) {
          table.workbookName = options.newWorkbookName;
        }
      });
      if (before) {
        const after = this.captureWorkbookState(workbookNames);
        this.reportMutations(this.buildStateMutations(before, after));
      }
    });
  }

  resetTables(newTables: Map<string, Map<string, TableDefinition>>): void {
    return this.batchMutations(() =>
      this.mutateWithStateDiff(() => {
        this.tables.clear();
        newTables.forEach((workbookTables, workbookName) => {
          const restoredTables = new Map<string, TableDefinition>();
          workbookTables.forEach((table, tableName) => {
            restoredTables.set(tableName, table);
          });
          this.tables.set(workbookName, restoredTables);
        });
      })
    );
  }

  toSnapshot(): TableManagerSnapshot {
    return this.tables;
  }

  restoreFromSnapshot(snapshot: TableManagerSnapshot): void {
    this.resetTables(snapshot);
  }

  /**
   * When adding a workbook, we need to initialize the new maps
   */
  addWorkbook(workbookName: string) {
    const names = new Set([workbookName]);
    const before = this.observingMutations
      ? this.captureWorkbookState(names)
      : undefined;
    this.batchMutations(() => {
      this.tables.set(workbookName, new Map());
      if (before) {
        const after = this.captureWorkbookState(names);
        this.reportMutations(this.buildStateMutations(before, after));
      }
    });
  }

  /**
   * When removing a workbook, we need to remove the maps
   */
  removeWorkbook(workbookName: string) {
    const names = new Set([workbookName]);
    const before = this.observingMutations
      ? this.captureWorkbookState(names)
      : undefined;
    this.batchMutations(() => {
      this.tables.delete(workbookName);
      if (before) {
        const after = this.captureWorkbookState(names);
        this.reportMutations(this.buildStateMutations(before, after));
      }
    });
  }

  /**
   * When removing a sheet, we need to remove the tables that belong to the sheet
   */
  removeSheet(opts: { sheetName: string; workbookName: string }): void {
    // Remove tables that belong to the removed sheet
    const wb = this.tables.get(opts.workbookName);
    if (!wb) {
      throw new Error("Workbook not found");
    }
    this.batchMutations(() => {
      let index = 0;
      const removed: TableEntryState[] = [];
      wb.forEach((table, tableName) => {
        if (table.sheetName === opts.sheetName) {
          if (this.observingMutations) {
            removed.push({
              workbookName: opts.workbookName,
              tableName,
              table: cloneTableDefinition(table),
              index,
            });
          }
          wb.delete(tableName);
        }
        index++;
      });
      this.reportMutations(
        removed.map((before) => ({
          kind: "table" as const,
          before,
          after: undefined,
        }))
      );
    });
  }

  isCellInTable(cellAddress: CellAddress): TableDefinition | undefined {
    const { rowIndex, colIndex } = cellAddress;

    // Get all tables for this sheet

    for (const table of this.getTables(cellAddress.workbookName).values()) {
      // Check each table to see if the cell is within its bounds
      if (table.sheetName !== cellAddress.sheetName) {
        continue;
      }

      const { start, endRow, headers } = table;

      // Check row bounds
      const isInRowRange =
        endRow.type === "infinity"
          ? rowIndex >= start.rowIndex
          : rowIndex >= start.rowIndex && rowIndex <= endRow.value;

      // Check column bounds
      const endColIndex = start.colIndex + headers.size - 1;
      const isInColRange =
        colIndex >= start.colIndex && colIndex <= endColIndex;

      if (isInRowRange && isInColRange) {
        return table;
      }
    }

    return undefined;
  }

  /**
   * Check if a range intersects with any table in the given workbook/sheet.
   * Used to prevent spilling into tables (Excel behavior).
   */
  doesRangeIntersectTable(
    workbookName: string,
    sheetName: string,
    range: SpreadsheetRange
  ): boolean {
    for (const table of this.getTables(workbookName).values()) {
      if (table.sheetName !== sheetName) {
        continue;
      }

      // Build the table's range
      const { start, endRow, headers } = table;
      const endColIndex = start.colIndex + headers.size - 1;

      const tableRange: SpreadsheetRange = {
        start: { col: start.colIndex, row: start.rowIndex },
        end: {
          col: { type: "number", value: endColIndex },
          row: endRow,
        },
      };

      if (checkRangeIntersection(range, tableRange)) {
        return true;
      }
    }

    return false;
  }
}
