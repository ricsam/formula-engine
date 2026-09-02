/**
 * Main FormulaEngine class
 * Core API implementation for spreadsheet calculations
 */

import {
  type CellAddress,
  type CellDataType,
  type CellStyle,
  type ConditionalStyle,
  type CopyCellsOptions,
  type DirectCellDataType,
  type DirectCellStyle,
  type FormulaEngineOptions,
  type NamedExpression,
  type RangeAddress,
  type RangeMetadata,
  type RangeMetadataInput,
  type ReplaceChange,
  type ReplaceTarget,
  type SearchMatch,
  type SearchOptions,
  type SerializedCellValue,
  type Sheet,
  type SingleEvaluationResult,
  type SpreadsheetRange,
  type SpreadsheetRangeEnd,
  type TableDefinition,
  type UndoRedoState,
} from "./types";

import type { FillDirection } from "@ricsam/selection-manager";
import { FormulaEvaluator } from "../evaluator/formula-evaluator";
import { AutoFill } from "./autofill-utils";
import {
  WorkbookManager,
  type WorkbookDataChange,
} from "./managers/workbook-manager";
import { deserialize, serialize } from "./map-serializer";
import { renameNamedExpressionInFormula } from "./named-expression-renamer";
import { renameSheetInFormula } from "./sheet-renamer";
import {
  renameTableColumnsInFormula,
  renameTableInFormula,
} from "./table-renamer";
import { renameWorkbookInFormula } from "./workbook-renamer";
import { getCellReference, parseCellReference } from "./utils";
import { CacheManager } from "./managers/cache-manager";
import {
  NamedExpressionManager,
  type NamedExpressionMutation,
} from "./managers/named-expression-manager";
import {
  TableManager,
  type TableHeaderUpdate,
  type TableMutation,
} from "./managers/table-manager";
import {
  EventManager,
  type ResourceEvent,
} from "./managers/event-manager";
import { EvaluationManager } from "./managers/evaluation-manager";
import { DependencyManager } from "./managers/dependency-manager";
import { StyleManager, type StyleDataChange } from "./managers/style-manager";
import { CopyManager } from "./managers/copy-manager";
import {
  ReferenceManager,
  type ReferenceDataChange,
} from "./managers/reference-manager";
import {
  RangeMetadataManager,
  type RangeMetadataDataChange,
} from "./managers/range-metadata-manager";
import { UndoRedoManager } from "./managers/undo-redo-manager";
import {
  createHistoryEntry,
  estimateHistoryValueBytes,
  isHistoryValueSafelyRetainable,
  type HistoryEntry,
} from "./history";
import {
  cloneHistoryValue,
  historyValuesEqual,
  type EngineHistoryStep,
  type SheetScopeState,
  type WorkbookScopeState,
} from "./engine-history";
import {
  ENGINE_SNAPSHOT_VERSION,
  LEGACY_ENGINE_SNAPSHOT_VERSION,
  type EngineHistorySnapshot,
  type EngineSnapshot,
  type StyleManagerSnapshot,
} from "./engine-snapshot";
import {
  buildFormulaTouchedCells,
  buildTableContextChangedCells,
  buildTableTouchedCells,
  buildTouchedCells,
  getNamedExpressionScopeResourceKeys,
  mergeTouchedCells,
  type MutationInvalidation,
} from "./mutation-invalidation";
import {
  getNamedExpressionResourceKey,
  getSheetResourceKey,
  getTableResourceKey,
  getWorkbookResourceKey,
} from "./resource-keys";

type Metadata = {
  cell?: unknown;
  range?: unknown;
  sheet?: unknown;
  workbook?: unknown;
};

type MetadataType<
  TMetadata extends Metadata,
  TKey extends keyof Metadata
> = TMetadata[TKey];

const MAX_OBSERVER_INVALIDATION_KEYS = 4_096;

function isPromiseLike(value: unknown): value is PromiseLike<unknown> {
  if (
    (typeof value !== "object" || value === null) &&
    typeof value !== "function"
  ) {
    return false;
  }
  try {
    return typeof (value as { then?: unknown }).then === "function";
  } catch {
    return true;
  }
}

class HistoryTransactionCapacityError extends Error {
  constructor() {
    super(
      "FormulaEngine.transact exceeded undoRedo.maxBytes or used metadata that cannot be retained safely"
    );
    this.name = "HistoryTransactionCapacityError";
  }
}

/**
 * Main FormulaEngine class
 * @template TMetadata - Consumer-defined metadata shape with optional cell, sheet, and workbook entries.
 */
export class FormulaEngine<TMetadata extends Metadata = Metadata> {
  private workbookManager: WorkbookManager;
  private namedExpressionManager: NamedExpressionManager;
  private tableManager: TableManager;
  private eventManager: EventManager;
  private evaluationManager: EvaluationManager;
  private autoFillManager: AutoFill;
  private dependencyManager: DependencyManager;
  private styleManager: StyleManager;
  private rangeMetadataManager: RangeMetadataManager<
    MetadataType<TMetadata, "range">
  >;
  private copyManager: CopyManager;
  private referenceManager: ReferenceManager;
  private undoRedoManager: UndoRedoManager<
    EngineHistoryStep<MetadataType<TMetadata, "range">>
  >;
  private historyTransactionDepth = 0;
  private pendingHistorySteps: EngineHistoryStep<
    MetadataType<TMetadata, "range">
  >[] = [];
  private pendingWorkbookChanges = new Map<string, WorkbookDataChange>();
  private pendingWorkbookChangeBytes = new Map<string, number>();
  private pendingHistoryEstimatedBytes = 0;
  private pendingHistoryOversized = false;
  private pendingResourceEvents: ResourceEvent[] = [];
  private observerInvalidatedCellKeys = new Set<string>();
  private observerInvalidationDeduplicationDisabled = false;
  private pendingUpdate = false;
  private isReplayingHistory = false;
  private workbookHistoryCaptureSuppressionDepth = 0;
  private workbookObserverInvalidationSuppressionDepth = 0;
  private explicitTransactionDepth = 0;

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
  public _rangeMetadataManager: RangeMetadataManager<
    MetadataType<TMetadata, "range">
  >;

  constructor(options: FormulaEngineOptions = {}) {
    this.undoRedoManager = new UndoRedoManager(options.undoRedo);
    this.eventManager = new EventManager();
    this.workbookManager = new WorkbookManager(
      (patches) => {
        for (const patch of patches) {
          if (patch.committed && patch.atomicGroupId !== undefined) {
            this.commitPendingWorkbookDataGroup(patch.atomicGroupId);
          } else {
            this.captureWorkbookDataChanges(patch.changes, patch.atomicGroupId);
          }
        }
      },
      () => this.shouldObserveWorkbookMutations(),
      () => this.shouldBatchMutationNotifications()
    );
    this.namedExpressionManager = new NamedExpressionManager(
      (changes) => {
        this.captureNamedExpressionDataChanges(changes);
      },
      () => this.shouldCaptureAncillaryHistory(),
      false,
      () => this.shouldBatchMutationNotifications()
    );
    this.tableManager = new TableManager(
      this.workbookManager,
      (changes) => {
        this.captureTableDataChanges(changes);
      },
      () => this.shouldCaptureAncillaryHistory(),
      false,
      () => this.shouldBatchMutationNotifications()
    );
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

    this.styleManager = new StyleManager(
      this.evaluationManager,
      (changes) => {
        this.captureStyleDataChanges(changes);
      },
      () => this.shouldCaptureAncillaryHistory(),
      false
    );
    this.rangeMetadataManager = new RangeMetadataManager<
      MetadataType<TMetadata, "range">
    >(
      (changes) => {
        this.captureRangeMetadataDataChanges(changes);
      },
      () => this.shouldCaptureAncillaryHistory(),
      false
    );
    this.copyManager = new CopyManager(
      this.workbookManager,
      this.evaluationManager,
      this.styleManager,
      this.rangeMetadataManager
    );

    this.autoFillManager = new AutoFill(
      this.workbookManager,
      this.styleManager,
      this.rangeMetadataManager
    );

    this.referenceManager = new ReferenceManager(
      (changes) => {
        this.captureReferenceDataChanges(changes);
      },
      () => this.shouldCaptureAncillaryHistory(),
      false
    );

    this._workbookManager = this.workbookManager;
    this._namedExpressionManager = this.namedExpressionManager;
    this._tableManager = this.tableManager;
    this._eventManager = this.eventManager;
    this._evaluationManager = this.evaluationManager;
    this._autoFillManager = this.autoFillManager;
    this._dependencyManager = this.dependencyManager;
    this._styleManager = this.styleManager;
    this._rangeMetadataManager = this.rangeMetadataManager;
  }

  /**
   * Static factory method to build an empty engine
   * @template TMetadata - Consumer-defined metadata shape with optional cell, sheet, and workbook entries.
   */
  static buildEmpty<TMetadata extends Metadata = Metadata>(
    options?: FormulaEngineOptions
  ) {
    return new FormulaEngine<TMetadata>(options);
  }

  undo(): boolean {
    this.assertHistoryControlAllowed("undo");
    if (!this.undoRedoManager.canUndo()) {
      return false;
    }

    const entry = this.undoRedoManager.popUndo();
    if (!entry) {
      return false;
    }

    let resourceEvents: ResourceEvent[];
    try {
      resourceEvents = this.replayHistoryEntry(entry, "undo");
    } catch (error) {
      this.undoRedoManager.pushUndoFromReplay(entry);
      throw error;
    }
    this.undoRedoManager.pushRedoFromReplay(entry);
    this.eventManager.emitResourceEvents(resourceEvents);
    this.eventManager.emitUpdate();
    return true;
  }

  redo(): boolean {
    this.assertHistoryControlAllowed("redo");
    if (!this.undoRedoManager.canRedo()) {
      return false;
    }

    const entry = this.undoRedoManager.popRedo();
    if (!entry) {
      return false;
    }

    let resourceEvents: ResourceEvent[];
    try {
      resourceEvents = this.replayHistoryEntry(entry, "redo");
    } catch (error) {
      this.undoRedoManager.pushRedoFromReplay(entry);
      throw error;
    }
    this.undoRedoManager.pushUndoFromReplay(entry);
    this.eventManager.emitResourceEvents(resourceEvents);
    this.eventManager.emitUpdate();
    return true;
  }

  canUndo(): boolean {
    return this.undoRedoManager.canUndo();
  }

  canRedo(): boolean {
    return this.undoRedoManager.canRedo();
  }

  getUndoRedoState(): UndoRedoState {
    return this.undoRedoManager.getState();
  }

  clearUndoRedoHistory(): void {
    this.assertHistoryControlAllowed("clear undo/redo history");
    this.undoRedoManager.clear();
  }

  /** Groups synchronous mutations into one atomic undo entry and update. */
  transact<TCallback extends () => unknown>(
    callback: [ReturnType<TCallback>] extends [never]
      ? TCallback
      : ReturnType<TCallback> extends PromiseLike<unknown>
      ? never
      : TCallback
  ): ReturnType<TCallback> {
    this.explicitTransactionDepth++;
    let result: ReturnType<TCallback>;
    try {
      result = this.withUndoRedoCheckpoint(() =>
        this.styleManager.batchMutations(() =>
          this.rangeMetadataManager.batchMutations(() =>
            this.referenceManager.batchMutations(callback)
          )
        )
      ) as ReturnType<TCallback>;
    } catch (error) {
      this.explicitTransactionDepth--;
      throw error;
    }

    if (isPromiseLike(result)) {
      return Promise.resolve(result).finally(() => {
        this.explicitTransactionDepth--;
      }) as ReturnType<TCallback>;
    }
    this.explicitTransactionDepth--;
    return result;
  }

  private emitMutation(footprint: MutationInvalidation): void {
    const touchedCells = this.observerInvalidationDeduplicationDisabled
      ? footprint.touchedCells
      : footprint.touchedCells.filter(
          ({ address }) =>
            !this.observerInvalidatedCellKeys.has(
              this.getHistoryAddressKey(address)
            )
        );
    if (
      touchedCells.length > 0 ||
      footprint.resourceKeys.length > 0 ||
      (footprint.tableContextChangedCells?.length ?? 0) > 0 ||
      (footprint.removedScopes?.length ?? 0) > 0
    ) {
      this.evaluationManager.invalidateFromMutation({
        ...footprint,
        touchedCells,
      });
    }
    this.observerInvalidatedCellKeys.clear();
    this.observerInvalidationDeduplicationDisabled = false;
    this.requestUpdate();
  }

  private emitUpdate(): void {
    this.requestUpdate();
  }

  private withUndoRedoCheckpoint<T>(callback: () => T): T {
    if (this.isReplayingHistory) {
      return callback();
    }

    const isOuterTransaction = this.historyTransactionDepth === 0;
    if (isOuterTransaction) {
      this.pendingHistorySteps = [];
      this.pendingWorkbookChanges.clear();
      this.pendingWorkbookChangeBytes.clear();
      this.pendingHistoryEstimatedBytes = 0;
      this.pendingHistoryOversized = false;
      this.pendingResourceEvents = [];
      this.observerInvalidatedCellKeys.clear();
      this.observerInvalidationDeduplicationDisabled = false;
      this.pendingUpdate = false;
    }

    this.historyTransactionDepth++;
    let result: T;
    try {
      result = callback();
    } catch (error) {
      this.abortUndoRedoCheckpoint(isOuterTransaction);
      throw error;
    }

    if (isPromiseLike(result)) {
      return Promise.resolve(result).then(
        () => {
          this.abortUndoRedoCheckpoint(isOuterTransaction);
          throw new Error(
            "FormulaEngine.transact callback must be synchronous"
          );
        },
        (error) => {
          this.abortUndoRedoCheckpoint(isOuterTransaction);
          throw error;
        }
      ) as T;
    }

    this.historyTransactionDepth--;
    if (isOuterTransaction) {
      this.flushPendingWorkbookChanges();
      if (this.pendingHistoryOversized) {
        this.undoRedoManager.recordBarrier();
      } else if (this.pendingHistorySteps.length > 0) {
        const entry = createHistoryEntry(this.pendingHistorySteps);
        if (
          this.explicitTransactionDepth > 0 &&
          entry.estimatedBytes > this.undoRedoManager.maxBytes
        ) {
          try {
            this.rollbackPendingHistory();
          } finally {
            this.resetPendingHistoryTransaction();
          }
          throw new HistoryTransactionCapacityError();
        }
        this.undoRedoManager.record(entry);
      }
      const resourceEvents = this.pendingResourceEvents;
      const shouldEmitUpdate = this.pendingUpdate;
      this.resetPendingHistoryTransaction();
      this.eventManager.emitResourceEvents(resourceEvents);
      if (shouldEmitUpdate) {
        this.eventManager.emitUpdate();
      }
    }

    return result;
  }

  private abortUndoRedoCheckpoint(isOuterTransaction: boolean): void {
    this.historyTransactionDepth--;
    if (!isOuterTransaction) {
      return;
    }
    this.flushPendingWorkbookChanges();
    if (this.pendingHistoryOversized) {
      // The inverse was intentionally discarded to honor the memory budget,
      // so older history must not cross the committed mutation.
      this.undoRedoManager.recordBarrier();
    } else {
      this.rollbackPendingHistory();
    }
    this.resetPendingHistoryTransaction();
  }

  private requestUpdate(): void {
    if (this.historyTransactionDepth > 0 || this.isReplayingHistory) {
      this.pendingUpdate = true;
      return;
    }
    this.eventManager.emitUpdate();
  }

  private shouldCaptureAncillaryHistory(): boolean {
    return (
      !this.isReplayingHistory &&
      this.historyTransactionDepth > 0 &&
      !this.pendingHistoryOversized
    );
  }

  private shouldObserveWorkbookMutations(): boolean {
    return (
      !this.isReplayingHistory &&
      this.historyTransactionDepth > 0 &&
      this.workbookHistoryCaptureSuppressionDepth === 0
    );
  }

  private shouldBatchMutationNotifications(): boolean {
    return this.explicitTransactionDepth === 0;
  }

  private assertHistoryControlAllowed(action: string): void {
    if (this.historyTransactionDepth > 0) {
      throw new Error(`Cannot ${action} during a transaction`);
    }
  }

  private resetPendingHistoryTransaction(): void {
    this.pendingHistorySteps = [];
    this.pendingWorkbookChanges.clear();
    this.pendingWorkbookChangeBytes.clear();
    this.pendingHistoryEstimatedBytes = 0;
    this.pendingHistoryOversized = false;
    this.pendingResourceEvents = [];
    this.observerInvalidatedCellKeys.clear();
    this.observerInvalidationDeduplicationDisabled = false;
    this.pendingUpdate = false;
  }

  private markPendingHistoryOversized(): void {
    this.pendingHistoryOversized = true;
    this.pendingHistorySteps = [];
    this.pendingWorkbookChanges.clear();
    this.pendingWorkbookChangeBytes.clear();
    this.pendingHistoryEstimatedBytes = 0;
  }

  private rejectExplicitTransactionOrMarkOversized(
    appliedStep?: EngineHistoryStep<MetadataType<TMetadata, "range">>
  ): void {
    if (this.explicitTransactionDepth === 0) {
      this.markPendingHistoryOversized();
      return;
    }

    if (appliedStep) {
      const wasReplaying = this.isReplayingHistory;
      this.isReplayingHistory = true;
      try {
        this.applyHistoryStep(appliedStep, "undo");
      } finally {
        this.isReplayingHistory = wasReplaying;
      }
    }
    throw new HistoryTransactionCapacityError();
  }

  private cloneMetadataHistoryValue(value: unknown): unknown {
    if (value === null || typeof value !== "object") {
      return value;
    }
    return cloneHistoryValue(value);
  }

  private cloneWorkbookDataChange(
    change: WorkbookDataChange
  ): WorkbookDataChange {
    if (change.kind === "cell-content") {
      return {
        ...change,
        address: { ...change.address },
      };
    }
    if (change.kind === "cell-metadata") {
      return {
        ...change,
        address: { ...change.address },
        before: this.cloneMetadataHistoryValue(change.before),
        after: this.cloneMetadataHistoryValue(change.after),
      };
    }
    return {
      ...change,
      before: this.cloneMetadataHistoryValue(change.before),
      after: this.cloneMetadataHistoryValue(change.after),
    };
  }

  private getWorkbookDataChangeKey(change: WorkbookDataChange): string {
    switch (change.kind) {
      case "cell-content":
      case "cell-metadata":
        return JSON.stringify([
          change.kind,
          change.address.workbookName,
          change.address.sheetName,
          change.address.rowIndex,
          change.address.colIndex,
        ]);
      case "sheet-metadata":
        return JSON.stringify([
          change.kind,
          change.workbookName,
          change.sheetName,
        ]);
      case "workbook-metadata":
        return JSON.stringify([change.kind, change.workbookName]);
    }
  }

  private getHistoryAddressKey(address: CellAddress): string {
    return JSON.stringify([
      address.workbookName,
      address.sheetName,
      address.rowIndex,
      address.colIndex,
    ]);
  }

  private workbookDataValuesEqual(
    change: WorkbookDataChange,
    before: unknown,
    after: unknown
  ): boolean {
    if (change.kind === "cell-content") {
      return (
        Object.is(before, after) &&
        Object.is(change.beforeIndex, change.afterIndex)
      );
    }
    if (change.kind === "cell-metadata") {
      return (
        historyValuesEqual(before, after) &&
        Object.is(change.beforeIndex, change.afterIndex)
      );
    }
    return historyValuesEqual(before, after);
  }

  private isAppendOnlyCellContentPatch(
    changes: readonly WorkbookDataChange[]
  ): boolean {
    return (
      changes.length > 0 &&
      changes.every(
        (change) =>
          change.kind === "cell-content" &&
          change.before === undefined &&
          change.after !== undefined &&
          change.beforeIndex === undefined &&
          change.afterIndex !== undefined
      )
    );
  }

  private isSequentialCellContentDeletionPatch(
    changes: readonly WorkbookDataChange[]
  ): boolean {
    const change = changes[0];
    return (
      changes.length === 1 &&
      change?.kind === "cell-content" &&
      change.before !== undefined &&
      change.after === undefined &&
      change.beforeIndex !== undefined &&
      change.afterIndex === undefined
    );
  }

  private captureWorkbookDataChanges(
    changes: readonly WorkbookDataChange[],
    atomicGroupId?: number
  ): void {
    if (
      this.isReplayingHistory ||
      this.workbookHistoryCaptureSuppressionDepth > 0 ||
      this.historyTransactionDepth === 0
    ) {
      return;
    }

    const contentChanges = changes.flatMap((change) =>
      change.kind === "cell-content"
        ? [
            {
              address: change.address,
              before: change.before,
              after: change.after,
            },
          ]
        : []
    );
    if (
      contentChanges.length > 0 &&
      this.workbookObserverInvalidationSuppressionDepth === 0
    ) {
      if (!this.observerInvalidationDeduplicationDisabled) {
        for (const { address } of contentChanges) {
          if (
            this.observerInvalidatedCellKeys.size >=
            MAX_OBSERVER_INVALIDATION_KEYS
          ) {
            this.observerInvalidatedCellKeys.clear();
            this.observerInvalidationDeduplicationDisabled = true;
            break;
          }
          this.observerInvalidatedCellKeys.add(
            this.getHistoryAddressKey(address)
          );
        }
      }
      this.evaluationManager.invalidateFromMutation({
        touchedCells: buildTouchedCells(contentChanges),
        resourceKeys: [],
      });
    }

    if (this.pendingHistoryOversized) {
      return;
    }

    const rawWorkbookStep: EngineHistoryStep<MetadataType<TMetadata, "range">> =
      {
        kind: "workbook-data",
        patches: [[...changes]],
      };
    if (!isHistoryValueSafelyRetainable(changes)) {
      this.rejectExplicitTransactionOrMarkOversized(
        atomicGroupId === undefined ? rawWorkbookStep : undefined
      );
      return;
    }

    // Collection indexes are relative to one observer patch. Keep indexed
    // patches atomic instead of coalescing them with later mutations: two
    // sequential deletions can legitimately both report index zero.
    if (
      changes.some(
        (change) =>
          (change.kind === "cell-content" || change.kind === "cell-metadata") &&
          (change.beforeIndex !== undefined || change.afterIndex !== undefined)
      )
    ) {
      this.flushPendingWorkbookChanges();
      const appendOnlyCellContent =
        atomicGroupId === undefined &&
        this.isAppendOnlyCellContentPatch(changes);
      const sequentialCellContentDeletions =
        atomicGroupId === undefined &&
        this.isSequentialCellContentDeletionPatch(changes);
      const previous = this.pendingHistorySteps.at(-1);
      const canAppendToPrevious =
        previous?.kind === "workbook-data" &&
        ((appendOnlyCellContent && previous.appendOnlyCellContent) ||
          (sequentialCellContentDeletions &&
            previous.sequentialCellContentDeletions));
      if (
        !canAppendToPrevious &&
        this.pendingHistoryEstimatedBytes +
          estimateHistoryValueBytes({
            kind: "workbook-data",
            patches: [changes],
            ...(appendOnlyCellContent
              ? { appendOnlyCellContent: true as const }
              : {}),
            ...(sequentialCellContentDeletions
              ? { sequentialCellContentDeletions: true as const }
              : {}),
          }) >
          this.undoRedoManager.maxBytes
      ) {
        this.rejectExplicitTransactionOrMarkOversized(
          atomicGroupId === undefined ? rawWorkbookStep : undefined
        );
        return;
      }
      const patch = changes.map((change) =>
        this.cloneWorkbookDataChange(change)
      );
      if (atomicGroupId === undefined) {
        if (canAppendToPrevious) {
          // Account exactly for appending one element to the existing patches
          // array without re-walking the growing step. Numeric array property
          // keys add two bytes per decimal digit in the estimator.
          const fragmentBytes =
            estimateHistoryValueBytes([patch]) -
            estimateHistoryValueBytes([]) +
            (String(previous.patches.length).length - 1) * 2;
          if (
            this.pendingHistoryEstimatedBytes + fragmentBytes >
            this.undoRedoManager.maxBytes
          ) {
            this.rejectExplicitTransactionOrMarkOversized(rawWorkbookStep);
            return;
          }
          previous.patches.push(patch);
          this.pendingHistoryEstimatedBytes += fragmentBytes;
        } else {
          this.recordHistoryStep({
            kind: "workbook-data",
            patches: [patch],
            ...(appendOnlyCellContent
              ? { appendOnlyCellContent: true as const }
              : {}),
            ...(sequentialCellContentDeletions
              ? { sequentialCellContentDeletions: true as const }
              : {}),
          });
        }
      } else {
        const fragmentBytes = estimateHistoryValueBytes({
          kind: "workbook-data",
          patches: [patch],
        });
        if (
          this.pendingHistoryEstimatedBytes + fragmentBytes >
          this.undoRedoManager.maxBytes
        ) {
          this.rejectExplicitTransactionOrMarkOversized();
          return;
        }

        const previous = this.pendingHistorySteps.at(-1);
        if (
          previous?.kind === "workbook-data" &&
          previous.atomicGroupId === atomicGroupId
        ) {
          previous.patches.push(patch);
        } else {
          this.pendingHistorySteps.push({
            kind: "workbook-data",
            patches: [patch],
            atomicGroupId,
            committed: false,
          });
        }
        this.pendingHistoryEstimatedBytes += fragmentBytes;
      }
      return;
    }

    for (const rawChange of changes) {
      const key = this.getWorkbookDataChangeKey(rawChange);
      const existing = this.pendingWorkbookChanges.get(key);
      const existingBytes = this.pendingWorkbookChangeBytes.get(key) ?? 0;
      if (
        this.pendingHistoryEstimatedBytes -
          existingBytes +
          estimateHistoryValueBytes(rawChange) >
        this.undoRedoManager.maxBytes
      ) {
        this.rejectExplicitTransactionOrMarkOversized({
          kind: "workbook-data",
          patches: [[rawChange]],
        });
        return;
      }

      const change = this.cloneWorkbookDataChange(rawChange);
      const before = existing ? existing.before : change.before;
      let merged = { ...change, before } as WorkbookDataChange;
      if (
        (change.kind === "cell-content" || change.kind === "cell-metadata") &&
        existing &&
        (existing.kind === "cell-content" || existing.kind === "cell-metadata")
      ) {
        const beforeIndex = existing.beforeIndex ?? change.beforeIndex;
        const afterIndex =
          change.afterIndex ??
          (change.after === undefined ? undefined : existing.afterIndex);
        merged = {
          ...change,
          before,
          ...(beforeIndex === undefined ? {} : { beforeIndex }),
          ...(afterIndex === undefined ? {} : { afterIndex }),
        } as WorkbookDataChange;
      }
      if (this.workbookDataValuesEqual(merged, before, merged.after)) {
        this.pendingWorkbookChanges.delete(key);
        this.pendingWorkbookChangeBytes.delete(key);
        this.pendingHistoryEstimatedBytes -= existingBytes;
      } else {
        const mergedBytes = estimateHistoryValueBytes(merged);
        const nextEstimatedBytes =
          this.pendingHistoryEstimatedBytes - existingBytes + mergedBytes;
        if (nextEstimatedBytes > this.undoRedoManager.maxBytes) {
          this.rejectExplicitTransactionOrMarkOversized({
            kind: "workbook-data",
            patches: [[rawChange]],
          });
          return;
        }
        this.pendingWorkbookChanges.set(key, merged);
        this.pendingWorkbookChangeBytes.set(key, mergedBytes);
        this.pendingHistoryEstimatedBytes = nextEstimatedBytes;
      }
    }
  }

  private captureNamedExpressionDataChanges(
    changes: readonly NamedExpressionMutation[]
  ): void {
    this.recordRawHistoryStep({
      kind: "named-expression-data",
      changes: [...changes],
    });
  }

  private captureTableDataChanges(changes: readonly TableMutation[]): void {
    this.recordRawHistoryStep({ kind: "table-data", changes: [...changes] });
  }

  private captureStyleDataChanges(changes: readonly StyleDataChange[]): void {
    this.recordRawHistoryStep({ kind: "style-data", changes: [...changes] });
  }

  private captureRangeMetadataDataChanges(
    changes: readonly RangeMetadataDataChange<
      MetadataType<TMetadata, "range">
    >[]
  ): void {
    this.recordRawHistoryStep({
      kind: "range-metadata-data",
      changes: [...changes],
    });
  }

  private captureReferenceDataChanges(
    changes: readonly ReferenceDataChange[]
  ): void {
    this.recordRawHistoryStep({
      kind: "reference-data",
      changes: [...changes],
    });
  }

  private flushPendingWorkbookChanges(): void {
    if (this.pendingWorkbookChanges.size === 0) {
      return;
    }
    this.pendingHistorySteps.push({
      kind: "workbook-data",
      patches: [Array.from(this.pendingWorkbookChanges.values())],
    });
    this.pendingWorkbookChanges.clear();
    this.pendingWorkbookChangeBytes.clear();
  }

  private commitPendingWorkbookDataGroup(atomicGroupId: number): void {
    for (let index = this.pendingHistorySteps.length - 1; index >= 0; index--) {
      const step = this.pendingHistorySteps[index];
      if (
        step?.kind === "workbook-data" &&
        step.atomicGroupId === atomicGroupId
      ) {
        step.committed = true;
        return;
      }
    }
  }

  private recordHistoryStep(
    step: EngineHistoryStep<MetadataType<TMetadata, "range">>
  ): void {
    if (
      this.isReplayingHistory ||
      this.historyTransactionDepth === 0 ||
      this.pendingHistoryOversized
    ) {
      return;
    }
    if (!isHistoryValueSafelyRetainable(step)) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }
    this.flushPendingWorkbookChanges();
    const stepBytes = estimateHistoryValueBytes(step);
    if (
      this.pendingHistoryEstimatedBytes + stepBytes >
      this.undoRedoManager.maxBytes
    ) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }
    this.pendingHistoryEstimatedBytes += stepBytes;
    this.pendingHistorySteps.push(step);
  }

  /** Estimate manager-owned raw deltas before paying to detach them. */
  private recordRawHistoryStep(
    step: EngineHistoryStep<MetadataType<TMetadata, "range">>
  ): void {
    if (
      this.isReplayingHistory ||
      this.historyTransactionDepth === 0 ||
      this.pendingHistoryOversized
    ) {
      return;
    }
    if (!isHistoryValueSafelyRetainable(step)) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }

    this.flushPendingWorkbookChanges();
    const stepBytes = estimateHistoryValueBytes(step);
    if (
      this.pendingHistoryEstimatedBytes + stepBytes >
      this.undoRedoManager.maxBytes
    ) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }

    const detached = cloneHistoryValue(step);
    if (
      typeof step === "object" &&
      step !== null &&
      Object.is(detached, step)
    ) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }
    this.pendingHistoryEstimatedBytes += stepBytes;
    this.pendingHistorySteps.push(detached);
  }

  private recordHistoryStepBeforePendingWorkbookChanges(
    step: EngineHistoryStep<MetadataType<TMetadata, "range">>
  ): void {
    if (
      this.isReplayingHistory ||
      this.historyTransactionDepth === 0 ||
      this.pendingHistoryOversized
    ) {
      return;
    }
    if (!isHistoryValueSafelyRetainable(step)) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }
    const stepBytes = estimateHistoryValueBytes(step);
    if (
      this.pendingHistoryEstimatedBytes + stepBytes >
      this.undoRedoManager.maxBytes
    ) {
      this.rejectExplicitTransactionOrMarkOversized(step);
      return;
    }
    this.pendingHistoryEstimatedBytes += stepBytes;
    this.pendingHistorySteps.push(step);
  }

  private withoutWorkbookHistoryCapture<T>(callback: () => T): T {
    this.flushPendingWorkbookChanges();
    this.workbookHistoryCaptureSuppressionDepth++;
    try {
      return callback();
    } finally {
      this.workbookHistoryCaptureSuppressionDepth--;
    }
  }

  private rollbackPendingHistory(): void {
    if (this.pendingHistorySteps.length === 0) {
      return;
    }

    this.isReplayingHistory = true;
    try {
      for (
        let index = this.pendingHistorySteps.length - 1;
        index >= 0;
        index--
      ) {
        const step = this.pendingHistorySteps[index]!;
        if (step.kind === "workbook-data" && step.committed === false) {
          continue;
        }
        this.applyHistoryStep(step, "undo");
      }
      this.evaluationManager.clearEvaluationCache();
    } finally {
      this.isReplayingHistory = false;
      this.pendingUpdate = false;
    }
  }

  private replayHistoryEntry(
    entry: HistoryEntry<EngineHistoryStep<MetadataType<TMetadata, "range">>>,
    direction: "undo" | "redo"
  ): ResourceEvent[] {
    this.isReplayingHistory = true;
    let requiresFullInvalidation = false;
    const resourceEvents: ResourceEvent[] = [];
    try {
      const steps =
        direction === "undo" ? [...entry.steps].reverse() : entry.steps;
      for (const step of steps) {
        requiresFullInvalidation =
          this.applyHistoryStep(step, direction) || requiresFullInvalidation;
        resourceEvents.push(
          ...this.getHistoryResourceEvents(step, direction)
        );
      }
      if (requiresFullInvalidation) {
        this.evaluationManager.clearEvaluationCache();
      }
      return resourceEvents;
    } finally {
      this.isReplayingHistory = false;
      this.pendingUpdate = false;
    }
  }

  private getHistoryResourceEvents(
    step: EngineHistoryStep<MetadataType<TMetadata, "range">>,
    direction: "undo" | "redo"
  ): ResourceEvent[] {
    const useAfter = direction === "redo";
    if (step.kind === "rename-sheet") {
      return [
        {
          type: "sheet-rename",
          workbookName: step.workbookName,
          sheetName: useAfter ? step.before : step.after,
          newSheetName: useAfter ? step.after : step.before,
        },
      ];
    }
    if (step.kind === "rename-workbook") {
      return [
        {
          type: "workbook-rename",
          workbookName: useAfter ? step.before : step.after,
          newWorkbookName: useAfter ? step.after : step.before,
        },
      ];
    }
    if (step.kind === "workbook-scope") {
      const source = useAfter ? step.before : step.after;
      const target = useAfter ? step.after : step.before;
      if (source.workbook !== undefined && target.workbook === undefined) {
        return [
          ...Array.from(source.workbook.sheets.keys(), (sheetName) => ({
            type: "sheet-delete" as const,
            workbookName: source.workbookName,
            sheetName,
          })),
          {
            type: "workbook-delete",
            workbookName: source.workbookName,
          },
        ];
      }
    }
    if (step.kind === "sheet-scope") {
      const source = useAfter ? step.before : step.after;
      const target = useAfter ? step.after : step.before;
      if (source.sheet !== undefined && target.sheet === undefined) {
        return [
          {
            type: "sheet-delete",
            workbookName: source.workbookName,
            sheetName: source.sheetName,
          },
        ];
      }
    }
    return [];
  }

  private applyHistoryStep(
    step: EngineHistoryStep<MetadataType<TMetadata, "range">>,
    direction: "undo" | "redo"
  ): boolean {
    const useAfter = direction === "redo";
    switch (step.kind) {
      case "workbook-data": {
        const contentDataChanges = function* (): Generator<
          Extract<WorkbookDataChange, { kind: "cell-content" }>
        > {
          for (const patch of step.patches) {
            for (const change of patch) {
              if (change.kind === "cell-content") {
                yield change;
              }
            }
          }
        };
        if (step.sequentialCellContentDeletions && step.patches.length > 1) {
          this.workbookManager.applySequentialCellContentDeletionsForHistory(
            step.patches,
            direction
          );
        } else {
          this.workbookManager.applyCellContentChangesForHistory(
            contentDataChanges(),
            direction
          );
        }
        const cellMetadataChanges = function* (): Generator<
          Extract<WorkbookDataChange, { kind: "cell-metadata" }>
        > {
          for (const patch of step.patches) {
            for (const change of patch) {
              if (change.kind === "cell-metadata") {
                yield change;
              }
            }
          }
        };
        this.workbookManager.applyCellMetadataChangesForHistory(
          cellMetadataChanges(),
          direction,
          (value) => this.cloneMetadataHistoryValue(value)
        );

        const patches =
          direction === "undo" ? [...step.patches].reverse() : step.patches;
        for (const patch of patches) {
          const changes = direction === "undo" ? [...patch].reverse() : patch;
          for (const change of changes) {
            switch (change.kind) {
              case "cell-content":
              case "cell-metadata":
                break;
              case "sheet-metadata": {
                const value = this.cloneMetadataHistoryValue(
                  useAfter ? change.after : change.before
                );
                this.workbookManager.setSheetMetadata(
                  {
                    workbookName: change.workbookName,
                    sheetName: change.sheetName,
                  },
                  value
                );
                break;
              }
              case "workbook-metadata": {
                const value = this.cloneMetadataHistoryValue(
                  useAfter ? change.after : change.before
                );
                this.workbookManager.setWorkbookMetadata(
                  change.workbookName,
                  value
                );
                break;
              }
            }
          }
        }

        for (const patch of step.patches) {
          const contentChanges = patch.flatMap((change) => {
            if (change.kind !== "cell-content") {
              return [];
            }
            const value = useAfter ? change.after : change.before;
            return [
              {
                address: change.address,
                before: useAfter ? change.before : change.after,
                after: value,
              },
            ];
          });
          if (contentChanges.length > 0) {
            this.evaluationManager.invalidateFromMutation({
              touchedCells: buildTouchedCells(contentChanges),
              resourceKeys: [],
            });
          }
        }
        return false;
      }
      case "named-expression-data":
        this.namedExpressionManager.applyHistoryChanges(
          step.changes,
          direction
        );
        this.invalidateNamedExpressionHistoryChanges(step.changes);
        return false;
      case "table-data": {
        this.tableManager.applyHistoryChanges(step.changes, direction);
        const tables = step.changes.flatMap((change) =>
          change.kind === "table"
            ? [change.before?.table, change.after?.table]
            : []
        );
        const resourceKeys = new Set<string>();
        for (const change of step.changes) {
          if (change.kind !== "table") {
            continue;
          }
          for (const state of [change.before, change.after]) {
            if (state) {
              resourceKeys.add(
                getTableResourceKey({
                  workbookName: state.workbookName,
                  tableName: state.tableName,
                })
              );
            }
          }
        }
        this.evaluationManager.invalidateFromMutation({
          touchedCells: buildTableTouchedCells(this.workbookManager, tables),
          tableContextChangedCells: buildTableContextChangedCells(
            this.workbookManager,
            tables
          ),
          resourceKeys: Array.from(resourceKeys),
        });
        return false;
      }
      case "style-data":
        this.styleManager.applyHistoryChanges(step.changes, direction);
        return false;
      case "range-metadata-data":
        this.rangeMetadataManager.applyHistoryChanges(step.changes, direction);
        return false;
      case "reference-data":
        this.referenceManager.applyHistoryChanges(step.changes, direction);
        return false;
      case "rename-sheet":
        this.workbookManager.renameSheet({
          workbookName: step.workbookName,
          sheetName: useAfter ? step.before : step.after,
          newSheetName: useAfter ? step.after : step.before,
        });
        return true;
      case "rename-workbook":
        this.workbookManager.renameWorkbook({
          workbookName: useAfter ? step.before : step.after,
          newWorkbookName: useAfter ? step.after : step.before,
        });
        return true;
      case "workbook-scope":
        this.restoreWorkbookScopeState(useAfter ? step.after : step.before);
        return true;
      case "sheet-scope":
        this.restoreSheetScopeState(useAfter ? step.after : step.before);
        return true;
    }
  }

  private invalidateNamedExpressionHistoryChanges(
    changes: readonly NamedExpressionMutation[]
  ): void {
    const resourceKeys = new Set<string>();
    for (const change of changes) {
      if (change.kind !== "named-expression") {
        continue;
      }
      for (const state of [change.before, change.after]) {
        if (!state) {
          continue;
        }
        const scope =
          state.scope.type === "global"
            ? {}
            : state.scope.type === "workbook"
            ? { workbookName: state.scope.workbookName }
            : {
                workbookName: state.scope.workbookName,
                sheetName: state.scope.sheetName,
              };
        resourceKeys.add(
          getNamedExpressionResourceKey({
            expressionName: state.expressionName,
            ...scope,
          })
        );
      }
    }
    if (resourceKeys.size > 0) {
      this.evaluationManager.invalidateFromMutation({
        touchedCells: [],
        resourceKeys: Array.from(resourceKeys),
      });
    }
  }

  private captureWorkbookScopeState(
    workbookName: string,
    detach = true
  ): WorkbookScopeState {
    const workbooks = this.workbookManager.getWorkbooks();
    const workbook = workbooks.get(workbookName);
    return {
      workbookName,
      workbookOrder: Array.from(workbooks.keys()),
      ...(workbook === undefined
        ? {}
        : { workbook: detach ? cloneHistoryValue(workbook) : workbook }),
    };
  }

  private captureSheetScopeState(
    workbookName: string,
    sheetName: string,
    detach = true
  ): SheetScopeState {
    const sheets = this.workbookManager.getSheets(workbookName);
    const sheet = sheets.get(sheetName);
    return {
      workbookName,
      sheetName,
      sheetOrder: Array.from(sheets.keys()),
      ...(sheet === undefined
        ? {}
        : { sheet: detach ? cloneHistoryValue(sheet) : sheet }),
    };
  }

  private restoreWorkbookScopeState(state: WorkbookScopeState): void {
    this.workbookManager.restoreWorkbookForHistory({
      workbookName: state.workbookName,
      workbookOrder: state.workbookOrder,
      workbook:
        state.workbook === undefined
          ? undefined
          : cloneHistoryValue(state.workbook),
    });
  }

  private restoreSheetScopeState(state: SheetScopeState): void {
    this.workbookManager.restoreSheetForHistory({
      workbookName: state.workbookName,
      sheetName: state.sheetName,
      sheetOrder: state.sheetOrder,
      sheet:
        state.sheet === undefined ? undefined : cloneHistoryValue(state.sheet),
    });
  }

  private withWorkbookScopeHistory<T>(
    workbookName: string,
    callback: () => T
  ): T {
    const beforeView = this.captureWorkbookScopeState(workbookName, false);
    const minimumStep = {
      kind: "workbook-scope" as const,
      before: beforeView,
      after: {
        workbookName,
        workbookOrder: beforeView.workbookOrder,
      },
    };
    if (
      !isHistoryValueSafelyRetainable(beforeView) ||
      this.pendingHistoryEstimatedBytes +
        estimateHistoryValueBytes(minimumStep) >
        this.undoRedoManager.maxBytes
    ) {
      this.rejectExplicitTransactionOrMarkOversized();
      return this.withoutWorkbookHistoryCapture(callback);
    }

    const before = cloneHistoryValue(beforeView);
    try {
      return this.withoutWorkbookHistoryCapture(callback);
    } finally {
      if (!this.pendingHistoryOversized) {
        const afterView = this.captureWorkbookScopeState(workbookName, false);
        if (!isHistoryValueSafelyRetainable(afterView)) {
          this.rejectExplicitTransactionOrMarkOversized({
            kind: "workbook-scope",
            before,
            after: afterView,
          });
        } else if (!historyValuesEqual(before, afterView)) {
          const candidate = {
            kind: "workbook-scope" as const,
            before,
            after: afterView,
          };
          if (
            this.pendingHistoryEstimatedBytes +
              estimateHistoryValueBytes(candidate) >
            this.undoRedoManager.maxBytes
          ) {
            this.rejectExplicitTransactionOrMarkOversized(candidate);
          } else {
            this.recordHistoryStep({
              ...candidate,
              after: cloneHistoryValue(afterView),
            });
          }
        }
      }
    }
  }

  private withSheetScopeHistory<T>(
    workbookName: string,
    sheetName: string,
    callback: () => T
  ): T {
    const beforeView = this.captureSheetScopeState(
      workbookName,
      sheetName,
      false
    );
    const minimumStep = {
      kind: "sheet-scope" as const,
      before: beforeView,
      after: {
        workbookName,
        sheetName,
        sheetOrder: beforeView.sheetOrder,
      },
    };
    if (
      !isHistoryValueSafelyRetainable(beforeView) ||
      this.pendingHistoryEstimatedBytes +
        estimateHistoryValueBytes(minimumStep) >
        this.undoRedoManager.maxBytes
    ) {
      this.rejectExplicitTransactionOrMarkOversized();
      return this.withoutWorkbookHistoryCapture(callback);
    }

    const before = cloneHistoryValue(beforeView);
    try {
      return this.withoutWorkbookHistoryCapture(callback);
    } finally {
      if (!this.pendingHistoryOversized) {
        const afterView = this.captureSheetScopeState(
          workbookName,
          sheetName,
          false
        );
        if (!isHistoryValueSafelyRetainable(afterView)) {
          this.rejectExplicitTransactionOrMarkOversized({
            kind: "sheet-scope",
            before,
            after: afterView,
          });
        } else if (!historyValuesEqual(before, afterView)) {
          const candidate = {
            kind: "sheet-scope" as const,
            before,
            after: afterView,
          };
          if (
            this.pendingHistoryEstimatedBytes +
              estimateHistoryValueBytes(candidate) >
            this.undoRedoManager.maxBytes
          ) {
            this.rejectExplicitTransactionOrMarkOversized(candidate);
          } else {
            this.recordHistoryStep({
              ...candidate,
              after: cloneHistoryValue(afterView),
            });
          }
        }
      }
    }
  }

  private materializeDefaultTableHeaders(
    table: TableDefinition
  ): MutationInvalidation["touchedCells"] {
    const changes: Array<{
      address: CellAddress;
      before: SerializedCellValue;
      after: SerializedCellValue;
    }> = [];

    for (const header of table.headers.values()) {
      const address = {
        workbookName: table.workbookName,
        sheetName: table.sheetName,
        rowIndex: table.start.rowIndex,
        colIndex: table.start.colIndex + header.index,
      };
      const before = this.workbookManager.getCellContent(address);
      if (before !== undefined && before !== "") {
        continue;
      }

      this.workbookManager.setCellContent(address, header.name);
      changes.push({ address, before, after: header.name });
    }

    return buildTouchedCells(changes);
  }

  private renameTableHeaderReferences(updates: TableHeaderUpdate[]): {
    changedFormulaCells: CellAddress[];
    resourceKeys: string[];
  } {
    const renamesByTable = new Map<TableDefinition, Map<string, string>>();
    for (const update of updates) {
      if (update.oldName === update.newName) {
        continue;
      }
      let renames = renamesByTable.get(update.table);
      if (!renames) {
        renames = new Map();
        renamesByTable.set(update.table, renames);
      }
      renames.set(update.oldName, update.newName);
    }

    const changedFormulaCells: CellAddress[] = [];
    const resourceKeys: string[] = [];
    for (const [table, columnRenames] of renamesByTable) {
      changedFormulaCells.push(
        ...this.workbookManager.updateAllFormulas((formula, formulaAddress) =>
          renameTableColumnsInFormula({
            formula,
            tableName: table.name,
            tableWorkbookName: table.workbookName,
            formulaWorkbookName: formulaAddress.workbookName,
            columnRenames,
            includeImplicitReferences:
              this.tableManager.isCellInTable(formulaAddress) === table,
          })
        )
      );
      resourceKeys.push(
        getTableResourceKey({
          workbookName: table.workbookName,
          tableName: table.name,
        }),
        ...this.namedExpressionManager.updateAllNamedExpressions(
          (formula, scope) =>
            renameTableColumnsInFormula({
              formula,
              tableName: table.name,
              tableWorkbookName: table.workbookName,
              formulaWorkbookName: scope.workbookName,
              columnRenames,
            })
        )
      );
    }

    return { changedFormulaCells, resourceKeys };
  }

  private getWorkbookResourceKeys(workbookName: string): string[] {
    const resourceKeys = new Set<string>([
      getWorkbookResourceKey(workbookName),
    ]);

    for (const sheetName of this.workbookManager
      .getWorkbooks()
      .get(workbookName)
      ?.sheets.keys() ?? []) {
      resourceKeys.add(getSheetResourceKey({ workbookName, sheetName }));
    }

    for (const tableName of this.tableManager.getTables(workbookName).keys()) {
      resourceKeys.add(getTableResourceKey({ workbookName, tableName }));
    }

    const namedExpressions = this.namedExpressionManager.getNamedExpressions();
    for (const name of namedExpressions.workbookExpressions
      .get(workbookName)
      ?.keys() ?? []) {
      resourceKeys.add(
        getNamedExpressionResourceKey({ expressionName: name, workbookName })
      );
    }
    for (const [
      sheetName,
      expressions,
    ] of namedExpressions.sheetExpressions.get(workbookName) ?? []) {
      resourceKeys.add(getSheetResourceKey({ workbookName, sheetName }));
      for (const name of expressions.keys()) {
        resourceKeys.add(
          getNamedExpressionResourceKey({
            expressionName: name,
            workbookName,
            sheetName,
          })
        );
      }
    }

    return Array.from(resourceKeys);
  }

  private getSheetResourceKeys(opts: {
    workbookName: string;
    sheetName: string;
  }): string[] {
    const resourceKeys = new Set<string>([
      getWorkbookResourceKey(opts.workbookName),
      getSheetResourceKey(opts),
    ]);

    for (const [tableName, table] of this.tableManager.getTables(
      opts.workbookName
    )) {
      if (table.sheetName === opts.sheetName) {
        resourceKeys.add(
          getTableResourceKey({ workbookName: opts.workbookName, tableName })
        );
      }
    }

    const sheetExpressions = this.namedExpressionManager
      .getNamedExpressions()
      .sheetExpressions.get(opts.workbookName)
      ?.get(opts.sheetName);
    for (const name of sheetExpressions?.keys() ?? []) {
      resourceKeys.add(
        getNamedExpressionResourceKey({
          expressionName: name,
          workbookName: opts.workbookName,
          sheetName: opts.sheetName,
        })
      );
    }

    return Array.from(resourceKeys);
  }

  private assertNamedExpressionScopeExists(opts: {
    workbookName?: string;
    sheetName?: string;
  }): void {
    if (opts.sheetName && !opts.workbookName) {
      throw new Error("Missing workbookName");
    }

    if (!opts.workbookName) {
      return;
    }

    if (!this.workbookManager.getWorkbooks().has(opts.workbookName)) {
      throw new Error(`Workbook not found: ${opts.workbookName}`);
    }

    if (
      opts.sheetName &&
      !this.workbookManager.getSheet({
        workbookName: opts.workbookName,
        sheetName: opts.sheetName,
      })
    ) {
      throw new Error(`Sheet not found: ${opts.sheetName}`);
    }
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
    metadata: MetadataType<TMetadata, "cell"> | undefined
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.setCellMetadata(address, metadata);
      this.emitUpdate();
    });
  }

  /**
   * Get metadata for a cell
   */
  getCellMetadata(
    address: CellAddress
  ): MetadataType<TMetadata, "cell"> | undefined {
    const metadata = this.workbookManager.getCellMetadata(address);
    return metadata as MetadataType<TMetadata, "cell"> | undefined;
  }

  /**
   * Get all cell metadata for a sheet (serialized as Map)
   */
  getSheetMetadataSerialized(opts: {
    sheetName: string;
    workbookName: string;
  }): Map<string, MetadataType<TMetadata, "sheet">> {
    return this.workbookManager.getSheetMetadataSerialized(opts) as Map<
      string,
      MetadataType<TMetadata, "sheet">
    >;
  }

  /**
   * Set metadata for a sheet
   * Sheet metadata can contain text boxes, frozen panes, print settings, or any consumer-defined data
   */
  setSheetMetadata(
    opts: { workbookName: string; sheetName: string },
    metadata: MetadataType<TMetadata, "sheet">
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.setSheetMetadata(opts, metadata);
      this.emitUpdate();
    });
  }

  /**
   * Get metadata for a sheet
   */
  getSheetMetadata(opts: {
    workbookName: string;
    sheetName: string;
  }): MetadataType<TMetadata, "sheet"> | undefined {
    return this.workbookManager.getSheetMetadata(opts) as
      | MetadataType<TMetadata, "sheet">
      | undefined;
  }

  /**
   * Set metadata for a workbook
   * Workbook metadata can contain themes, document properties, settings, or any consumer-defined data
   */
  setWorkbookMetadata(
    workbookName: string,
    metadata: MetadataType<TMetadata, "workbook">
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.setWorkbookMetadata(workbookName, metadata);
      this.emitUpdate();
    });
  }

  /**
   * Get metadata for a workbook
   */
  getWorkbookMetadata(
    workbookName: string
  ): MetadataType<TMetadata, "workbook"> | undefined {
    return this.workbookManager.getWorkbookMetadata(workbookName) as
      | MetadataType<TMetadata, "workbook">
      | undefined;
  }

  /**
   * Add metadata attached to one or more ranges.
   * Range metadata is arbitrary consumer-defined data that follows range
   * copy/paste, autofill, workbook/sheet rename, and serialization flows.
   */
  addRangeMetadata(
    metadata: RangeMetadataInput<MetadataType<TMetadata, "range">>
  ): string {
    return this.withUndoRedoCheckpoint(() => {
      const id = this.rangeMetadataManager.addRangeMetadata(metadata);
      this.emitUpdate();
      return id;
    });
  }

  /**
   * Remove a range metadata entry by id.
   */
  removeRangeMetadata(id: string): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.rangeMetadataManager.removeRangeMetadata(id);
      if (removed) {
        this.emitUpdate();
      }
    });
  }

  /**
   * Get all range metadata entries.
   */
  getAllRangeMetadata(): RangeMetadata<MetadataType<TMetadata, "range">>[] {
    return this.rangeMetadataManager.getAllRangeMetadata();
  }

  /**
   * Get range metadata entries that apply to a specific cell.
   */
  getRangeMetadataForCell(
    address: CellAddress
  ): RangeMetadata<MetadataType<TMetadata, "range">>[] {
    return this.rangeMetadataManager.getRangeMetadataForCell(address);
  }

  /**
   * Get range metadata entries intersecting with a range.
   */
  getRangeMetadataIntersectingWithRange(
    range: RangeAddress
  ): RangeMetadata<MetadataType<TMetadata, "range">>[] {
    return this.rangeMetadataManager.getRangeMetadataIntersectingWithRange(
      range
    );
  }

  /**
   * Clear range metadata from a range, preserving non-overlapping portions.
   */
  clearRangeMetadata(range: RangeAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      this.rangeMetadataManager.clearRangeMetadataInRange(range);
      this.emitUpdate();
    });
  }

  //#endregion

  //#region Reference Tracking
  /**
   * Create a tracked reference to a range
   * Returns a stable UUID that can be used to retrieve the address later
   * The reference automatically updates when sheets/workbooks are renamed
   */
  createRef(address: RangeAddress): string {
    return this.withUndoRedoCheckpoint(() => {
      const id = this.referenceManager.createRef(address);
      this.emitUpdate();
      return id;
    });
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
    return this.withUndoRedoCheckpoint(() => {
      const deleted = this.referenceManager.deleteRef(refId);
      if (deleted) {
        this.emitUpdate();
      }
      return deleted;
    });
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
  addNamedExpression(opts: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }) {
    return this.withUndoRedoCheckpoint(() => {
      this.assertNamedExpressionScopeExists(opts);
      this.namedExpressionManager.addNamedExpression(opts);
      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getNamedExpressionResourceKey(opts)],
      });
    });
  }

  removeNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.namedExpressionManager.removeNamedExpression(opts);
      if (removed) {
        this.emitMutation({
          touchedCells: [],
          resourceKeys: [getNamedExpressionResourceKey(opts)],
        });
      }
    });
  }

  /**
   * Check if a named expression exists
   */
  hasNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): boolean {
    const scope =
      opts.sheetName && opts.workbookName
        ? {
            type: "sheet" as const,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }
        : opts.workbookName
        ? { type: "workbook" as const, workbookName: opts.workbookName }
        : { type: "global" as const };

    return !!this.namedExpressionManager.getNamedExpression({
      name: opts.expressionName,
      scope,
    });
  }

  updateNamedExpression(opts: {
    expression: string;
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      this.namedExpressionManager.updateNamedExpression(opts);
      this.emitMutation({
        touchedCells: [],
        resourceKeys: [getNamedExpressionResourceKey(opts)],
      });
    });
  }

  renameNamedExpression(opts: {
    expressionName: string;
    sheetName?: string;
    workbookName?: string;
    newName: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      this.namedExpressionManager.renameNamedExpression(opts);

      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameNamedExpressionInFormula(
          formula,
          opts.expressionName,
          opts.newName
        )
      );

      const changedNamedExpressions =
        this.namedExpressionManager.updateAllNamedExpressions((formula) =>
          renameNamedExpressionInFormula(
            formula,
            opts.expressionName,
            opts.newName
          )
        );

      this.emitMutation({
        touchedCells: buildFormulaTouchedCells(changedCells),
        resourceKeys: [
          getNamedExpressionResourceKey({
            expressionName: opts.expressionName,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }),
          getNamedExpressionResourceKey({
            expressionName: opts.newName,
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          }),
          ...changedNamedExpressions,
        ],
      });
    });
  }

  setNamedExpressions(
    opts: (
      | { type: "global" }
      | { type: "sheet"; sheetName: string; workbookName: string }
      | { type: "workbook"; workbookName: string }
    ) & {
      expressions: Map<string, NamedExpression>;
    }
  ) {
    return this.withUndoRedoCheckpoint(() => {
      const allExpressions = this.namedExpressionManager.getNamedExpressions();
      let previousExpressions: Map<string, NamedExpression> | undefined;

      if (opts.type === "global") {
        previousExpressions = new Map(allExpressions.globalExpressions);
      } else if (opts.type === "workbook") {
        previousExpressions = new Map(
          allExpressions.workbookExpressions.get(opts.workbookName) || []
        );
      } else {
        const sheetExpressions = allExpressions.sheetExpressions
          .get(opts.workbookName)
          ?.get(opts.sheetName);
        previousExpressions = new Map(sheetExpressions || []);
      }

      this.namedExpressionManager.setNamedExpressions(opts);

      const scope =
        opts.type === "global"
          ? {}
          : opts.type === "workbook"
          ? { workbookName: opts.workbookName }
          : {
              workbookName: opts.workbookName,
              sheetName: opts.sheetName,
            };

      this.emitMutation({
        touchedCells: [],
        resourceKeys: [
          ...getNamedExpressionScopeResourceKeys(
            previousExpressions.keys(),
            scope
          ),
          ...getNamedExpressionScopeResourceKeys(
            opts.expressions.keys(),
            scope
          ),
        ],
      });
    });
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
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const table = this.tableManager.addTable({
        ...props,
        getCellValue: (cellAddress: CellAddress) =>
          this.getCellValue(cellAddress),
      });
      const generatedHeaderCells = this.materializeDefaultTableHeaders(table);

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, [table]),
          generatedHeaderCells
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [table]
        ),
        resourceKeys: [
          getTableResourceKey({
            workbookName: props.workbookName,
            tableName: props.tableName,
          }),
        ],
      });
    });
  }

  renameTable(
    workbookName: string,
    names: { oldName: string; newName: string }
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTable = this.tableManager.getTable({
        workbookName,
        name: names.oldName,
      });
      const oldTableSnapshot = oldTable
        ? { ...oldTable, headers: new Map(oldTable.headers) }
        : undefined;

      this.tableManager.renameTable(workbookName, names);

      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameTableInFormula(formula, names.oldName, names.newName)
      );

      const changedNamedExpressions =
        this.namedExpressionManager.updateAllNamedExpressions((formula) =>
          renameTableInFormula(formula, names.oldName, names.newName)
        );

      const newTable = this.tableManager.getTable({
        workbookName,
        name: names.newName,
      });

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, [oldTableSnapshot]),
          buildTableTouchedCells(this.workbookManager, [newTable]),
          buildFormulaTouchedCells(changedCells)
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [oldTableSnapshot, newTable]
        ),
        resourceKeys: [
          getTableResourceKey({
            workbookName,
            tableName: names.oldName,
          }),
          getTableResourceKey({
            workbookName,
            tableName: names.newName,
          }),
          ...changedNamedExpressions,
        ],
      });
    });
  }

  updateTable(opts: {
    tableName: string;
    sheetName?: string;
    start?: string;
    numRows?: SpreadsheetRangeEnd;
    numCols?: number;
    workbookName: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTable = this.tableManager.getTable({
        workbookName: opts.workbookName,
        name: opts.tableName,
      });
      const oldTableSnapshot = oldTable
        ? { ...oldTable, headers: new Map(oldTable.headers) }
        : undefined;

      this.tableManager.updateTable({
        ...opts,
        getCellValue: (cellAddress: CellAddress) =>
          this.getCellValue(cellAddress),
      });

      const newTable = this.tableManager.getTable({
        workbookName: opts.workbookName,
        name: opts.tableName,
      });
      const generatedHeaderCells = newTable
        ? this.materializeDefaultTableHeaders(newTable)
        : [];

      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, [oldTableSnapshot]),
          buildTableTouchedCells(this.workbookManager, [newTable]),
          generatedHeaderCells
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [oldTableSnapshot, newTable]
        ),
        resourceKeys: [
          getTableResourceKey({
            workbookName: opts.workbookName,
            tableName: opts.tableName,
          }),
        ],
      });
    });
  }

  removeTable(opts: { tableName: string; workbookName: string }): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTable = this.tableManager.getTable({
        workbookName: opts.workbookName,
        name: opts.tableName,
      });
      const oldTableSnapshot = oldTable
        ? { ...oldTable, headers: new Map(oldTable.headers) }
        : undefined;

      const found = this.tableManager.removeTable(opts);
      if (found) {
        this.emitMutation({
          touchedCells: buildTableTouchedCells(this.workbookManager, [
            oldTableSnapshot,
          ]),
          tableContextChangedCells: buildTableContextChangedCells(
            this.workbookManager,
            [oldTableSnapshot]
          ),
          resourceKeys: [
            getTableResourceKey({
              workbookName: opts.workbookName,
              tableName: opts.tableName,
            }),
          ],
        });
      }
    });
  }

  private getAllTables(): TableDefinition[] {
    return Array.from(this.tableManager.tables.values()).flatMap((tables) =>
      Array.from(tables.values()).map((table) => ({
        ...table,
        headers: new Map(table.headers),
      }))
    );
  }

  /**
   * Check if a table exists
   */
  hasTable(opts: { tableName: string; workbookName: string }): boolean {
    return !!this.tableManager.getTable({
      workbookName: opts.workbookName,
      name: opts.tableName,
    });
  }

  /**
   * Get a table definition by name
   */
  getTable(opts: {
    tableName: string;
    workbookName: string;
  }): TableDefinition | undefined {
    return this.tableManager.getTable({
      workbookName: opts.workbookName,
      name: opts.tableName,
    });
  }

  resetTables(tables: Map<string, Map<string, TableDefinition>>): void {
    return this.withUndoRedoCheckpoint(() => {
      const oldTables = this.getAllTables();
      const newTables = Array.from(tables.values()).flatMap((workbookTables) =>
        Array.from(workbookTables.values())
      );
      const resourceKeys = new Set<string>();
      for (const table of [...oldTables, ...newTables]) {
        resourceKeys.add(
          getTableResourceKey({
            workbookName: table.workbookName,
            tableName: table.name,
          })
        );
      }

      this.tableManager.resetTables(tables);
      this.emitMutation({
        touchedCells: mergeTouchedCells(
          buildTableTouchedCells(this.workbookManager, oldTables),
          buildTableTouchedCells(this.workbookManager, newTables)
        ),
        tableContextChangedCells: buildTableContextChangedCells(
          this.workbookManager,
          [...oldTables, ...newTables]
        ),
        resourceKeys: Array.from(resourceKeys),
      });
    });
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
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.addConditionalStyle(style);
      this.emitUpdate();
    });
  }

  /**
   * Remove a conditional style rule by index
   */
  removeConditionalStyle(workbookName: string, index: number): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.styleManager.removeConditionalStyle(
        workbookName,
        index
      );
      if (removed) {
        this.emitUpdate();
      }
    });
  }

  /**
   * Get the count of conditional styles for a workbook
   */
  getConditionalStyleCount(workbookName: string): number {
    const allStyles = this.styleManager.getAllConditionalStyles();
    return allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === workbookName)
    ).length;
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

  getCellDataType(cellAddress: CellAddress): CellDataType {
    return this.styleManager.getCellDataType(cellAddress);
  }

  getDataTypeForRange(range: RangeAddress): CellDataType | undefined {
    return this.styleManager.getDataTypeForRange(range);
  }

  getAllCellDataTypes(): DirectCellDataType[] {
    return this.styleManager.getAllCellDataTypes();
  }

  addCellDataType(dataType: DirectCellDataType): void {
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.addCellDataType(dataType);
      this.emitUpdate();
    });
  }

  clearCellDataTypes(range: RangeAddress): void {
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.clearCellDataTypesInRange(range);
      this.emitUpdate();
    });
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
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.addCellStyle(style);
      this.emitUpdate();
    });
  }

  /**
   * Remove a direct cell style rule by index
   */
  removeCellStyle(workbookName: string, index: number): void {
    return this.withUndoRedoCheckpoint(() => {
      const removed = this.styleManager.removeCellStyle(workbookName, index);
      if (removed) {
        this.emitUpdate();
      }
    });
  }

  /**
   * Get the count of direct cell styles for a workbook
   */
  getCellStyleCount(workbookName: string): number {
    const allStyles = this.styleManager.getAllCellStyles();
    return allStyles.filter((s) =>
      s.areas.some((a) => a.workbookName === workbookName)
    ).length;
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
    return this.withUndoRedoCheckpoint(() => {
      this.styleManager.clearCellStyles(range);
      this.emitUpdate();
    });
  }

  //#endregion

  //#region Copy/Paste
  private batchCopyAncillaryMutations<TResult>(
    callback: () => TResult
  ): TResult {
    return this.styleManager.batchMutations(() =>
      this.rangeMetadataManager.batchMutations(callback)
    );
  }

  /**
   * Paste cells from source to target
   */
  pasteCells(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      const movedTables =
        options.cut === true
          ? this.tableManager.getTablesContainedInCells(source)
          : [];
      this.pasteCellsWithTables(source, target, options, movedTables);
    });
  }

  private pasteCellsWithTables(
    source: CellAddress[],
    target: CellAddress,
    options: CopyCellsOptions,
    movedTables: TableDefinition[]
  ): void {
    if (source.length === 0) {
      return;
    }

    const sourceBounds = this.getBoundsFromCells(source);
    const rowOffset = target.rowIndex - sourceBounds.minRow;
    const colOffset = target.colIndex - sourceBounds.minCol;

    this.batchCopyAncillaryMutations(() =>
      this.copyManager.pasteCells(source, target, options)
    );

    const relocatedTables = movedTables.map((table) =>
      this.tableManager.moveTable({
        workbookName: table.workbookName,
        tableName: table.name,
        target: {
          workbookName: target.workbookName,
          sheetName: target.sheetName,
          rowIndex: table.start.rowIndex + rowOffset,
          colIndex: table.start.colIndex + colOffset,
        },
      })
    );
    const tableResourceKeys = new Set<string>();
    for (const table of [...movedTables, ...relocatedTables]) {
      tableResourceKeys.add(
        getTableResourceKey({
          workbookName: table.workbookName,
          tableName: table.name,
        })
      );
    }

    this.emitMutation({
      touchedCells: mergeTouchedCells(
        buildTableTouchedCells(this.workbookManager, movedTables),
        buildTableTouchedCells(this.workbookManager, relocatedTables)
      ),
      tableContextChangedCells: buildTableContextChangedCells(
        this.workbookManager,
        [...movedTables, ...relocatedTables]
      ),
      resourceKeys: Array.from(tableResourceKeys),
    });
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
    return this.withUndoRedoCheckpoint(() => {
      this.batchCopyAncillaryMutations(() =>
        this.copyManager.fillAreas(seedRange, targetRanges, options)
      );

      this.emitMutation({
        touchedCells: [],
        resourceKeys: [],
      });
    });
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
    return this.withUndoRedoCheckpoint(() => {
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
        const shouldFill =
          pasteWidth > sourceWidth || pasteHeight > sourceHeight;

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
    });
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
    return this.withUndoRedoCheckpoint(() => {
      this.pasteCells([source], target, {
        cut: true,
        type: "formula",
        include: "all",
      });
    });
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
    return this.withUndoRedoCheckpoint(() => {
      const cells = this.copyManager.expandRangeToCells(sourceRange);
      const movedTables =
        this.tableManager.getTablesContainedInRange(sourceRange);
      this.pasteCellsWithTables(
        cells,
        target,
        {
          cut: true,
          type: "formula",
          include: "all",
        },
        movedTables
      );
    });
  }
  //#endregion

  //#region Sheets
  addSheet(opts: { workbookName: string; sheetName: string }): Sheet {
    return this.withUndoRedoCheckpoint(() => {
      return this.withSheetScopeHistory(
        opts.workbookName,
        opts.sheetName,
        () => {
          const sheet = this.workbookManager.addSheet(opts);
          this.namedExpressionManager.addSheet(opts);
          this.emitMutation({
            touchedCells: [],
            resourceKeys: [getSheetResourceKey(opts)],
          });
          return sheet;
        }
      );
    });
  }

  createSheet(opts: {
    workbookName: string;
    sheetName?: string;
    baseName?: string;
  }): Sheet {
    return this.withUndoRedoCheckpoint(() => {
      const sheetName =
        opts.sheetName ??
        this.workbookManager.getNextAvailableSheetName(
          opts.workbookName,
          opts.baseName
        );

      return this.addSheet({
        workbookName: opts.workbookName,
        sheetName,
      });
    });
  }

  /**
   * Clone a sheet and all sheet-scoped engine state into a new sheet in the
   * same workbook. Tables receive unique workbook-scoped names.
   */
  cloneSheet(opts: {
    workbookName: string;
    sheetName: string;
    newSheetName: string;
  }): Sheet {
    return this.withUndoRedoCheckpoint(() => {
      return this.withSheetScopeHistory(
        opts.workbookName,
        opts.newSheetName,
        () => {
          const sourceSheet = this.workbookManager.getSheet({
            workbookName: opts.workbookName,
            sheetName: opts.sheetName,
          });
          if (!sourceSheet) {
            throw new Error(`Source sheet "${opts.sheetName}" not found`);
          }
          if (
            this.workbookManager.getSheet({
              workbookName: opts.workbookName,
              sheetName: opts.newSheetName,
            })
          ) {
            throw new Error(
              `Target sheet "${opts.newSheetName}" already exists`
            );
          }

          const sourceTables = Array.from(
            this.tableManager.getTables(opts.workbookName).values()
          ).filter((table) => table.sheetName === opts.sheetName);
          const usedTableNames = new Set(
            this.tableManager.getTables(opts.workbookName).keys()
          );
          const tableNameMap = new Map<string, string>();
          for (const table of sourceTables) {
            let suffix = 2;
            let clonedName = `${table.name}_${suffix}`;
            while (usedTableNames.has(clonedName)) {
              suffix++;
              clonedName = `${table.name}_${suffix}`;
            }
            usedTableNames.add(clonedName);
            tableNameMap.set(table.name, clonedName);
          }

          const rewriteFormula = (formula: string): string => {
            let rewritten = renameSheetInFormula({
              formula,
              oldSheetName: opts.sheetName,
              newSheetName: opts.newSheetName,
              workbookName: opts.workbookName,
            });
            for (const [oldTableName, newTableName] of tableNameMap) {
              rewritten = renameTableInFormula(
                rewritten,
                oldTableName,
                newTableName,
                opts.workbookName
              );
            }
            return rewritten;
          };

          const clonedContent = new Map<string, SerializedCellValue>();
          for (const [cellReference, content] of sourceSheet.content) {
            clonedContent.set(
              cellReference,
              typeof content === "string" && content.startsWith("=")
                ? `=${rewriteFormula(content.slice(1))}`
                : content
            );
          }

          const targetSheet = this.workbookManager.addSheet({
            workbookName: opts.workbookName,
            sheetName: opts.newSheetName,
          });
          this.namedExpressionManager.addSheet({
            workbookName: opts.workbookName,
            sheetName: opts.newSheetName,
          });
          this.workbookManager.setSheetContent(
            {
              workbookName: opts.workbookName,
              sheetName: opts.newSheetName,
            },
            clonedContent
          );
          targetSheet.metadata = cloneHistoryValue(sourceSheet.metadata);
          if (sourceSheet.sheetMetadata !== undefined) {
            targetSheet.sheetMetadata = cloneHistoryValue(
              sourceSheet.sheetMetadata
            );
          }

          const sourceSheetExpressions =
            this.namedExpressionManager
              .getNamedExpressions()
              .sheetExpressions.get(opts.workbookName)
              ?.get(opts.sheetName);
          if (sourceSheetExpressions) {
            for (const [expressionName, expression] of sourceSheetExpressions) {
              this.namedExpressionManager.addNamedExpression({
                expressionName,
                expression: rewriteFormula(expression.expression),
                workbookName: opts.workbookName,
                sheetName: opts.newSheetName,
              });
            }
          }

          const clonedTables: TableDefinition[] = [];
          for (const table of sourceTables) {
            const tableName = tableNameMap.get(table.name)!;
            this.tableManager.copyTable(
              { workbookName: opts.workbookName, tableName: table.name },
              {
                workbookName: opts.workbookName,
                tableName,
                sheetName: opts.newSheetName,
              }
            );
            const clonedTable = this.tableManager.getTable({
              workbookName: opts.workbookName,
              name: tableName,
            });
            if (clonedTable) {
              clonedTables.push(clonedTable);
            }
          }

          const cloneAreas = (areas: RangeAddress[]): RangeAddress[] =>
            areas
              .filter(
                (area) =>
                  area.workbookName === opts.workbookName &&
                  area.sheetName === opts.sheetName
              )
              .map((area) => ({ ...area, sheetName: opts.newSheetName }));

          for (const style of this.styleManager.getAllConditionalStyles()) {
            const areas = cloneAreas(style.areas);
            if (areas.length === 0) {
              continue;
            }
            const condition =
              style.condition.type === "formula"
                ? {
                    ...style.condition,
                    formula: rewriteFormula(style.condition.formula),
                  }
                : {
                    ...style.condition,
                    min:
                      style.condition.min.type === "number"
                        ? {
                            ...style.condition.min,
                            valueFormula: rewriteFormula(
                              style.condition.min.valueFormula
                            ),
                          }
                        : style.condition.min,
                    max:
                      style.condition.max.type === "number"
                        ? {
                            ...style.condition.max,
                            valueFormula: rewriteFormula(
                              style.condition.max.valueFormula
                            ),
                          }
                        : style.condition.max,
                  };
            this.styleManager.addConditionalStyle({
              areas,
              condition,
            });
          }

          for (const style of this.styleManager.getAllCellStyles()) {
            const areas = cloneAreas(style.areas);
            if (areas.length > 0) {
              this.styleManager.addCellStyle({ ...style, areas });
            }
          }

          for (const dataType of this.styleManager.getAllCellDataTypes()) {
            const areas = cloneAreas(dataType.areas);
            if (areas.length > 0) {
              this.styleManager.addCellDataType({ ...dataType, areas });
            }
          }

          for (const metadata of this.rangeMetadataManager.getAllRangeMetadata()) {
            const areas = cloneAreas(metadata.areas);
            if (areas.length > 0) {
              this.rangeMetadataManager.addRangeMetadata({
                areas,
                metadata: cloneHistoryValue(metadata.metadata),
              });
            }
          }

          const resourceKeys = [
            getSheetResourceKey({
              workbookName: opts.workbookName,
              sheetName: opts.newSheetName,
            }),
            ...Array.from(
              sourceSheetExpressions?.keys() ?? [],
              (expressionName) =>
                getNamedExpressionResourceKey({
                  workbookName: opts.workbookName,
                  sheetName: opts.newSheetName,
                  expressionName,
                })
            ),
            ...clonedTables.map((table) =>
              getTableResourceKey({
                workbookName: opts.workbookName,
                tableName: table.name,
              })
            ),
          ];
          this.emitMutation({
            touchedCells: buildTableTouchedCells(
              this.workbookManager,
              clonedTables
            ),
            tableContextChangedCells: buildTableContextChangedCells(
              this.workbookManager,
              clonedTables
            ),
            resourceKeys,
          });

          return targetSheet;
        }
      );
    });
  }

  removeSheet(opts: { workbookName: string; sheetName: string }): void {
    return this.withUndoRedoCheckpoint(() => {
      this.withSheetScopeHistory(opts.workbookName, opts.sheetName, () => {
        const resourceKeys = this.getSheetResourceKeys(opts);
        this.workbookManager.removeSheet(opts);
        this.namedExpressionManager.removeSheet(opts);
        this.tableManager.removeSheet(opts);
        this.styleManager.removeSheetStyles(opts.workbookName, opts.sheetName);
        this.rangeMetadataManager.removeSheetRangeMetadata(
          opts.workbookName,
          opts.sheetName
        );
        this.referenceManager.invalidateSheet(
          opts.workbookName,
          opts.sheetName
        );
        this.emitMutation({
          touchedCells: [],
          resourceKeys,
          removedScopes: [{ type: "sheet", ...opts }],
        });
        this.pendingResourceEvents.push({
          type: "sheet-delete",
          ...opts,
        });
      });
    });
  }

  renameSheet(opts: {
    sheetName: string;
    newSheetName: string;
    workbookName: string;
  }): void {
    return this.withUndoRedoCheckpoint(() => {
      this.flushPendingWorkbookChanges();
      const oldResourceKeys = this.getSheetResourceKeys(opts);

      this.workbookManager.renameSheet(opts);
      this.recordHistoryStepBeforePendingWorkbookChanges({
        kind: "rename-sheet",
        workbookName: opts.workbookName,
        before: opts.sheetName,
        after: opts.newSheetName,
      });
      this.namedExpressionManager.renameSheet(opts);
      this.tableManager.updateTablesForSheetRename(opts);
      this.styleManager.updateSheetName(
        opts.workbookName,
        opts.sheetName,
        opts.newSheetName
      );
      this.rangeMetadataManager.updateSheetName(
        opts.workbookName,
        opts.sheetName,
        opts.newSheetName
      );
      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameSheetInFormula({
          formula,
          oldSheetName: opts.sheetName,
          newSheetName: opts.newSheetName,
        })
      );
      this.referenceManager.updateSheetName(
        opts.workbookName,
        opts.sheetName,
        opts.newSheetName
      );

      this.emitMutation({
        touchedCells: buildFormulaTouchedCells(changedCells),
        resourceKeys: Array.from(
          new Set([
            ...oldResourceKeys,
            ...this.getSheetResourceKeys({
              workbookName: opts.workbookName,
              sheetName: opts.newSheetName,
            }),
          ])
        ),
      });
      this.pendingResourceEvents.push({
        type: "sheet-rename",
        workbookName: opts.workbookName,
        sheetName: opts.sheetName,
        newSheetName: opts.newSheetName,
      });
    });
  }

  /**
   * Check if a sheet exists
   */
  hasSheet(opts: { workbookName: string; sheetName: string }): boolean {
    return !!this.workbookManager.getSheet(opts);
  }

  getSheets(workbookName: string) {
    return this.workbookManager.getSheets(workbookName);
  }

  getOrderedSheets(workbookName: string) {
    return this.workbookManager.getOrderedSheets(workbookName);
  }

  getOrderedSheetNames(workbookName: string) {
    return this.workbookManager.getOrderedSheetNames(workbookName);
  }

  getNextAvailableSheetName(workbookName: string, baseName?: string) {
    return this.workbookManager.getNextAvailableSheetName(
      workbookName,
      baseName
    );
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

  /**
   * Search raw stored string content without evaluating cell values.
   */
  search(query: string, options?: SearchOptions): SearchMatch[] {
    return this.workbookManager.search(query, options);
  }

  /**
   * Replace one specific search occurrence inside one addressed cell.
   */
  replace(
    query: string,
    replacement: string,
    target: ReplaceTarget,
    options?: { caseSensitive?: boolean }
  ): ReplaceChange {
    return this.withUndoRedoCheckpoint(() => {
      const prepared = this.workbookManager.prepareReplace(
        query,
        replacement,
        target,
        options
      );

      this.workbookManager.setCellContent(
        prepared.address,
        prepared.afterContent
      );
      this.emitMutation({
        touchedCells: buildTouchedCells([
          {
            address: prepared.address,
            before: prepared.beforeContent,
            after: prepared.afterContent,
          },
        ]),
        resourceKeys: [],
      });

      return prepared.change;
    });
  }

  /**
   * Replace all matching raw string occurrences within the requested scope.
   */
  replaceAll(
    query: string,
    replacement: string,
    options?: SearchOptions
  ): ReplaceChange[] {
    return this.withUndoRedoCheckpoint(() => {
      const preparedReplacements = this.workbookManager.prepareReplaceAll(
        query,
        replacement,
        options
      );

      if (preparedReplacements.length === 0) {
        return [];
      }

      for (const prepared of preparedReplacements) {
        this.workbookManager.setCellContent(
          prepared.address,
          prepared.afterContent
        );
      }

      this.emitMutation({
        touchedCells: buildTouchedCells(
          preparedReplacements.map((prepared) => ({
            address: prepared.address,
            before: prepared.beforeContent,
            after: prepared.afterContent,
          }))
        ),
        resourceKeys: [],
      });

      return preparedReplacements.flatMap((prepared) => prepared.changes);
    });
  }

  //#endregion

  //#region Workbook
  addWorkbook(workbookName: string): void {
    return this.withUndoRedoCheckpoint(() => {
      this.withWorkbookScopeHistory(workbookName, () => {
        this.workbookManager.addWorkbook(workbookName);
        this.namedExpressionManager.addWorkbook(workbookName);
        this.tableManager.addWorkbook(workbookName);
        this.emitMutation({
          touchedCells: [],
          resourceKeys: [getWorkbookResourceKey(workbookName)],
        });
      });
    });
  }

  removeWorkbook(workbookName: string): void {
    return this.withUndoRedoCheckpoint(() => {
      this.withWorkbookScopeHistory(workbookName, () => {
        const resourceKeys = this.getWorkbookResourceKeys(workbookName);
        const sheetNames = Array.from(
          this.workbookManager.getSheets(workbookName).keys()
        );
        this.workbookManager.removeWorkbook(workbookName);
        this.namedExpressionManager.removeWorkbook(workbookName);
        this.tableManager.removeWorkbook(workbookName);
        this.styleManager.removeWorkbookStyles(workbookName);
        this.rangeMetadataManager.removeWorkbookRangeMetadata(workbookName);
        this.referenceManager.invalidateWorkbook(workbookName);
        this.emitMutation({
          touchedCells: [],
          resourceKeys,
          removedScopes: [{ type: "workbook", workbookName }],
        });
        this.pendingResourceEvents.push(
          ...sheetNames.map((sheetName) => ({
            type: "sheet-delete" as const,
            workbookName,
            sheetName,
          })),
          { type: "workbook-delete", workbookName }
        );
      });
    });
  }

  /**
   * Check if a workbook exists
   */
  hasWorkbook(workbookName: string): boolean {
    return this.workbookManager.getWorkbooks().has(workbookName);
  }

  cloneWorkbook(fromWorkbookName: string, toWorkbookName: string): void {
    return this.withUndoRedoCheckpoint(() => {
      this.withWorkbookScopeHistory(toWorkbookName, () => {
        const sourceWorkbook = this.workbookManager
          .getWorkbooks()
          .get(fromWorkbookName);
        if (!sourceWorkbook) {
          throw new Error(`Source workbook "${fromWorkbookName}" not found`);
        }
        if (this.workbookManager.getWorkbooks().has(toWorkbookName)) {
          throw new Error(`Target workbook "${toWorkbookName}" already exists`);
        }

        this.workbookManager.addWorkbook(toWorkbookName);
        this.namedExpressionManager.addWorkbook(toWorkbookName);
        this.tableManager.addWorkbook(toWorkbookName);

        for (const [sheetName, sheet] of sourceWorkbook.sheets) {
          this.workbookManager.addSheet({
            workbookName: toWorkbookName,
            sheetName,
          });
          this.namedExpressionManager.addSheet({
            workbookName: toWorkbookName,
            sheetName,
          });
          this.workbookManager.setSheetContent(
            { workbookName: toWorkbookName, sheetName },
            new Map(sheet.content)
          );

          const targetSheet = this.workbookManager.getSheet({
            workbookName: toWorkbookName,
            sheetName,
          });
          if (targetSheet) {
            targetSheet.metadata = new Map(sheet.metadata);
            if (sheet.sheetMetadata !== undefined) {
              targetSheet.sheetMetadata = structuredClone(sheet.sheetMetadata);
            }
          }
        }

        const targetWorkbook = this.workbookManager
          .getWorkbooks()
          .get(toWorkbookName);
        if (targetWorkbook && sourceWorkbook.workbookMetadata !== undefined) {
          targetWorkbook.workbookMetadata = structuredClone(
            sourceWorkbook.workbookMetadata
          );
        }

        const namedExpressions =
          this.namedExpressionManager.getNamedExpressions();
        const sourceWorkbookExpressions =
          namedExpressions.workbookExpressions.get(fromWorkbookName);
        if (sourceWorkbookExpressions) {
          for (const [
            expressionName,
            expression,
          ] of sourceWorkbookExpressions) {
            this.namedExpressionManager.addNamedExpression({
              expressionName,
              expression: expression.expression,
              workbookName: toWorkbookName,
            });
          }
        }

        const sourceSheetExpressions =
          namedExpressions.sheetExpressions.get(fromWorkbookName);
        if (sourceSheetExpressions) {
          for (const [sheetName, expressions] of sourceSheetExpressions) {
            for (const [expressionName, expression] of expressions) {
              this.namedExpressionManager.addNamedExpression({
                expressionName,
                expression: expression.expression,
                workbookName: toWorkbookName,
                sheetName,
              });
            }
          }
        }

        const sourceTables = this.tableManager.tables.get(fromWorkbookName);
        if (sourceTables) {
          for (const [tableName] of sourceTables) {
            this.tableManager.copyTable(
              { workbookName: fromWorkbookName, tableName },
              { workbookName: toWorkbookName, tableName }
            );
          }
        }

        for (const style of this.styleManager.getAllConditionalStyles()) {
          if (
            style.areas.some((area) => area.workbookName === fromWorkbookName)
          ) {
            this.styleManager.addConditionalStyle({
              ...style,
              areas: style.areas.map((area) =>
                area.workbookName === fromWorkbookName
                  ? { ...area, workbookName: toWorkbookName }
                  : area
              ),
            });
          }
        }

        for (const style of this.styleManager.getAllCellStyles()) {
          if (
            style.areas.some((area) => area.workbookName === fromWorkbookName)
          ) {
            this.styleManager.addCellStyle({
              ...style,
              areas: style.areas.map((area) =>
                area.workbookName === fromWorkbookName
                  ? { ...area, workbookName: toWorkbookName }
                  : area
              ),
            });
          }
        }

        for (const dataType of this.styleManager.getAllCellDataTypes()) {
          const clonedAreas = dataType.areas
            .filter((area) => area.workbookName === fromWorkbookName)
            .map((area) => ({ ...area, workbookName: toWorkbookName }));
          if (clonedAreas.length > 0) {
            this.styleManager.addCellDataType({
              ...dataType,
              areas: clonedAreas,
            });
          }
        }

        for (const metadata of this.rangeMetadataManager.getAllRangeMetadata()) {
          if (
            metadata.areas.some(
              (area) => area.workbookName === fromWorkbookName
            )
          ) {
            this.rangeMetadataManager.addRangeMetadata({
              metadata: metadata.metadata,
              areas: metadata.areas.map((area) =>
                area.workbookName === fromWorkbookName
                  ? { ...area, workbookName: toWorkbookName }
                  : area
              ),
            });
          }
        }

        this.workbookManager.updateFormulasForWorkbook(
          toWorkbookName,
          (formula) =>
            renameWorkbookInFormula({
              formula,
              oldWorkbookName: fromWorkbookName,
              newWorkbookName: toWorkbookName,
            })
        );

        this.emitMutation({
          touchedCells: [],
          resourceKeys: [getWorkbookResourceKey(toWorkbookName)],
        });
      });
    });
  }

  renameWorkbook(opts: { workbookName: string; newWorkbookName: string }) {
    return this.withUndoRedoCheckpoint(() => {
      this.flushPendingWorkbookChanges();
      const oldResourceKeys = this.getWorkbookResourceKeys(opts.workbookName);

      this.workbookManager.renameWorkbook(opts);
      this.recordHistoryStepBeforePendingWorkbookChanges({
        kind: "rename-workbook",
        before: opts.workbookName,
        after: opts.newWorkbookName,
      });
      this.namedExpressionManager.renameWorkbook(opts);
      this.tableManager.updateTablesForWorkbookRename(opts);
      this.styleManager.updateWorkbookName(
        opts.workbookName,
        opts.newWorkbookName
      );
      this.rangeMetadataManager.updateWorkbookName(
        opts.workbookName,
        opts.newWorkbookName
      );
      const changedCells = this.workbookManager.updateAllFormulas((formula) =>
        renameWorkbookInFormula({
          formula,
          oldWorkbookName: opts.workbookName,
          newWorkbookName: opts.newWorkbookName,
        })
      );
      this.referenceManager.updateWorkbookName(
        opts.workbookName,
        opts.newWorkbookName
      );

      this.emitMutation({
        touchedCells: buildFormulaTouchedCells(changedCells),
        resourceKeys: Array.from(
          new Set([
            ...oldResourceKeys,
            ...this.getWorkbookResourceKeys(opts.newWorkbookName),
          ])
        ),
      });
      this.pendingResourceEvents.push({
        type: "workbook-rename",
        workbookName: opts.workbookName,
        newWorkbookName: opts.newWorkbookName,
      });
    });
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
    return this.withUndoRedoCheckpoint(() => {
      const preparedHeaderUpdates =
        this.tableManager.prepareHeaderUpdatesForSheet({
          ...opts,
          getCellContent: (address) => content.get(getCellReference(address)),
        });
      let replacementContent = content;
      if (preparedHeaderUpdates.generatedHeaders.length > 0) {
        replacementContent = new Map(content);
        for (const generatedHeader of preparedHeaderUpdates.generatedHeaders) {
          replacementContent.set(
            getCellReference(generatedHeader.address),
            generatedHeader.name
          );
        }
      }

      const applyContent = () => {
        this.workbookObserverInvalidationSuppressionDepth++;
        try {
          this.workbookManager.setSheetContent(opts, replacementContent);
        } finally {
          this.workbookObserverInvalidationSuppressionDepth--;
        }
        this.tableManager.applyHeaderUpdates(preparedHeaderUpdates.updates);
        const renamedReferences = this.renameTableHeaderReferences(
          preparedHeaderUpdates.updates
        );
        // Rebuilding a sheet can change formula/spill evaluation precedence
        // even when every serialized cell value is identical. A cache reset is
        // both correct for that operation and avoids allocating a second
        // sheet-sized invalidation footprint.
        this.evaluationManager.clearEvaluationCache();
        this.emitMutation({
          touchedCells: [],
          resourceKeys: Array.from(new Set(renamedReferences.resourceKeys)),
        });
      };

      applyContent();
    });
  }

  /**
   * Set the content of a single cell.
   */
  setCellContent(address: CellAddress, content: SerializedCellValue) {
    return this.withUndoRedoCheckpoint(() => {
      const preparedHeaderUpdate = this.tableManager.prepareHeaderUpdate(
        address,
        content
      );
      const applyContent = () => {
        const previousValue = this.workbookManager.getCellContent(address);
        this.workbookManager.setCellContent(
          address,
          preparedHeaderUpdate.content
        );
        this.tableManager.applyHeaderUpdates(preparedHeaderUpdate.updates);
        const renamedReferences = this.renameTableHeaderReferences(
          preparedHeaderUpdate.updates
        );

        this.emitMutation({
          touchedCells: mergeTouchedCells(
            buildTouchedCells([
              {
                address,
                before: previousValue,
                after: preparedHeaderUpdate.content,
              },
            ]),
            buildFormulaTouchedCells(renamedReferences.changedFormulaCells)
          ),
          resourceKeys: Array.from(new Set(renamedReferences.resourceKeys)),
        });
      };

      applyContent();
    });
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
  ): void {
    return this.withUndoRedoCheckpoint(() => {
      this.batchCopyAncillaryMutations(() => {
        this.autoFillManager.fill(opts, seedRange, fillRanges, direction);
      });

      this.emitMutation({
        touchedCells: [],
        resourceKeys: [],
      });
    });
  }

  /**
   * Removes the content in the spreadsheet that is inside the range.
   */
  clearSpreadsheetRange(address: RangeAddress) {
    return this.withUndoRedoCheckpoint(() => {
      this.workbookManager.clearSpreadsheetRange(address);

      this.emitMutation({
        touchedCells: [],
        resourceKeys: [],
      });
    });
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
      cellDataTypes: this.styleManager.getAllCellDataTypes(),
      rangeMetadata: this.rangeMetadataManager.getAllRangeMetadata(),
      references: this.referenceManager.getAllReferences(),
    };
  }

  onUpdate(listener: () => void) {
    return this.eventManager.onUpdate(listener);
  }

  onWorkbookRename(
    workbookName: string,
    listener: (newWorkbookName: string) => void
  ): () => void {
    return this.eventManager.onWorkbookRename(workbookName, listener);
  }

  onSheetRename(
    opts: { workbookName: string; sheetName: string },
    listener: (newSheetName: string) => void
  ): () => void {
    return this.eventManager.onSheetRename(opts, listener);
  }

  onWorkbookDelete(workbookName: string, listener: () => void): () => void {
    return this.eventManager.onWorkbookDelete(workbookName, listener);
  }

  onSheetDelete(
    opts: { workbookName: string; sheetName: string },
    listener: () => void
  ): () => void {
    return this.eventManager.onSheetDelete(opts, listener);
  }

  private buildSerializedSnapshot(): EngineSnapshot {
    const evaluationSnapshots = this.dependencyManager.toSnapshot(
      this.evaluationManager
    );
    const historySnapshot = this.buildHistorySnapshot();

    return {
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        ...historySnapshot.managers,
        dependency: evaluationSnapshots.dependency,
        cache: evaluationSnapshots.cache,
      },
    };
  }

  private buildHistorySnapshot(): EngineHistorySnapshot {
    const workbookSnapshot = this.workbookManager.toSnapshot();

    return {
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        workbook: workbookSnapshot,
        namedExpression: this.buildNamedExpressionSnapshot(workbookSnapshot),
        table: this.tableManager.toSnapshot(),
        style: this.styleManager.toSnapshot(),
        rangeMetadata: this.rangeMetadataManager.toSnapshot(),
        reference: this.referenceManager.toSnapshot(),
      },
    };
  }

  private buildNamedExpressionSnapshot(
    workbookSnapshot: EngineSnapshot["managers"]["workbook"]
  ): EngineSnapshot["managers"]["namedExpression"] {
    const namedExpressions = this.namedExpressionManager.toSnapshot();
    const workbookExpressions = new Map<string, Map<string, NamedExpression>>();
    const sheetExpressions = new Map<
      string,
      Map<string, Map<string, NamedExpression>>
    >();

    workbookSnapshot.forEach((workbook) => {
      const sourceWorkbookExpressions =
        namedExpressions.workbookExpressions.get(workbook.name);
      workbookExpressions.set(
        workbook.name,
        new Map(sourceWorkbookExpressions ?? [])
      );

      const sourceSheets = namedExpressions.sheetExpressions.get(workbook.name);
      const workbookSheetExpressions = new Map<
        string,
        Map<string, NamedExpression>
      >();
      workbook.sheets.forEach((_, sheetName) => {
        workbookSheetExpressions.set(
          sheetName,
          new Map(sourceSheets?.get(sheetName) ?? [])
        );
      });
      sheetExpressions.set(workbook.name, workbookSheetExpressions);
    });

    return {
      sheetExpressions,
      workbookExpressions,
      globalExpressions: new Map(namedExpressions.globalExpressions),
    };
  }

  serializeEngine(): string {
    return serialize(this.buildSerializedSnapshot());
  }

  private normalizeSupportedSnapshot(
    snapshot: unknown
  ): EngineHistorySnapshot | EngineSnapshot {
    const candidate = snapshot as {
      version?: unknown;
      managers?: Record<string, unknown> & {
        style?: Partial<StyleManagerSnapshot>;
      };
    };

    if (
      !candidate ||
      typeof candidate !== "object" ||
      (candidate.version !== ENGINE_SNAPSHOT_VERSION &&
        candidate.version !== LEGACY_ENGINE_SNAPSHOT_VERSION) ||
      !candidate.managers ||
      !candidate.managers.style
    ) {
      throw new Error(
        `Unsupported serialized engine format. Expected EngineSnapshot version ${LEGACY_ENGINE_SNAPSHOT_VERSION} or ${ENGINE_SNAPSHOT_VERSION}.`
      );
    }

    const style = candidate.managers.style;
    if (
      candidate.version === ENGINE_SNAPSHOT_VERSION &&
      !Array.isArray(style.cellDataTypes)
    ) {
      throw new Error(
        `Unsupported serialized engine format. Expected EngineSnapshot version ${LEGACY_ENGINE_SNAPSHOT_VERSION} or ${ENGINE_SNAPSHOT_VERSION}.`
      );
    }

    return {
      ...candidate,
      version: ENGINE_SNAPSHOT_VERSION,
      managers: {
        ...candidate.managers,
        style: {
          ...style,
          cellDataTypes: style.cellDataTypes ?? [],
        },
      },
    } as unknown as EngineHistorySnapshot | EngineSnapshot;
  }

  private restoreDataManagersFromSnapshot(
    managers: EngineHistorySnapshot["managers"]
  ): void {
    this.namedExpressionManager.clear();
    this.workbookManager.restoreFromSnapshot(managers.workbook);

    managers.workbook.forEach((workbook) => {
      this.namedExpressionManager.addWorkbook(workbook.name);
      workbook.sheets.forEach((sheet) => {
        this.namedExpressionManager.addSheet({
          workbookName: workbook.name,
          sheetName: sheet.name,
        });
      });
    });

    this.namedExpressionManager.restoreFromSnapshot(managers.namedExpression);
    this.tableManager.restoreFromSnapshot(managers.table);
    this.styleManager.restoreFromSnapshot(managers.style);
    this.rangeMetadataManager.restoreFromSnapshot(managers.rangeMetadata);
    this.referenceManager.restoreFromSnapshot(managers.reference);
  }

  resetToSerializedEngine(data: string) {
    this.assertHistoryControlAllowed("reset serialized engine state");
    const deserialized = this.normalizeSupportedSnapshot(
      deserialize(data)
    ) as EngineSnapshot;
    this.restoreDataManagersFromSnapshot(deserialized.managers);
    this.dependencyManager.restoreFromSnapshot(
      {
        dependency: deserialized.managers.dependency,
        cache: deserialized.managers.cache,
      },
      this.evaluationManager
    );

    this.clearUndoRedoHistory();
    this.eventManager.emitUpdate();
  }
  //#endregion
}
