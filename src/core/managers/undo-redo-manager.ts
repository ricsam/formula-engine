import type { HistoryEntry, HistoryStep } from "../history";
import type { UndoRedoOptions, UndoRedoState } from "../types";

const DEFAULT_UNDO_REDO_MAX_ENTRIES = 100;
const DEFAULT_UNDO_REDO_MAX_BYTES = 64 * 1024 * 1024;

export type HistoryRecordResult = "recorded" | "oversized";

function resolvePositiveSafeInteger(
  value: number | undefined,
  fallback: number,
  optionName: string
): number {
  const resolved = value ?? fallback;
  if (!Number.isSafeInteger(resolved) || resolved <= 0) {
    throw new Error(`undoRedo.${optionName} must be a positive safe integer`);
  }
  return resolved;
}

/**
 * Owns incremental undo/redo entries and enforces bounded retention.
 *
 * Applying entries remains the engine's responsibility. Replay moves use a
 * pop followed by the matching `push*FromReplay`; unlike `record`, those pushes
 * never clear the opposite stack.
 */
export class UndoRedoManager<TStep extends HistoryStep = HistoryStep> {
  private undoStack: HistoryEntry<TStep>[] = [];
  private redoStack: HistoryEntry<TStep>[] = [];
  private _undoBytes = 0;
  private _redoBytes = 0;

  readonly maxEntries: number;
  readonly maxBytes: number;

  constructor(options: UndoRedoOptions | undefined) {
    this.maxEntries = resolvePositiveSafeInteger(
      options?.maxEntries,
      DEFAULT_UNDO_REDO_MAX_ENTRIES,
      "maxEntries"
    );
    this.maxBytes = resolvePositiveSafeInteger(
      options?.maxBytes,
      DEFAULT_UNDO_REDO_MAX_BYTES,
      "maxBytes"
    );
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Records a new mutation and invalidates redo history.
   *
   * An entry larger than the complete byte budget creates an undo barrier: all
   * retained history is cleared and the oversized entry is not retained. This
   * prevents a later undo from crossing an unrecorded mutation.
   */
  record(entry: HistoryEntry<TStep>): HistoryRecordResult {
    this.assertValidEntry(entry);
    this.clearRedo();

    if (entry.estimatedBytes > this.maxBytes) {
      this.clearUndo();
      return "oversized";
    }

    this.undoStack.push(entry);
    this._undoBytes += entry.estimatedBytes;
    this.trimUndoToLimits();
    return "recorded";
  }

  /**
   * Marks a committed mutation whose inverse payload could not be retained.
   * History on both sides is cleared so replay can never cross the gap.
   */
  recordBarrier(): void {
    this.clear();
  }

  popUndo(): HistoryEntry<TStep> | undefined {
    const entry = this.undoStack.pop();
    if (entry) {
      this._undoBytes -= entry.estimatedBytes;
    }
    return entry;
  }

  popRedo(): HistoryEntry<TStep> | undefined {
    const entry = this.redoStack.pop();
    if (entry) {
      this._redoBytes -= entry.estimatedBytes;
    }
    return entry;
  }

  pushUndoFromReplay(entry: HistoryEntry<TStep>): void {
    this.assertReplayEntryFits(entry);
    this.undoStack.push(entry);
    this._undoBytes += entry.estimatedBytes;
  }

  pushRedoFromReplay(entry: HistoryEntry<TStep>): void {
    this.assertReplayEntryFits(entry);
    this.redoStack.push(entry);
    this._redoBytes += entry.estimatedBytes;
  }

  clear(): void {
    this.clearUndo();
    this.clearRedo();
  }

  getState(): UndoRedoState {
    return {
      enabled: true,
      canUndo: this.canUndo(),
      canRedo: this.canRedo(),
      undoDepth: this.undoStack.length,
      redoDepth: this.redoStack.length,
      maxEntries: this.maxEntries,
      maxBytes: this.maxBytes,
      undoBytes: this._undoBytes,
      redoBytes: this._redoBytes,
    };
  }

  private assertValidEntry(entry: HistoryEntry<TStep>): void {
    if (
      !Number.isSafeInteger(entry.estimatedBytes) ||
      entry.estimatedBytes < 0
    ) {
      throw new Error(
        "history entry estimatedBytes must be a non-negative safe integer"
      );
    }
  }

  private assertReplayEntryFits(entry: HistoryEntry<TStep>): void {
    this.assertValidEntry(entry);
    if (entry.estimatedBytes > this.maxBytes) {
      throw new Error("cannot replay a history entry larger than maxBytes");
    }
  }

  private trimUndoToLimits(): void {
    while (
      this.undoStack.length > this.maxEntries ||
      this._undoBytes > this.maxBytes
    ) {
      const removed = this.undoStack.shift();
      if (!removed) {
        break;
      }
      this._undoBytes -= removed.estimatedBytes;
    }
  }

  private clearUndo(): void {
    this.undoStack = [];
    this._undoBytes = 0;
  }

  private clearRedo(): void {
    this.redoStack = [];
    this._redoBytes = 0;
  }
}
