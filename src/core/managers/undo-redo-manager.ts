import type { UndoRedoOptions, UndoRedoState } from "../types";

const DEFAULT_UNDO_REDO_MAX_DEPTH = 100;

function resolveUndoRedoOptions(
  options: UndoRedoOptions | undefined
): { maxDepth: number } {
  const maxDepth =
    options && options.maxDepth !== undefined
      ? options.maxDepth
      : DEFAULT_UNDO_REDO_MAX_DEPTH;

  if (!Number.isInteger(maxDepth) || maxDepth <= 0) {
    throw new Error("undoRedo.maxDepth must be a positive integer");
  }

  return { maxDepth };
}

export class UndoRedoManager {
  private undoStack: string[] = [];
  private redoStack: string[] = [];
  readonly maxDepth: number;

  constructor(options: UndoRedoOptions | undefined) {
    const resolved = resolveUndoRedoOptions(options);
    this.maxDepth = resolved.maxDepth;
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  recordMutation(before: string, after: string): void {
    if (before === after) {
      return;
    }

    this.pushUndo(before);
    this.redoStack = [];
  }

  popUndo(): string | undefined {
    return this.undoStack.pop();
  }

  popRedo(): string | undefined {
    return this.redoStack.pop();
  }

  pushUndo(snapshot: string): void {
    this.undoStack.push(snapshot);
    while (this.undoStack.length > this.maxDepth) {
      this.undoStack.shift();
    }
  }

  pushRedo(snapshot: string): void {
    this.redoStack.push(snapshot);
    while (this.redoStack.length > this.maxDepth) {
      this.redoStack.shift();
    }
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  getState(): UndoRedoState {
    return {
      enabled: true,
      canUndo: this.canUndo(),
      canRedo: this.canRedo(),
      undoDepth: this.undoStack.length,
      redoDepth: this.redoStack.length,
      maxDepth: this.maxDepth,
    };
  }
}
