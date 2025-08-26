import type {
  CellAddress,
  CellValue,
  FormulaEngineEvents,
  NamedExpression,
  SerializedCellValue,
  Sheet,
  SpreadsheetRange,
  TableDefinition,
} from "./types";

// Simple event emitter for internal use
type EventListener<T = any> = (data: T) => void;

class EventEmitter<T extends Record<string, any>> {
  private listeners: { [K in keyof T]?: EventListener<T[K]>[] } = {};

  on<K extends keyof T>(event: K, listener: EventListener<T[K]>): () => void {
    if (!this.listeners[event]) {
      this.listeners[event] = [];
    }
    this.listeners[event]!.push(listener);

    // Return unsubscribe function
    return () => {
      const listeners = this.listeners[event];
      if (listeners) {
        const index = listeners.indexOf(listener);
        if (index > -1) {
          listeners.splice(index, 1);
        }
      }
    };
  }

  emit<K extends keyof T>(event: K, data: T[K]): void {
    const listeners = this.listeners[event];
    if (listeners) {
      listeners.forEach((listener) => listener(data));
    }
  }

  removeAllListeners(): void {
    this.listeners = {};
  }
}

export class SheetHandler {
  sheets: Map<string, Sheet> = new Map();
  scopedNamedExpressions: Map<string, Map<string, NamedExpression>> = new Map();
  globalNamedExpressions: Map<string, NamedExpression> = new Map();
  tables: Map<string, TableDefinition> = new Map();

  private eventEmitter: EventEmitter<FormulaEngineEvents>;

  cellsUpdateListeners: Map<
    /**
     * sheetName -> listeners
     */
    string,
    Set<() => void>
  > = new Map();

  constructor() {
    this.eventEmitter = new EventEmitter<FormulaEngineEvents>();
  }

  // ===== Event System =====

  /**
   * Subscribe to FormulaEngine events
   * @param event The event name
   * @param listener The event listener function
   * @returns Unsubscribe function
   */
  on<K extends keyof FormulaEngineEvents>(
    event: K,
    listener: EventListener<FormulaEngineEvents[K]>
  ): () => void {
    return this.eventEmitter.on(event, listener);
  }

  /**
   * Subscribe to FormulaEngine events (alias for on)
   * @param event The event name
   * @param listener The event listener function
   * @returns Unsubscribe function
   */
  subscribe<K extends keyof FormulaEngineEvents>(
    event: K,
    listener: EventListener<FormulaEngineEvents[K]>
  ): () => void {
    return this.eventEmitter.on(event, listener);
  }

  /**
   * Remove all event listeners
   */
  removeAllListeners(): void {
    this.eventEmitter.removeAllListeners();
  }

  /**
   * Emit an event (internal use)
   */
  emit<K extends keyof FormulaEngineEvents>(
    event: K,
    data: FormulaEngineEvents[K]
  ): void {
    this.eventEmitter.emit(event, data);
  }

  /**
   * Register listener for batched sheet updates. Returns an unsubscribe function.
   */
  onCellsUpdate(sheetName: string, listener: () => void): () => void {
    if (!this.cellsUpdateListeners.has(sheetName)) {
      this.cellsUpdateListeners.set(sheetName, new Set());
    }
    const set = this.cellsUpdateListeners.get(sheetName)!;
    set.add(listener);
    return () => {
      const listeners = this.cellsUpdateListeners.get(sheetName);
      if (listeners) {
        listeners.delete(listener);
        if (listeners.size === 0) this.cellsUpdateListeners.delete(sheetName);
      }
    };
  }

  getSheetSerialized(sheetName: string): Map<string, SerializedCellValue> {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) return new Map();

    return sheet.content;
  }

  /**
   * Returns true if the range is a single cell, false otherwise
   */
  isRangeOneCell(range: SpreadsheetRange) {
    if (
      range.end.col.type === "infinity" ||
      range.end.row.type === "infinity"
    ) {
      return false;
    }
    return (
      range.start.col === range.end.col.value &&
      range.start.row === range.end.row.value
    );
  }
}
