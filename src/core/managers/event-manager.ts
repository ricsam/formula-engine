import type { FormulaEngineEvents } from "../types";

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

export class EventManager {
  private eventEmitter: EventEmitter<FormulaEngineEvents>;
  private cellsUpdateListeners: Map<
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

  getCellsUpdateListeners(): Map<string, Set<() => void>> {
    return this.cellsUpdateListeners;
  }

  triggerCellsUpdateEvent(): void {
    this.cellsUpdateListeners.forEach((sheetListeners) =>
      sheetListeners.forEach((listener) => listener())
    );
  }

  removeCellsUpdateListenersForSheet(sheetName: string): void {
    this.cellsUpdateListeners.delete(sheetName);
  }

  renameCellsUpdateListenersForSheet(oldName: string, newName: string): void {
    const listeners = this.cellsUpdateListeners.get(oldName);
    if (listeners) {
      this.cellsUpdateListeners.set(newName, listeners);
      this.cellsUpdateListeners.delete(oldName);
    }
  }
}
