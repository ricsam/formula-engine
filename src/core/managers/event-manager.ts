export class EventManager {
  private updateListeners: Set<() => void> = new Set();

  /**
   * Register listener for batched sheet updates. Returns an unsubscribe function.
   */
  onUpdate(listener: () => void): () => void {
    this.updateListeners.add(listener);
    return () => {
      this.updateListeners.delete(listener);
    };
  }

  emitUpdate(): void {
    this.updateListeners.forEach((listener) => listener());
  }
}
