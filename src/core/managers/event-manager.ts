export type ResourceEvent =
  | {
      type: "workbook-rename";
      workbookName: string;
      newWorkbookName: string;
    }
  | {
      type: "sheet-rename";
      workbookName: string;
      sheetName: string;
      newSheetName: string;
    }
  | {
      type: "workbook-delete";
      workbookName: string;
    }
  | {
      type: "sheet-delete";
      workbookName: string;
      sheetName: string;
    };

type WorkbookRenameSubscription = {
  workbookName: string;
  listener: (newWorkbookName: string) => void;
};

type SheetRenameSubscription = {
  workbookName: string;
  sheetName: string;
  listener: (newSheetName: string) => void;
};

type WorkbookDeleteSubscription = {
  workbookName: string;
  listener: () => void;
};

type SheetDeleteSubscription = {
  workbookName: string;
  sheetName: string;
  listener: () => void;
};

export class EventManager {
  private updateListeners: Set<() => void> = new Set();
  private workbookRenameSubscriptions = new Set<WorkbookRenameSubscription>();
  private sheetRenameSubscriptions = new Set<SheetRenameSubscription>();
  private workbookDeleteSubscriptions = new Set<WorkbookDeleteSubscription>();
  private sheetDeleteSubscriptions = new Set<SheetDeleteSubscription>();

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

  /**
   * Subscribe to the identity of one workbook. The subscription follows the
   * workbook through subsequent renames until it is unsubscribed.
   */
  onWorkbookRename(
    workbookName: string,
    listener: (newWorkbookName: string) => void
  ): () => void {
    const subscription = { workbookName, listener };
    this.workbookRenameSubscriptions.add(subscription);
    return () => {
      this.workbookRenameSubscriptions.delete(subscription);
    };
  }

  /**
   * Subscribe to the identity of one sheet. The subscription follows both
   * sheet renames and renames of its containing workbook.
   */
  onSheetRename(
    opts: { workbookName: string; sheetName: string },
    listener: (newSheetName: string) => void
  ): () => void {
    const subscription = { ...opts, listener };
    this.sheetRenameSubscriptions.add(subscription);
    return () => {
      this.sheetRenameSubscriptions.delete(subscription);
    };
  }

  /** Subscribe to deletion of one workbook. */
  onWorkbookDelete(workbookName: string, listener: () => void): () => void {
    const subscription = { workbookName, listener };
    this.workbookDeleteSubscriptions.add(subscription);
    return () => {
      this.workbookDeleteSubscriptions.delete(subscription);
    };
  }

  /**
   * Subscribe to deletion of one sheet. The subscription follows both sheet
   * renames and renames of its containing workbook.
   */
  onSheetDelete(
    opts: { workbookName: string; sheetName: string },
    listener: () => void
  ): () => void {
    const subscription = { ...opts, listener };
    this.sheetDeleteSubscriptions.add(subscription);
    return () => {
      this.sheetDeleteSubscriptions.delete(subscription);
    };
  }

  emitResourceEvents(events: readonly ResourceEvent[]): void {
    for (const event of events) {
      switch (event.type) {
        case "workbook-rename":
          this.emitWorkbookRename(event);
          break;
        case "sheet-rename":
          this.emitSheetRename(event);
          break;
        case "workbook-delete":
          this.emitWorkbookDelete(event);
          break;
        case "sheet-delete":
          this.emitSheetDelete(event);
          break;
      }
    }
  }

  private emitWorkbookRename(
    event: Extract<ResourceEvent, { type: "workbook-rename" }>
  ): void {
    const matchingSubscriptions = Array.from(
      this.workbookRenameSubscriptions
    ).filter(
      (subscription) => subscription.workbookName === event.workbookName
    );

    // Move every subscription to the new identity before invoking callbacks so
    // a listener-triggered rename still reaches the same subscriptions.
    for (const subscription of matchingSubscriptions) {
      subscription.workbookName = event.newWorkbookName;
    }
    for (const subscription of this.sheetRenameSubscriptions) {
      if (subscription.workbookName === event.workbookName) {
        subscription.workbookName = event.newWorkbookName;
      }
    }
    for (const subscription of this.workbookDeleteSubscriptions) {
      if (subscription.workbookName === event.workbookName) {
        subscription.workbookName = event.newWorkbookName;
      }
    }
    for (const subscription of this.sheetDeleteSubscriptions) {
      if (subscription.workbookName === event.workbookName) {
        subscription.workbookName = event.newWorkbookName;
      }
    }
    for (const { listener } of matchingSubscriptions) {
      listener(event.newWorkbookName);
    }
  }

  private emitSheetRename(
    event: Extract<ResourceEvent, { type: "sheet-rename" }>
  ): void {
    const matchingSubscriptions = Array.from(
      this.sheetRenameSubscriptions
    ).filter(
      (subscription) =>
        subscription.workbookName === event.workbookName &&
        subscription.sheetName === event.sheetName
    );

    for (const subscription of matchingSubscriptions) {
      subscription.sheetName = event.newSheetName;
    }
    for (const subscription of this.sheetDeleteSubscriptions) {
      if (
        subscription.workbookName === event.workbookName &&
        subscription.sheetName === event.sheetName
      ) {
        subscription.sheetName = event.newSheetName;
      }
    }
    for (const { listener } of matchingSubscriptions) {
      listener(event.newSheetName);
    }
  }

  private emitWorkbookDelete(
    event: Extract<ResourceEvent, { type: "workbook-delete" }>
  ): void {
    const matchingSubscriptions = Array.from(
      this.workbookDeleteSubscriptions
    ).filter(
      (subscription) => subscription.workbookName === event.workbookName
    );
    for (const { listener } of matchingSubscriptions) {
      listener();
    }
  }

  private emitSheetDelete(
    event: Extract<ResourceEvent, { type: "sheet-delete" }>
  ): void {
    const matchingSubscriptions = Array.from(
      this.sheetDeleteSubscriptions
    ).filter(
      (subscription) =>
        subscription.workbookName === event.workbookName &&
        subscription.sheetName === event.sheetName
    );
    for (const { listener } of matchingSubscriptions) {
      listener();
    }
  }
}
