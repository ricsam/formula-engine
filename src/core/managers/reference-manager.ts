/**
 * ReferenceManager - Manages tracked references for external elements
 *
 * Allows consumers to create stable references to ranges that automatically
 * update when sheets/workbooks are renamed and become invalid when deleted.
 */

import type { RangeAddress, TrackedReference } from "../types";
import type { ReferenceManagerSnapshot } from "../engine-snapshot";
import {
  MutationObserverDispatcher,
  applyIndexedChanges,
  type IndexedMutationValue,
  type MutationDirection,
} from "./mutation-observer";

export type ReferenceDataChange = {
  readonly kind: "reference";
  readonly id: string;
  readonly before?: IndexedMutationValue<TrackedReference>;
  readonly after?: IndexedMutationValue<TrackedReference>;
};

export type ReferenceMutationObserver = (
  changes: readonly ReferenceDataChange[]
) => void;

export class ReferenceManager {
  private references: Map<string, TrackedReference>;
  private readonly mutationDispatcher: MutationObserverDispatcher<ReferenceDataChange>;
  private mutationBatchDepth = 0;
  private mutationBatchBefore?: TrackedReference[];

  constructor(
    mutationObserver?: ReferenceMutationObserver,
    shouldObserve?: () => boolean,
    detachMutationValues = true,
    private readonly shouldBatchMutations: () => boolean = () => true
  ) {
    this.references = new Map();
    this.mutationDispatcher = new MutationObserverDispatcher(
      mutationObserver,
      shouldObserve,
      () => this.captureMutationBatchBefore(),
      detachMutationValues
    );
  }

  batchMutations<TResult>(callback: () => TResult): TResult {
    if (!this.shouldBatchMutations()) {
      return callback();
    }

    if (!this.mutationDispatcher.observed || this.mutationBatchDepth > 0) {
      this.mutationBatchDepth++;
      try {
        return callback();
      } finally {
        this.mutationBatchDepth--;
      }
    }

    this.mutationBatchBefore = undefined;
    this.mutationBatchDepth++;
    try {
      return this.mutationDispatcher.suppress(callback);
    } finally {
      this.mutationBatchDepth--;
      const before = this.takeMutationBatchBefore();
      if (before) {
        const beforeById = new Map(
          before.map((reference, index) => [reference.id, { reference, index }])
        );
        const afterValues = Array.from(this.references.values());
        const afterById = new Map(
          afterValues.map((reference, index) => [
            reference.id,
            { reference, index },
          ])
        );
        const changes: ReferenceDataChange[] = [];
        for (const id of new Set([...beforeById.keys(), ...afterById.keys()])) {
          const beforeState = beforeById.get(id);
          const afterState = afterById.get(id);
          if (beforeState?.reference === afterState?.reference) {
            continue;
          }
          changes.push({
            kind: "reference",
            id,
            ...(beforeState
              ? {
                  before: {
                    index: beforeState.index,
                    value: this.mutationDispatcher.retain(
                      beforeState.reference
                    ),
                  },
                }
              : {}),
            ...(afterState
              ? {
                  after: {
                    index: afterState.index,
                    value: this.mutationDispatcher.retain(afterState.reference),
                  },
                }
              : {}),
          });
        }
        this.mutationDispatcher.report(changes);
      }
    }
  }

  private captureMutationBatchBefore(): void {
    if (this.mutationBatchDepth === 0 || this.mutationBatchBefore) {
      return;
    }
    this.mutationBatchBefore = Array.from(this.references.values());
  }

  private takeMutationBatchBefore(): TrackedReference[] | undefined {
    const before = this.mutationBatchBefore;
    this.mutationBatchBefore = undefined;
    return before;
  }

  /**
   * Create a new tracked reference
   * Returns UUID for the reference
   */
  createRef(address: RangeAddress): string {
    const uuid = crypto.randomUUID();
    const reference: TrackedReference = {
      id: uuid,
      address: {
        workbookName: address.workbookName,
        sheetName: address.sheetName,
        range: address.range,
      },
      isValid: true,
    };
    const after = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(reference)
      : undefined;
    const index = this.references.size;
    this.references.set(uuid, reference);
    if (after) {
      this.mutationDispatcher.report([
        { kind: "reference", id: uuid, after: { index, value: after } },
      ]);
    }
    return uuid;
  }

  /**
   * Get the current address for a reference
   * Returns undefined if reference doesn't exist or is invalid
   */
  getRefAddress(refId: string): RangeAddress | undefined {
    const ref = this.references.get(refId);
    if (!ref || !ref.isValid) {
      return undefined;
    }
    return {
      workbookName: ref.address.workbookName,
      sheetName: ref.address.sheetName,
      range: ref.address.range,
    };
  }

  /**
   * Delete a reference
   * Returns true if reference was deleted, false if it didn't exist
   */
  deleteRef(refId: string): boolean {
    const reference = this.references.get(refId);
    if (!reference) {
      return false;
    }
    const index = Array.from(this.references.keys()).indexOf(refId);
    const before = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(reference)
      : undefined;
    this.references.delete(refId);
    if (before) {
      this.mutationDispatcher.report([
        { kind: "reference", id: refId, before: { index, value: before } },
      ]);
    }
    return true;
  }

  /**
   * Get all invalid reference IDs
   */
  getInvalidRefs(): string[] {
    const invalid: string[] = [];
    for (const [id, ref] of this.references) {
      if (!ref.isValid) {
        invalid.push(id);
      }
    }
    return invalid;
  }

  /**
   * Update references when sheet is renamed
   */
  updateSheetName(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    const changes: ReferenceDataChange[] = [];
    let index = 0;
    for (const [id, ref] of this.references) {
      if (
        ref.address.workbookName === workbookName &&
        ref.address.sheetName === oldSheetName
      ) {
        const after: TrackedReference = {
          ...ref,
          address: { ...ref.address, sheetName: newSheetName },
        };
        if (this.mutationDispatcher.observed) {
          changes.push(this.buildChangedReference(id, index, ref, after));
        }
        this.references.set(id, after);
      }
      index++;
    }
    this.mutationDispatcher.report(changes);
  }

  /**
   * Update references when workbook is renamed
   */
  updateWorkbookName(oldWorkbookName: string, newWorkbookName: string): void {
    const changes: ReferenceDataChange[] = [];
    let index = 0;
    for (const [id, ref] of this.references) {
      if (ref.address.workbookName === oldWorkbookName) {
        const after: TrackedReference = {
          ...ref,
          address: { ...ref.address, workbookName: newWorkbookName },
        };
        if (this.mutationDispatcher.observed) {
          changes.push(this.buildChangedReference(id, index, ref, after));
        }
        this.references.set(id, after);
      }
      index++;
    }
    this.mutationDispatcher.report(changes);
  }

  /**
   * Mark references as invalid when sheet is removed
   */
  invalidateSheet(workbookName: string, sheetName: string): void {
    const changes: ReferenceDataChange[] = [];
    let index = 0;
    for (const [id, ref] of this.references) {
      if (
        ref.isValid &&
        ref.address.workbookName === workbookName &&
        ref.address.sheetName === sheetName
      ) {
        const after = { ...ref, isValid: false };
        if (this.mutationDispatcher.observed) {
          changes.push(this.buildChangedReference(id, index, ref, after));
        }
        this.references.set(id, after);
      }
      index++;
    }
    this.mutationDispatcher.report(changes);
  }

  /**
   * Mark references as invalid when workbook is removed
   */
  invalidateWorkbook(workbookName: string): void {
    const changes: ReferenceDataChange[] = [];
    let index = 0;
    for (const [id, ref] of this.references) {
      if (ref.isValid && ref.address.workbookName === workbookName) {
        const after = { ...ref, isValid: false };
        if (this.mutationDispatcher.observed) {
          changes.push(this.buildChangedReference(id, index, ref, after));
        }
        this.references.set(id, after);
      }
      index++;
    }
    this.mutationDispatcher.report(changes);
  }

  /**
   * Get all references for serialization
   */
  getAllReferences(): Map<string, TrackedReference> {
    return new Map(this.references);
  }

  /**
   * Restore references from serialization
   */
  resetReferences(refs: Map<string, TrackedReference>): void {
    const next = new Map<string, TrackedReference>();
    for (const [id, ref] of refs) {
      next.set(id, {
        id: ref.id,
        address: {
          workbookName: ref.address.workbookName,
          sheetName: ref.address.sheetName,
          range: ref.address.range,
        },
        isValid: ref.isValid,
      });
    }
    if (this.mutationDispatcher.observed) {
      const changes: ReferenceDataChange[] = [];
      const beforeIndex = new Map(
        Array.from(this.references.keys(), (id, index) => [id, index])
      );
      const afterIndex = new Map(
        Array.from(next.keys(), (id, index) => [id, index])
      );
      for (const id of new Set([...this.references.keys(), ...next.keys()])) {
        const before = this.references.get(id);
        const after = next.get(id);
        const beforePosition = beforeIndex.get(id);
        const afterPosition = afterIndex.get(id);
        if (
          before &&
          after &&
          beforePosition === afterPosition &&
          this.referencesEqual(before, after)
        ) {
          continue;
        }
        changes.push({
          kind: "reference",
          id,
          ...(before && beforePosition !== undefined
            ? {
                before: {
                  index: beforePosition,
                  value: this.mutationDispatcher.retain(before),
                },
              }
            : {}),
          ...(after && afterPosition !== undefined
            ? {
                after: {
                  index: afterPosition,
                  value: this.mutationDispatcher.retain(after),
                },
              }
            : {}),
        });
      }
      this.references = next;
      this.mutationDispatcher.report(changes);
      return;
    }
    this.references = next;
  }

  toSnapshot(): ReferenceManagerSnapshot {
    return this.getAllReferences();
  }

  restoreFromSnapshot(snapshot: ReferenceManagerSnapshot): void {
    this.resetReferences(snapshot);
  }

  /** Applies retained deltas directly without notifying the observer. */
  applyHistoryChanges(
    changes: readonly ReferenceDataChange[],
    direction: MutationDirection
  ): void {
    const values = applyIndexedChanges(
      Array.from(this.references.values()),
      changes,
      direction
    );
    this.references = new Map(
      values.map((reference) => [reference.id, reference])
    );
  }

  private buildChangedReference(
    id: string,
    index: number,
    before: TrackedReference,
    after: TrackedReference
  ): ReferenceDataChange {
    return {
      kind: "reference",
      id,
      before: { index, value: this.mutationDispatcher.retain(before) },
      after: { index, value: this.mutationDispatcher.retain(after) },
    };
  }

  private referencesEqual(
    left: TrackedReference,
    right: TrackedReference
  ): boolean {
    return (
      left.id === right.id &&
      left.isValid === right.isValid &&
      left.address.workbookName === right.address.workbookName &&
      left.address.sheetName === right.address.sheetName &&
      JSON.stringify(left.address.range) === JSON.stringify(right.address.range)
    );
  }
}
