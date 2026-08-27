/**
 * RangeMetadataManager - Manages arbitrary consumer-defined metadata attached
 * to ranges rather than individual cells.
 */

import type {
  CellAddress,
  RangeAddress,
  RangeMetadata,
  RangeMetadataInput,
} from "../types";
import type { RangeMetadataManagerSnapshot } from "../engine-snapshot";
import { isCellInRange } from "../utils";
import { rangesIntersect, subtractRange } from "../utils/range-utils";
import {
  MutationObserverDispatcher,
  applyIndexedChanges,
  type IndexedMutationValue,
  type MutationDirection,
} from "./mutation-observer";

export type RangeMetadataDataChange<TMetadata = unknown> = {
  readonly kind: "range-metadata";
  readonly id: string;
  readonly before?: IndexedMutationValue<RangeMetadata<TMetadata>>;
  readonly after?: IndexedMutationValue<RangeMetadata<TMetadata>>;
};

export type RangeMetadataMutationObserver<TMetadata = unknown> = (
  changes: readonly RangeMetadataDataChange<TMetadata>[]
) => void;

const cloneArea = (area: RangeAddress): RangeAddress => ({
  workbookName: area.workbookName,
  sheetName: area.sheetName,
  range: {
    start: { ...area.range.start },
    end: {
      col: { ...area.range.end.col },
      row: { ...area.range.end.row },
    },
  },
});

const cloneEntry = <TMetadata>(
  entry: RangeMetadata<TMetadata>
): RangeMetadata<TMetadata> => ({
  id: entry.id,
  areas: entry.areas.map(cloneArea),
  metadata: entry.metadata,
});

export class RangeMetadataManager<TMetadata = unknown> {
  private rangeMetadata: RangeMetadata<TMetadata>[] = [];
  private readonly mutationDispatcher: MutationObserverDispatcher<
    RangeMetadataDataChange<TMetadata>
  >;
  private mutationBatchDepth = 0;
  private mutationBatchBefore?: RangeMetadata<TMetadata>[];

  constructor(
    mutationObserver?: RangeMetadataMutationObserver<TMetadata>,
    shouldObserve?: () => boolean,
    detachMutationValues = true,
    private readonly shouldBatchMutations: () => boolean = () => true
  ) {
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
          before.map((entry, index) => [entry.id, { entry, index }])
        );
        const afterById = new Map(
          this.rangeMetadata.map((entry, index) => [entry.id, { entry, index }])
        );
        const changes: RangeMetadataDataChange<TMetadata>[] = [];
        for (const id of new Set([...beforeById.keys(), ...afterById.keys()])) {
          const beforeState = beforeById.get(id);
          const afterState = afterById.get(id);
          if (beforeState?.entry === afterState?.entry) {
            continue;
          }
          changes.push({
            kind: "range-metadata",
            id,
            ...(beforeState
              ? {
                  before: {
                    index: beforeState.index,
                    value: this.mutationDispatcher.retain(beforeState.entry),
                  },
                }
              : {}),
            ...(afterState
              ? {
                  after: {
                    index: afterState.index,
                    value: this.mutationDispatcher.retain(afterState.entry),
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
    this.mutationBatchBefore = [...this.rangeMetadata];
  }

  private takeMutationBatchBefore(): RangeMetadata<TMetadata>[] | undefined {
    const before = this.mutationBatchBefore;
    this.mutationBatchBefore = undefined;
    return before;
  }

  addRangeMetadata(entry: RangeMetadataInput<TMetadata>): string {
    const id = entry.id ?? crypto.randomUUID();
    if (this.rangeMetadata.some((existing) => existing.id === id)) {
      throw new Error(`Range metadata with id "${id}" already exists`);
    }

    const storedEntry: RangeMetadata<TMetadata> = {
      id,
      areas: entry.areas.map(cloneArea),
      metadata: entry.metadata,
    };
    const after = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(storedEntry)
      : undefined;
    const index = this.rangeMetadata.length;
    this.rangeMetadata.push(storedEntry);
    if (after) {
      this.mutationDispatcher.report([
        { kind: "range-metadata", id, after: { index, value: after } },
      ]);
    }

    return id;
  }

  removeRangeMetadata(id: string): boolean {
    const index = this.rangeMetadata.findIndex((entry) => entry.id === id);
    if (index === -1) {
      return false;
    }
    const before = this.mutationDispatcher.observed
      ? this.mutationDispatcher.retain(this.rangeMetadata[index]!)
      : undefined;
    this.rangeMetadata.splice(index, 1);
    if (before) {
      this.mutationDispatcher.report([
        { kind: "range-metadata", id, before: { index, value: before } },
      ]);
    }
    return true;
  }

  getAllRangeMetadata(): RangeMetadata<TMetadata>[] {
    return this.rangeMetadata.map(cloneEntry);
  }

  getRangeMetadataForCell(
    cellAddress: CellAddress
  ): RangeMetadata<TMetadata>[] {
    return this.rangeMetadata
      .filter((entry) =>
        entry.areas.some(
          (area) =>
            area.workbookName === cellAddress.workbookName &&
            area.sheetName === cellAddress.sheetName &&
            isCellInRange(cellAddress, area.range)
        )
      )
      .map(cloneEntry);
  }

  getRangeMetadataIntersectingWithRange(
    range: RangeAddress
  ): RangeMetadata<TMetadata>[] {
    return this.rangeMetadata
      .filter((entry) =>
        entry.areas.some(
          (area) =>
            area.workbookName === range.workbookName &&
            area.sheetName === range.sheetName &&
            rangesIntersect(area.range, range.range)
        )
      )
      .map(cloneEntry);
  }

  clearRangeMetadataInRange(range: RangeAddress): void {
    this.transformEntries((entry) => {
      let changed = false;
      const areas = entry.areas.flatMap((area) => {
        if (
          area.workbookName !== range.workbookName ||
          area.sheetName !== range.sheetName ||
          !rangesIntersect(area.range, range.range)
        ) {
          return [area];
        }

        changed = true;
        return subtractRange(area.range, range.range).map((remainingRange) => ({
          ...area,
          range: remainingRange,
        }));
      });
      if (!changed) {
        return entry;
      }
      return areas.length > 0 ? { ...entry, areas } : undefined;
    });
  }

  removeWorkbookRangeMetadata(workbookName: string): void {
    this.transformEntries((entry) =>
      entry.areas.some((area) => area.workbookName === workbookName)
        ? undefined
        : entry
    );
  }

  removeSheetRangeMetadata(workbookName: string, sheetName: string): void {
    this.transformEntries((entry) =>
      entry.areas.some(
        (area) =>
          area.workbookName === workbookName && area.sheetName === sheetName
      )
        ? undefined
        : entry
    );
  }

  updateWorkbookName(oldName: string, newName: string): void {
    this.transformEntries((entry) => {
      if (!entry.areas.some((area) => area.workbookName === oldName)) {
        return entry;
      }
      return {
        ...entry,
        areas: entry.areas.map((area) =>
          area.workbookName === oldName
            ? { ...area, workbookName: newName }
            : area
        ),
      };
    });
  }

  updateSheetName(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    this.transformEntries((entry) => {
      if (
        !entry.areas.some(
          (area) =>
            area.workbookName === workbookName &&
            area.sheetName === oldSheetName
        )
      ) {
        return entry;
      }
      return {
        ...entry,
        areas: entry.areas.map((area) =>
          area.workbookName === workbookName && area.sheetName === oldSheetName
            ? { ...area, sheetName: newSheetName }
            : area
        ),
      };
    });
  }

  resetRangeMetadata(rangeMetadata?: RangeMetadata<TMetadata>[]): void {
    const next = rangeMetadata ? rangeMetadata.map(cloneEntry) : [];
    if (this.mutationDispatcher.observed) {
      const changes: RangeMetadataDataChange<TMetadata>[] = [];
      const beforeById = new Map(
        this.rangeMetadata.map((entry, index) => [entry.id, { entry, index }])
      );
      const afterById = new Map(
        next.map((entry, index) => [entry.id, { entry, index }])
      );
      for (const id of new Set([...beforeById.keys(), ...afterById.keys()])) {
        const before = beforeById.get(id);
        const after = afterById.get(id);
        if (before?.index === after?.index && before?.entry === after?.entry) {
          continue;
        }
        changes.push({
          kind: "range-metadata",
          id,
          ...(before
            ? {
                before: {
                  index: before.index,
                  value: this.mutationDispatcher.retain(before.entry),
                },
              }
            : {}),
          ...(after
            ? {
                after: {
                  index: after.index,
                  value: this.mutationDispatcher.retain(after.entry),
                },
              }
            : {}),
        });
      }
      this.rangeMetadata = next;
      this.mutationDispatcher.report(changes);
      return;
    }
    this.rangeMetadata = next;
  }

  toSnapshot(): RangeMetadataManagerSnapshot {
    return this.getAllRangeMetadata();
  }

  restoreFromSnapshot(snapshot: RangeMetadataManagerSnapshot): void {
    this.resetRangeMetadata(snapshot as RangeMetadata<TMetadata>[]);
  }

  /** Applies retained deltas directly without notifying the observer. */
  applyHistoryChanges(
    changes: readonly RangeMetadataDataChange<TMetadata>[],
    direction: MutationDirection
  ): void {
    this.rangeMetadata = applyIndexedChanges(
      this.rangeMetadata,
      changes,
      direction
    );
  }

  private transformEntries(
    transform: (
      entry: RangeMetadata<TMetadata>
    ) => RangeMetadata<TMetadata> | undefined
  ): void {
    const next: RangeMetadata<TMetadata>[] = [];
    const changes: RangeMetadataDataChange<TMetadata>[] = [];
    let afterIndex = 0;

    for (
      let beforeIndex = 0;
      beforeIndex < this.rangeMetadata.length;
      beforeIndex++
    ) {
      const before = this.rangeMetadata[beforeIndex]!;
      const after = transform(before);
      if (this.mutationDispatcher.observed && after !== before) {
        changes.push({
          kind: "range-metadata",
          id: before.id,
          before: {
            index: beforeIndex,
            value: this.mutationDispatcher.retain(before),
          },
          ...(after
            ? {
                after: {
                  index: afterIndex,
                  value: this.mutationDispatcher.retain(after),
                },
              }
            : {}),
        });
      }
      if (after) {
        next.push(after);
        afterIndex++;
      }
    }

    this.rangeMetadata = next;
    this.mutationDispatcher.report(changes);
  }
}
