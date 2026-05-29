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

  addRangeMetadata(entry: RangeMetadataInput<TMetadata>): string {
    const id = entry.id ?? crypto.randomUUID();
    if (this.rangeMetadata.some((existing) => existing.id === id)) {
      throw new Error(`Range metadata with id "${id}" already exists`);
    }

    this.rangeMetadata.push({
      id,
      areas: entry.areas.map(cloneArea),
      metadata: entry.metadata,
    });

    return id;
  }

  removeRangeMetadata(id: string): boolean {
    const beforeLength = this.rangeMetadata.length;
    this.rangeMetadata = this.rangeMetadata.filter((entry) => entry.id !== id);
    return this.rangeMetadata.length !== beforeLength;
  }

  getAllRangeMetadata(): RangeMetadata<TMetadata>[] {
    return this.rangeMetadata.map(cloneEntry);
  }

  getRangeMetadataForCell(cellAddress: CellAddress): RangeMetadata<TMetadata>[] {
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
    this.rangeMetadata = this.rangeMetadata
      .map((entry) => ({
        ...entry,
        areas: entry.areas.flatMap((area) => {
          if (
            area.workbookName !== range.workbookName ||
            area.sheetName !== range.sheetName
          ) {
            return [area];
          }

          return subtractRange(area.range, range.range).map((remainingRange) => ({
            ...area,
            range: remainingRange,
          }));
        }),
      }))
      .filter((entry) => entry.areas.length > 0);
  }

  removeWorkbookRangeMetadata(workbookName: string): void {
    this.rangeMetadata = this.rangeMetadata.filter(
      (entry) => !entry.areas.some((area) => area.workbookName === workbookName)
    );
  }

  removeSheetRangeMetadata(workbookName: string, sheetName: string): void {
    this.rangeMetadata = this.rangeMetadata.filter(
      (entry) =>
        !entry.areas.some(
          (area) =>
            area.workbookName === workbookName && area.sheetName === sheetName
        )
    );
  }

  updateWorkbookName(oldName: string, newName: string): void {
    this.rangeMetadata = this.rangeMetadata.map((entry) => ({
      ...entry,
      areas: entry.areas.map((area) =>
        area.workbookName === oldName ? { ...area, workbookName: newName } : area
      ),
    }));
  }

  updateSheetName(
    workbookName: string,
    oldSheetName: string,
    newSheetName: string
  ): void {
    this.rangeMetadata = this.rangeMetadata.map((entry) => ({
      ...entry,
      areas: entry.areas.map((area) =>
        area.workbookName === workbookName && area.sheetName === oldSheetName
          ? { ...area, sheetName: newSheetName }
          : area
      ),
    }));
  }

  resetRangeMetadata(rangeMetadata?: RangeMetadata<TMetadata>[]): void {
    this.rangeMetadata = rangeMetadata ? rangeMetadata.map(cloneEntry) : [];
  }

  toSnapshot(): RangeMetadataManagerSnapshot {
    return this.getAllRangeMetadata();
  }

  restoreFromSnapshot(snapshot: RangeMetadataManagerSnapshot): void {
    this.resetRangeMetadata(snapshot as RangeMetadata<TMetadata>[]);
  }
}
