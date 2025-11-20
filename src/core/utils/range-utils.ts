/**
 * Range utility functions for handling range arithmetic and operations
 */

import type { SpreadsheetRange, SpreadsheetRangeEnd } from "../types";

/**
 * Check if two ranges intersect
 */
export function rangesIntersect(
  range1: SpreadsheetRange,
  range2: SpreadsheetRange
): boolean {
  // Get finite end values
  const r1EndCol =
    range1.end.col.type === "number" ? range1.end.col.value : Infinity;
  const r1EndRow =
    range1.end.row.type === "number" ? range1.end.row.value : Infinity;
  const r2EndCol =
    range2.end.col.type === "number" ? range2.end.col.value : Infinity;
  const r2EndRow =
    range2.end.row.type === "number" ? range2.end.row.value : Infinity;

  // Check if ranges overlap
  const colOverlap =
    range1.start.col <= r2EndCol && range2.start.col <= r1EndCol;
  const rowOverlap =
    range1.start.row <= r2EndRow && range2.start.row <= r1EndRow;

  return colOverlap && rowOverlap;
}

/**
 * Check if one range is completely contained within another
 */
export function isRangeContained(
  inner: SpreadsheetRange,
  outer: SpreadsheetRange
): boolean {
  const innerEndCol =
    inner.end.col.type === "number" ? inner.end.col.value : Infinity;
  const innerEndRow =
    inner.end.row.type === "number" ? inner.end.row.value : Infinity;
  const outerEndCol =
    outer.end.col.type === "number" ? outer.end.col.value : Infinity;
  const outerEndRow =
    outer.end.row.type === "number" ? outer.end.row.value : Infinity;

  return (
    inner.start.col >= outer.start.col &&
    inner.start.row >= outer.start.row &&
    innerEndCol <= outerEndCol &&
    innerEndRow <= outerEndRow
  );
}

/**
 * Subtract one range from another, returning the remaining ranges
 * Returns an array of ranges representing `original - subtract`
 * 
 * For infinite ranges in the original:
 * - If subtract range is finite, we still return the infinite portions
 * - This can result in up to 4 ranges (4 sides around the subtracted area)
 */
export function subtractRange(
  original: SpreadsheetRange,
  subtract: SpreadsheetRange
): SpreadsheetRange[] {
  // If ranges don't intersect, return original unchanged
  if (!rangesIntersect(original, subtract)) {
    return [original];
  }

  // If original is completely contained in subtract, return empty
  if (isRangeContained(original, subtract)) {
    return [];
  }

  const result: SpreadsheetRange[] = [];

  // Get finite end values for easier arithmetic
  const origEndCol =
    original.end.col.type === "number" ? original.end.col.value : Infinity;
  const origEndRow =
    original.end.row.type === "number" ? original.end.row.value : Infinity;
  const subEndCol =
    subtract.end.col.type === "number" ? subtract.end.col.value : Infinity;
  const subEndRow =
    subtract.end.row.type === "number" ? subtract.end.row.value : Infinity;

  // Calculate the intersection bounds
  const intersectStartCol = Math.max(original.start.col, subtract.start.col);
  const intersectStartRow = Math.max(original.start.row, subtract.start.row);
  const intersectEndCol = Math.min(origEndCol, subEndCol);
  const intersectEndRow = Math.min(origEndRow, subEndRow);

  // Create up to 4 rectangles around the intersection:
  // Top, Bottom, Left, Right

  // Top rectangle (above the intersection)
  if (original.start.row < intersectStartRow) {
    result.push({
      start: { col: original.start.col, row: original.start.row },
      end: {
        col: createRangeEnd(origEndCol),
        row: createRangeEnd(intersectStartRow - 1),
      },
    });
  }

  // Bottom rectangle (below the intersection)
  if (origEndRow > intersectEndRow) {
    result.push({
      start: { col: original.start.col, row: intersectEndRow + 1 },
      end: {
        col: createRangeEnd(origEndCol),
        row: createRangeEnd(origEndRow),
      },
    });
  }

  // Left rectangle (left of the intersection, within the vertical bounds of intersection)
  if (original.start.col < intersectStartCol) {
    result.push({
      start: { col: original.start.col, row: intersectStartRow },
      end: {
        col: createRangeEnd(intersectStartCol - 1),
        row: createRangeEnd(intersectEndRow),
      },
    });
  }

  // Right rectangle (right of the intersection, within the vertical bounds of intersection)
  if (origEndCol > intersectEndCol) {
    result.push({
      start: { col: intersectEndCol + 1, row: intersectStartRow },
      end: {
        col: createRangeEnd(origEndCol),
        row: createRangeEnd(intersectEndRow),
      },
    });
  }

  return result;
}

/**
 * Helper to create a SpreadsheetRangeEnd from a number (handles infinity)
 */
function createRangeEnd(value: number): SpreadsheetRangeEnd {
  if (!isFinite(value)) {
    return { type: "infinity", sign: "positive" };
  }
  return { type: "number", value };
}

