import { describe, expect, test } from "bun:test";
import {
  subtractRange,
  rangesIntersect,
  isRangeContained,
} from "../../../../src/core/utils/range-utils";
import type { SpreadsheetRange } from "../../../../src/core/types";

describe("range-utils", () => {
  describe("rangesIntersect", () => {
    test("detects overlapping ranges", () => {
      const range1: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 5 }, row: { type: "number", value: 5 } },
      };
      const range2: SpreadsheetRange = {
        start: { col: 3, row: 3 },
        end: { col: { type: "number", value: 7 }, row: { type: "number", value: 7 } },
      };
      expect(rangesIntersect(range1, range2)).toBe(true);
    });

    test("detects non-overlapping ranges", () => {
      const range1: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
      };
      const range2: SpreadsheetRange = {
        start: { col: 5, row: 5 },
        end: { col: { type: "number", value: 7 }, row: { type: "number", value: 7 } },
      };
      expect(rangesIntersect(range1, range2)).toBe(false);
    });

    test("handles infinite ranges", () => {
      const finite: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 5 }, row: { type: "number", value: 5 } },
      };
      const infinite: SpreadsheetRange = {
        start: { col: 3, row: 3 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "infinity", sign: "positive" } },
      };
      expect(rangesIntersect(finite, infinite)).toBe(true);
    });
  });

  describe("isRangeContained", () => {
    test("detects contained range", () => {
      const inner: SpreadsheetRange = {
        start: { col: 2, row: 2 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      };
      const outer: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 10 }, row: { type: "number", value: 10 } },
      };
      expect(isRangeContained(inner, outer)).toBe(true);
    });

    test("detects not contained range", () => {
      const inner: SpreadsheetRange = {
        start: { col: 2, row: 2 },
        end: { col: { type: "number", value: 12 }, row: { type: "number", value: 12 } },
      };
      const outer: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 10 }, row: { type: "number", value: 10 } },
      };
      expect(isRangeContained(inner, outer)).toBe(false);
    });
  });

  describe("subtractRange", () => {
    test("returns original if ranges don't intersect", () => {
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 5, row: 5 },
        end: { col: { type: "number", value: 7 }, row: { type: "number", value: 7 } },
      };
      const result = subtractRange(original, subtract);
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual(original);
    });

    test("returns empty array if original is completely contained", () => {
      const original: SpreadsheetRange = {
        start: { col: 2, row: 2 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 10 }, row: { type: "number", value: 10 } },
      };
      const result = subtractRange(original, subtract);
      expect(result).toHaveLength(0);
    });

    test("creates hole in middle - splits into 4 rectangles", () => {
      // Original: A1:E5 (0,0 to 4,4)
      // Subtract: B2:D4 (1,1 to 3,3)
      // Result: Top (A1:E1), Bottom (A5:E5), Left (A2:A4), Right (E2:E4)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 1, row: 1 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 3 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(4);
      
      // Top: A1:E1 (0,0 to 4,0)
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 0 } },
      });
      
      // Bottom: A5:E5 (0,4 to 4,4)
      expect(result[1]).toEqual({
        start: { col: 0, row: 4 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      });
      
      // Left: A2:A4 (0,1 to 0,3)
      expect(result[2]).toEqual({
        start: { col: 0, row: 1 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 3 } },
      });
      
      // Right: E2:E4 (4,1 to 4,3)
      expect(result[3]).toEqual({
        start: { col: 4, row: 1 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 3 } },
      });
    });

    test("shrinks from top edge", () => {
      // Original: A1:C5 (0,0 to 2,4)
      // Subtract: A1:C2 (0,0 to 2,1)
      // Result: A3:C5 (0,2 to 2,4)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 4 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 1 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 4 } },
      });
    });

    test("shrinks from bottom edge", () => {
      // Original: A1:C5 (0,0 to 2,4)
      // Subtract: A4:C5 (0,3 to 2,4)
      // Result: A1:C3 (0,0 to 2,2)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 4 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 0, row: 3 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 4 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
      });
    });

    test("shrinks from left edge", () => {
      // Original: A1:E3 (0,0 to 4,2)
      // Subtract: A1:B3 (0,0 to 1,2)
      // Result: C1:E3 (2,0 to 4,2)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 2 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 2 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        start: { col: 2, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 2 } },
      });
    });

    test("shrinks from right edge", () => {
      // Original: A1:E3 (0,0 to 4,2)
      // Subtract: D1:E3 (3,0 to 4,2)
      // Result: A1:C3 (0,0 to 2,2)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 2 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 3, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 2 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
      });
    });

    test("handles corner overlap - creates L-shape (2 rectangles)", () => {
      // Original: A1:C3 (0,0 to 2,2)
      // Subtract: B2:D4 (1,1 to 3,3) - overlaps bottom-right corner
      // Result: Top (A1:C1), Left (A2:A3)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 1, row: 1 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 3 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(2);
      
      // Top strip
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 0 } },
      });
      
      // Left strip
      expect(result[1]).toEqual({
        start: { col: 0, row: 1 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 2 } },
      });
    });

    test("handles vertical strip subtraction", () => {
      // Original: A1:E5 (0,0 to 4,4)
      // Subtract: B1:D5 (1,0 to 3,4) - full height, partial width
      // Result: Left (A1:A5), Right (E1:E5)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 4 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(2);
      
      // Left
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 4 } },
      });
      
      // Right
      expect(result[1]).toEqual({
        start: { col: 4, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      });
    });

    test("handles horizontal strip subtraction", () => {
      // Original: A1:E5 (0,0 to 4,4)
      // Subtract: A2:E4 (0,1 to 4,3) - full width, partial height
      // Result: Top (A1:E1), Bottom (A5:E5)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 0, row: 1 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 3 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(2);
      
      // Top
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 0 } },
      });
      
      // Bottom
      expect(result[1]).toEqual({
        start: { col: 0, row: 4 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      });
    });

    test("handles infinite original range with finite subtract", () => {
      // Original: A1:INFINITY (0,0 to ∞,∞)
      // Subtract: B2:D4 (1,1 to 3,3)
      // Result: 4 ranges (top, bottom, left, right)
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "infinity", sign: "positive" } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 1, row: 1 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 3 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(4);
      
      // Top: A1:∞1 (0,0 to ∞,0)
      expect(result[0]).toEqual({
        start: { col: 0, row: 0 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "number", value: 0 } },
      });
      
      // Bottom: A4:∞∞ (0,4 to ∞,∞)
      expect(result[1]).toEqual({
        start: { col: 0, row: 4 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "infinity", sign: "positive" } },
      });
      
      // Left: A2:A4 (0,1 to 0,3)
      expect(result[2]).toEqual({
        start: { col: 0, row: 1 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 3 } },
      });
      
      // Right: D2:∞4 (4,1 to ∞,3)
      expect(result[3]).toEqual({
        start: { col: 4, row: 1 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "number", value: 3 } },
      });
    });

    test("handles single cell subtraction", () => {
      // Original: A1:C3 (0,0 to 2,2)
      // Subtract: B2:B2 (1,1 to 1,1) - single cell
      // Result: Top, Bottom, Left, Right
      const original: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 1, row: 1 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(4);
    });

    test("handles edge case - subtract extends beyond original on one side", () => {
      // Original: B2:D4 (1,1 to 3,3)
      // Subtract: A3:E5 (0,2 to 4,4) - extends left, right, and bottom
      // Result: Top portion only (B2:D2)
      const original: SpreadsheetRange = {
        start: { col: 1, row: 1 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 3 } },
      };
      const subtract: SpreadsheetRange = {
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
      };
      const result = subtractRange(original, subtract);
      
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        start: { col: 1, row: 1 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 1 } },
      });
    });
  });
});

