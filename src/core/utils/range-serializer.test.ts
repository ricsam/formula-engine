import { describe, test, expect } from "bun:test";
import { serializeRange, serializeRangeWithSheet } from "./range-serializer";
import type { SpreadsheetRange } from "../types";

describe("serializeRange", () => {
  test("should serialize finite ranges", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "number", value: 9 }, // Row 10
        col: { type: "number", value: 1 }, // Column B
      },
    };

    expect(serializeRange(range)).toBe("A2:B10");
  });

  test("should serialize column-infinite ranges", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "infinity", sign: "positive" },
        col: { type: "number", value: 1 }, // Column B
      },
    };

    expect(serializeRange(range)).toBe("A2:B");
  });

  test("should serialize row-infinite ranges", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "number", value: 9 }, // Row 10
        col: { type: "infinity", sign: "positive" },
      },
    };

    expect(serializeRange(range)).toBe("A2:10");
  });

  test("should serialize fully infinite ranges", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "infinity", sign: "positive" },
        col: { type: "infinity", sign: "positive" },
      },
    };

    expect(serializeRange(range)).toBe("A2:INFINITY");
  });

  test("should handle single cell ranges", () => {
    const range: SpreadsheetRange = {
      start: { row: 0, col: 0 }, // A1
      end: {
        row: { type: "number", value: 0 }, // Row 1
        col: { type: "number", value: 0 }, // Column A
      },
    };

    expect(serializeRange(range)).toBe("A1:A1");
  });

  test("should handle large column indices", () => {
    const range: SpreadsheetRange = {
      start: { row: 0, col: 25 }, // Z1
      end: {
        row: { type: "number", value: 99 }, // Row 100
        col: { type: "number", value: 701 }, // Column ZZ
      },
    };

    expect(serializeRange(range)).toBe("Z1:ZZ100");
  });

  test("should handle ranges starting from different positions", () => {
    const range: SpreadsheetRange = {
      start: { row: 4, col: 3 }, // D5
      end: {
        row: { type: "infinity", sign: "positive" },
        col: { type: "number", value: 5 }, // Column F
      },
    };

    expect(serializeRange(range)).toBe("D5:F");
  });
});

describe("serializeRangeWithSheet", () => {
  test("should serialize range without sheet name", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "number", value: 9 }, // Row 10
        col: { type: "number", value: 1 }, // Column B
      },
    };

    expect(serializeRangeWithSheet(range)).toBe("A2:B10");
  });

  test("should serialize range with simple sheet name", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "number", value: 9 }, // Row 10
        col: { type: "number", value: 1 }, // Column B
      },
    };

    expect(serializeRangeWithSheet(range, "Sheet1")).toBe("Sheet1!A2:B10");
  });

  test("should quote sheet names with spaces", () => {
    const range: SpreadsheetRange = {
      start: { row: 1, col: 0 }, // A2
      end: {
        row: { type: "number", value: 9 }, // Row 10
        col: { type: "number", value: 1 }, // Column B
      },
    };

    expect(serializeRangeWithSheet(range, "My Sheet")).toBe("'My Sheet'!A2:B10");
  });

  test("should quote sheet names with special characters", () => {
    const range: SpreadsheetRange = {
      start: { row: 0, col: 0 }, // A1
      end: {
        row: { type: "infinity", sign: "positive" },
        col: { type: "infinity", sign: "positive" },
      },
    };

    expect(serializeRangeWithSheet(range, "Sheet-1")).toBe("'Sheet-1'!A1:INFINITY");
  });

  test("should not quote sheet names with underscores and numbers", () => {
    const range: SpreadsheetRange = {
      start: { row: 4, col: 3 }, // D5
      end: {
        row: { type: "number", value: 14 }, // Row 15
        col: { type: "infinity", sign: "positive" },
      },
    };

    expect(serializeRangeWithSheet(range, "Sheet_1")).toBe("Sheet_1!D5:15");
  });

  test("should handle infinite ranges with sheet names", () => {
    const range: SpreadsheetRange = {
      start: { row: 0, col: 0 }, // A1
      end: {
        row: { type: "infinity", sign: "positive" },
        col: { type: "number", value: 2 }, // Column C
      },
    };

    expect(serializeRangeWithSheet(range, "Data")).toBe("Data!A1:C");
  });
});

describe("Edge cases", () => {
  test("should handle maximum Excel column", () => {
    const range: SpreadsheetRange = {
      start: { row: 0, col: 16383 }, // XFD1 (Excel's max column)
      end: {
        row: { type: "number", value: 1048575 }, // Row 1048576 (Excel's max row)
        col: { type: "number", value: 16383 }, // Column XFD
      },
    };

    expect(serializeRange(range)).toBe("XFD1:XFD1048576");
  });

  test("should throw error for invalid range configuration", () => {
    // This shouldn't happen in practice, but testing error handling
    const invalidRange = {
      start: { row: 0, col: 0 },
      end: {
        row: { type: "invalid" as any },
        col: { type: "number", value: 0 },
      },
    } as SpreadsheetRange;

    expect(() => serializeRange(invalidRange)).toThrow("Invalid range end configuration");
  });
});
