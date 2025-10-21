import { describe, expect, test } from "bun:test";

import type { CellAddress } from "../types";
import { cellAddressToKey, keyToCellAddress } from "../utils";

describe("cell address to key", () => {
  describe("Cell nodes", () => {
    test("should generate unique keys for cell references", () => {
      const cellAddress: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      const key = cellAddressToKey(cellAddress);
      expect(key).toBe("cell-value:Workbook1:Sheet1:A1");
    });

    test("should handle different sheets", () => {
      const cellAddress1: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        colIndex: 10,
        rowIndex: 5,
      };

      const cellAddress2: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet2",
        colIndex: 10,
        rowIndex: 5,
      };

      expect(cellAddressToKey(cellAddress1)).toBe("cell-value:Workbook1:Sheet1:K6");
      expect(cellAddressToKey(cellAddress2)).toBe("cell-value:Workbook1:Sheet2:K6");
      expect(cellAddressToKey(cellAddress1)).not.toBe(
        cellAddressToKey(cellAddress2)
      );
    });

    test("should handle sheets with spaces in names", () => {
      const cellAddress: CellAddress = {
        workbookName: "My Workbook",
        sheetName: "My Sheet Name",
        colIndex: 2,
        rowIndex: 1,
      };

      expect(cellAddressToKey(cellAddress)).toBe(
        "cell-value:My Workbook:My Sheet Name:C2"
      );
    });
  });

  describe("Edge cases", () => {
    test("should handle empty sheet names", () => {
      const cellAddress: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "",
        colIndex: 0,
        rowIndex: 0,
      };

      expect(cellAddressToKey(cellAddress)).toBe("cell-value:Workbook1::A1");
    });

    test("should handle large numbers", () => {
      const cellAddress: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        colIndex: 16383,
        rowIndex: 999999,
      };

      expect(cellAddressToKey(cellAddress)).toBe(
        "cell-value:Workbook1:Sheet1:XFD1000000"
      );
    });

    test("should throw error for undefined rowIndex or colIndex", () => {
      const cellAddressWithUndefinedRow: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        colIndex: 0,
        rowIndex: undefined as any,
      };

      const cellAddressWithUndefinedCol: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        rowIndex: 0,
        colIndex: undefined as any,
      };

      const cellAddressWithBothUndefined: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        colIndex: undefined as any,
        rowIndex: undefined as any,
      };

      expect(() => cellAddressToKey(cellAddressWithUndefinedRow)).toThrow(
        "Invalid cell address: rowIndex and colIndex must be defined"
      );
      expect(() => cellAddressToKey(cellAddressWithUndefinedCol)).toThrow(
        "Invalid cell address: rowIndex and colIndex must be defined"
      );
      expect(() => cellAddressToKey(cellAddressWithBothUndefined)).toThrow(
        "Invalid cell address: rowIndex and colIndex must be defined"
      );
    });
  });
});

describe("keyToDependencyNode", () => {
  describe("Cell nodes", () => {
    test("should parse cell keys correctly", () => {
      const key = "cell:Workbook1:Sheet1:A1";
      const expected: CellAddress = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      expect(keyToCellAddress(key)).toEqual(expected);
    });

    test("should handle different sheets", () => {
      const key1 = "cell:Workbook1:Sheet1:K6";
      const key2 = "cell:Workbook1:Sheet2:K6";

      const node1 = keyToCellAddress(key1);
      const node2 = keyToCellAddress(key2);

      expect(node1.workbookName).toBe("Workbook1");
      expect(node1.sheetName).toBe("Sheet1");
      expect(node2.sheetName).toBe("Sheet2");
      expect(node1.rowIndex).toBe(5);
      expect(node1.colIndex).toBe(10);
    });

    test("should handle sheets with spaces in names", () => {
      const key = "cell:My Workbook:My Sheet Name:C2";
      const node = keyToCellAddress(key);

      expect(node.workbookName).toBe("My Workbook");
      expect(node.sheetName).toBe("My Sheet Name");
      expect(node.rowIndex).toBe(1);
      expect(node.colIndex).toBe(2);
    });

    test("should handle empty sheet names", () => {
      const key = "cell:Workbook1::A1";
      const node = keyToCellAddress(key);

      expect(node.workbookName).toBe("Workbook1");
      expect(node.sheetName).toBe("");
      expect(node.rowIndex).toBe(0);
      expect(node.colIndex).toBe(0);
    });

    test("should handle large numbers", () => {
      const key = "cell:Workbook1:Sheet1:XFD1000000";
      const node = keyToCellAddress(key);

      expect(node.workbookName).toBe("Workbook1");
      expect(node.rowIndex).toBe(999999);
      expect(node.colIndex).toBe(16383);
    });
  });

  describe("Error handling", () => {
    test("should throw error for invalid key format", () => {
      expect(() => keyToCellAddress("invalid")).toThrow(
        "Invalid dependency key format"
      );
    });

    test("should throw error for unknown node type", () => {
      expect(() => keyToCellAddress("unknown:type:here")).toThrow(
        "Invalid cell key format: unknown:type:here"
      );
    });

    test("should throw error for invalid cell key parts", () => {
      expect(() => keyToCellAddress("cell:Workbook1:Sheet1:A")).toThrow(
        "Invalid cell reference"
      );
      expect(() => keyToCellAddress("cell:Workbook1:Sheet1:A1:extra")).toThrow(
        "Invalid cell key format"
      );
    });
  });
});

describe("Roundtrip conversion (node -> key -> node)", () => {
  test("should preserve cell nodes through roundtrip", () => {
    const originalNode: CellAddress = {
      workbookName: "Workbook1",
      sheetName: "Sheet1",
      rowIndex: 5,
      colIndex: 10,
    };

    const key = cellAddressToKey(originalNode);
    const roundtripNode = keyToCellAddress(key);

    expect(roundtripNode).toEqual(originalNode);
  });

  test("should preserve edge cases through roundtrip", () => {
    const emptySheetNode: CellAddress = {
      workbookName: "Workbook1",
      sheetName: "",
      rowIndex: 0,
      colIndex: 0,
    };

    const largeNumberNode: CellAddress = {
      workbookName: "Workbook1",
      sheetName: "Sheet1",
      rowIndex: 999999,
      colIndex: 16383,
    };

    expect(keyToCellAddress(cellAddressToKey(emptySheetNode))).toEqual(
      emptySheetNode
    );

    expect(keyToCellAddress(cellAddressToKey(largeNumberNode))).toEqual(
      largeNumberNode
    );
  });
});
