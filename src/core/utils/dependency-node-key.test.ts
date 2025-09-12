import { describe, test, expect } from "bun:test";
import {
  dependencyNodeToKey,
  keyToDependencyNode,
} from "./dependency-node-key";
import type { DependencyNode } from "../types";

describe("dependencyNodeToKey", () => {
  describe("Cell nodes", () => {
    test("should generate unique keys for cell references", () => {
      const cellNode: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: {
          rowIndex: 0,
          colIndex: 0,
        },
      };

      const key = dependencyNodeToKey(cellNode);
      expect(key).toBe("cell:Workbook1:Sheet1:A1");
    });

    test("should handle different sheets", () => {
      const cellNode1: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: 5, colIndex: 10 },
      };

      const cellNode2: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet2",
        address: { rowIndex: 5, colIndex: 10 },
      };

      expect(dependencyNodeToKey(cellNode1)).toBe("cell:Workbook1:Sheet1:K6");
      expect(dependencyNodeToKey(cellNode2)).toBe("cell:Workbook1:Sheet2:K6");
      expect(dependencyNodeToKey(cellNode1)).not.toBe(
        dependencyNodeToKey(cellNode2)
      );
    });

    test("should handle sheets with spaces in names", () => {
      const cellNode: DependencyNode = {
        workbookName: "My Workbook",
        sheetName: "My Sheet Name",
        address: { rowIndex: 1, colIndex: 2 },
      };

      expect(dependencyNodeToKey(cellNode)).toBe(
        "cell:My Workbook:My Sheet Name:C2"
      );
    });
  });

  describe("Edge cases", () => {
    test("should handle empty sheet names", () => {
      const cellNode: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "",
        address: { rowIndex: 0, colIndex: 0 },
      };

      expect(dependencyNodeToKey(cellNode)).toBe("cell:Workbook1::A1");
    });

    test("should handle large numbers", () => {
      const cellNode: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: 999999, colIndex: 16383 },
      };

      expect(dependencyNodeToKey(cellNode)).toBe(
        "cell:Workbook1:Sheet1:XFD1000000"
      );
    });

    test("should throw error for undefined rowIndex or colIndex", () => {
      const cellNodeWithUndefinedRow: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: undefined as any, colIndex: 0 },
      };

      const cellNodeWithUndefinedCol: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: {
          rowIndex: 0,
          colIndex: undefined as any,
        },
      };

      const cellNodeWithBothUndefined: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: {
          rowIndex: undefined as any,
          colIndex: undefined as any,
        },
      };

      expect(() => dependencyNodeToKey(cellNodeWithUndefinedRow)).toThrow(
        "Invalid cell address: rowIndex and colIndex must be defined"
      );
      expect(() => dependencyNodeToKey(cellNodeWithUndefinedCol)).toThrow(
        "Invalid cell address: rowIndex and colIndex must be defined"
      );
      expect(() => dependencyNodeToKey(cellNodeWithBothUndefined)).toThrow(
        "Invalid cell address: rowIndex and colIndex must be defined"
      );
    });
  });
});

describe("keyToDependencyNode", () => {
  describe("Cell nodes", () => {
    test("should parse cell keys correctly", () => {
      const key = "cell:Workbook1:Sheet1:A1";
      const expected: DependencyNode = {
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: 0, colIndex: 0 },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should handle different sheets", () => {
      const key1 = "cell:Workbook1:Sheet1:K6";
      const key2 = "cell:Workbook1:Sheet2:K6";

      const node1 = keyToDependencyNode(key1);
      const node2 = keyToDependencyNode(key2);

      expect(node1.workbookName).toBe("Workbook1");
      expect(node1.sheetName).toBe("Sheet1");
      expect(node2.sheetName).toBe("Sheet2");
      expect(node1.address.rowIndex).toBe(5);
      expect(node1.address.colIndex).toBe(10);
    });

    test("should handle sheets with spaces in names", () => {
      const key = "cell:My Workbook:My Sheet Name:C2";
      const node = keyToDependencyNode(key);

      expect(node.workbookName).toBe("My Workbook");
      expect(node.sheetName).toBe("My Sheet Name");
      expect(node.address.rowIndex).toBe(1);
      expect(node.address.colIndex).toBe(2);
    });

    test("should handle empty sheet names", () => {
      const key = "cell:Workbook1::A1";
      const node = keyToDependencyNode(key);

      expect(node.workbookName).toBe("Workbook1");
      expect(node.sheetName).toBe("");
      expect(node.address.rowIndex).toBe(0);
      expect(node.address.colIndex).toBe(0);
    });

    test("should handle large numbers", () => {
      const key = "cell:Workbook1:Sheet1:XFD1000000";
      const node = keyToDependencyNode(key);

      expect(node.workbookName).toBe("Workbook1");
      expect(node.address.rowIndex).toBe(999999);
      expect(node.address.colIndex).toBe(16383);
    });
  });

  describe("Error handling", () => {
    test("should throw error for invalid key format", () => {
      expect(() => keyToDependencyNode("invalid")).toThrow(
        "Invalid dependency key format"
      );
    });

    test("should throw error for unknown node type", () => {
      expect(() => keyToDependencyNode("unknown:type:here")).toThrow(
        "Invalid cell key format: unknown:type:here"
      );
    });

    test("should throw error for invalid cell key parts", () => {
      expect(() => keyToDependencyNode("cell:Workbook1:Sheet1:A")).toThrow(
        "Invalid cell reference"
      );
      expect(() =>
        keyToDependencyNode("cell:Workbook1:Sheet1:A1:extra")
      ).toThrow("Invalid cell key format");
    });
  });
});

describe("Roundtrip conversion (node -> key -> node)", () => {
  test("should preserve cell nodes through roundtrip", () => {
    const originalNode: DependencyNode = {
      workbookName: "Workbook1",
      sheetName: "Sheet1",
      address: { rowIndex: 5, colIndex: 10 },
    };

    const key = dependencyNodeToKey(originalNode);
    const roundtripNode = keyToDependencyNode(key);

    expect(roundtripNode).toEqual(originalNode);
  });

  test("should preserve edge cases through roundtrip", () => {
    const emptySheetNode: DependencyNode = {
      workbookName: "Workbook1",
      sheetName: "",
      address: { rowIndex: 0, colIndex: 0 },
    };

    const largeNumberNode: DependencyNode = {
      workbookName: "Workbook1",
      sheetName: "Sheet1",
      address: { rowIndex: 999999, colIndex: 16383 },
    };

    expect(keyToDependencyNode(dependencyNodeToKey(emptySheetNode))).toEqual(
      emptySheetNode
    );

    expect(keyToDependencyNode(dependencyNodeToKey(largeNumberNode))).toEqual(
      largeNumberNode
    );
  });
});
