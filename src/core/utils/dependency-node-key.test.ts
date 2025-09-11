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
        type: "cell",
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
        type: "cell",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: 5, colIndex: 10 },
      };

      const cellNode2: DependencyNode = {
        type: "cell",
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
        type: "cell",
        workbookName: "My Workbook",
        sheetName: "My Sheet Name",
        address: { rowIndex: 1, colIndex: 2 },
      };

      expect(dependencyNodeToKey(cellNode)).toBe("cell:My Workbook:My Sheet Name:C2");
    });
  });

  describe("Range nodes", () => {
    test("should generate keys for finite ranges", () => {
      const rangeNode: DependencyNode = {
        type: "range",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 9 },
            col: { type: "number", value: 4 },
          },
        },
      };

      expect(dependencyNodeToKey(rangeNode)).toBe("range:Workbook1:Sheet1:A1:E10");
    });

    test("should generate keys for infinite column ranges", () => {
      const rangeNode: DependencyNode = {
        type: "range",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "infinity", sign: "positive" },
            col: { type: "number", value: 0 },
          },
        },
      };

      expect(dependencyNodeToKey(rangeNode)).toBe(
        "range:Workbook1:Sheet1:A1:A"
      );
    });

    test("should generate keys for infinite row ranges", () => {
      const rangeNode: DependencyNode = {
        type: "range",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 0 },
            col: { type: "infinity", sign: "positive" },
          },
        },
      };

      expect(dependencyNodeToKey(rangeNode)).toBe(
        "range:Workbook1:Sheet1:A1:1"
      );
    });

    test("should generate keys for fully infinite ranges", () => {
      const rangeNode: DependencyNode = {
        type: "range",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "infinity", sign: "positive" },
            col: { type: "infinity", sign: "positive" },
          },
        },
      };

      expect(dependencyNodeToKey(rangeNode)).toBe(
        "range:Workbook1:Sheet1:A1:INFINITY"
      );
    });
  });

  describe("Multi-spreadsheet range nodes", () => {
    test("should generate keys for list-based multi-sheet ranges", () => {
      const multiRangeNode: DependencyNode = {
        type: "multi-spreadsheet-range",
        ranges: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 9 },
            col: { type: "number", value: 4 },
          },
        },
        sheetNames: {
          type: "list",
          list: ["Sheet1", "Sheet2", "Sheet3"],
        },
      };

      expect(dependencyNodeToKey(multiRangeNode)).toBe(
        "multi-range:list:Sheet1,Sheet2,Sheet3:A1:E10"
      );
    });

    test("should generate keys for range-based multi-sheet ranges", () => {
      const multiRangeNode: DependencyNode = {
        type: "multi-spreadsheet-range",
        ranges: {
          start: { row: 1, col: 1 },
          end: {
            row: { type: "number", value: 5 },
            col: { type: "number", value: 5 },
          },
        },
        sheetNames: {
          type: "range",
          startSpreadsheetName: "Q1",
          endSpreadsheetName: "Q4",
        },
      };

      expect(dependencyNodeToKey(multiRangeNode)).toBe(
        "multi-range:range:Q1:Q4:B2:F6"
      );
    });

    test("should handle infinite ranges in multi-sheet scenarios", () => {
      const multiRangeNode: DependencyNode = {
        type: "multi-spreadsheet-range",
        ranges: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "infinity", sign: "positive" },
            col: { type: "infinity", sign: "positive" },
          },
        },
        sheetNames: {
          type: "list",
          list: ["Sheet1"],
        },
      };

      expect(dependencyNodeToKey(multiRangeNode)).toBe(
        "multi-range:list:Sheet1:A1:INFINITY"
      );
    });
  });

  describe("Named expression nodes", () => {
    test("should generate keys for global named expressions", () => {
      const namedNode: DependencyNode = {
        type: "named-expression",
        name: "SALES_TAX",
        scope: { type: "global" },
      };

      expect(dependencyNodeToKey(namedNode)).toBe("named:global:SALES_TAX");
    });

    test("should generate keys for workbook-scoped named expressions", () => {
      const namedNode: DependencyNode = {
        type: "named-expression",
        name: "WORKBOOK_RATE",
        scope: { type: "workbook", workbookName: "Workbook1" },
      };

      expect(dependencyNodeToKey(namedNode)).toBe("named:workbook:Workbook1:WORKBOOK_RATE");
    });

    test("should generate keys for sheet-scoped named expressions", () => {
      const namedNode: DependencyNode = {
        type: "named-expression",
        name: "LOCAL_RATE",
        scope: { type: "sheet", workbookName: "Workbook1", sheetName: "Sheet1" },
      };

      expect(dependencyNodeToKey(namedNode)).toBe("named:sheet:Workbook1:Sheet1:LOCAL_RATE");
    });
  });

  describe("Table nodes", () => {
    test("should generate keys for table header areas", () => {
      const tableNode: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: { kind: "Headers" },
      };

      expect(dependencyNodeToKey(tableNode)).toBe(
        "table:Workbook1:Sheet1:SalesData:Headers"
      );
    });

    test("should generate keys for table data areas", () => {
      const tableNode: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: { kind: "AllData" },
      };

      expect(dependencyNodeToKey(tableNode)).toBe(
        "table:Workbook1:Sheet1:SalesData:AllData"
      );
    });

    test("should generate keys for table data with specific columns", () => {
      const tableNode: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: {
          kind: "Data",
          columns: ["Product", "Sales", "Region"],
          isCurrentRow: false,
        },
      };

      expect(dependencyNodeToKey(tableNode)).toBe(
        "table:Workbook1:Sheet1:SalesData:data:Product,Sales,Region"
      );
    });

    test("should generate keys for all table areas", () => {
      const tableNode: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: { kind: "All" },
      };

      expect(dependencyNodeToKey(tableNode)).toBe("table:Workbook1:Sheet1:SalesData:All");
    });

    test("should handle different table names on same sheet", () => {
      const table1: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "Table1",
        sheetName: "Sheet1",
        area: { kind: "All" },
      };

      const table2: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "Table2",
        sheetName: "Sheet1",
        area: { kind: "All" },
      };

      expect(dependencyNodeToKey(table1)).toBe("table:Workbook1:Sheet1:Table1:All");
      expect(dependencyNodeToKey(table2)).toBe("table:Workbook1:Sheet1:Table2:All");
      expect(dependencyNodeToKey(table1)).not.toBe(dependencyNodeToKey(table2));
    });
  });

  describe("Key uniqueness", () => {
    test("should generate unique keys for different node types", () => {
      const cellNode: DependencyNode = {
        type: "cell",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: 0, colIndex: 0 },
      };

      const rangeNode: DependencyNode = {
        type: "range",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 0 },
            col: { type: "number", value: 0 },
          },
        },
      };

      const namedNode: DependencyNode = {
        type: "named-expression",
        name: "Sheet1",
        scope: { type: "sheet", workbookName: "Workbook1", sheetName: "0" },
      };

      const cellKey = dependencyNodeToKey(cellNode);
      const rangeKey = dependencyNodeToKey(rangeNode);
      const namedKey = dependencyNodeToKey(namedNode);

      expect(cellKey).not.toBe(rangeKey);
      expect(cellKey).not.toBe(namedKey);
      expect(rangeKey).not.toBe(namedKey);
    });
  });

  describe("Edge cases", () => {
    test("should handle empty sheet names", () => {
      const cellNode: DependencyNode = {
        type: "cell",
        workbookName: "Workbook1",
        sheetName: "",
        address: { rowIndex: 0, colIndex: 0 },
      };

      expect(dependencyNodeToKey(cellNode)).toBe("cell:Workbook1::A1");
    });

    test("should handle special characters in names", () => {
      const namedNode: DependencyNode = {
        type: "named-expression",
        name: "MY_NAME!@#$%^&*()",
        scope: { type: "sheet", workbookName: "Workbook1", sheetName: "Sheet1" },
      };

      expect(dependencyNodeToKey(namedNode)).toBe(
        "named:sheet:Workbook1:Sheet1:MY_NAME!@#$%^&*()"
      );
    });

    test("should handle large numbers", () => {
      const cellNode: DependencyNode = {
        type: "cell",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: 999999, colIndex: 16383 },
      };

      expect(dependencyNodeToKey(cellNode)).toBe("cell:Workbook1:Sheet1:XFD1000000");
    });

    test("should throw error for undefined rowIndex or colIndex", () => {
      const cellNodeWithUndefinedRow: DependencyNode = {
        type: "cell" as const,
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: { rowIndex: undefined as any, colIndex: 0 },
      };

      const cellNodeWithUndefinedCol: DependencyNode = {
        type: "cell" as const,
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        address: {
          rowIndex: 0,
          colIndex: undefined as any,
        },
      };

      const cellNodeWithBothUndefined: DependencyNode = {
        type: "cell" as const,
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
        type: "cell",
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

      expect(node1.type).toBe("cell");
      expect(node2.type).toBe("cell");
      if (node1.type === "cell" && node2.type === "cell") {
        expect(node1.workbookName).toBe("Workbook1");
        expect(node1.sheetName).toBe("Sheet1");
        expect(node2.sheetName).toBe("Sheet2");
        expect(node1.address.rowIndex).toBe(5);
        expect(node1.address.colIndex).toBe(10);
      }
    });

    test("should handle sheets with spaces in names", () => {
      const key = "cell:My Workbook:My Sheet Name:C2";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("cell");
      if (node.type === "cell") {
        expect(node.workbookName).toBe("My Workbook");
        expect(node.sheetName).toBe("My Sheet Name");
        expect(node.address.rowIndex).toBe(1);
        expect(node.address.colIndex).toBe(2);
      }
    });

    test("should handle empty sheet names", () => {
      const key = "cell:Workbook1::A1";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("cell");
      if (node.type === "cell") {
        expect(node.workbookName).toBe("Workbook1");
        expect(node.sheetName).toBe("");
        expect(node.address.rowIndex).toBe(0);
        expect(node.address.colIndex).toBe(0);
      }
    });

    test("should handle large numbers", () => {
      const key = "cell:Workbook1:Sheet1:XFD1000000";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("cell");
      if (node.type === "cell") {
        expect(node.workbookName).toBe("Workbook1");
        expect(node.address.rowIndex).toBe(999999);
        expect(node.address.colIndex).toBe(16383);
      }
    });
  });

  describe("Range nodes", () => {
    test("should parse finite range keys", () => {
      const key = "range:Workbook1:Sheet1:A1:E10";
      const expected: DependencyNode = {
        type: "range",
        workbookName: "Workbook1",
        sheetName: "Sheet1",
        range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 9 },
            col: { type: "number", value: 4 },
          },
        },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse infinite column range keys", () => {
      const key = "range:Workbook1:Sheet1:A1:A";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("range");
      if (node.type === "range") {
        expect(node.workbookName).toBe("Workbook1");
        expect(node.range.end.row.type).toBe("infinity");
        if (node.range.end.row.type === "infinity") {
          expect(node.range.end.row.sign).toBe("positive");
        }
        expect(node.range.end.col).toEqual({ type: "number", value: 0 });
      }
    });

    test("should parse infinite row range keys", () => {
      const key = "range:Workbook1:Sheet1:A1:1";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("range");
      if (node.type === "range") {
        expect(node.workbookName).toBe("Workbook1");
        expect(node.range.end.row).toEqual({ type: "number", value: 0 });
        expect(node.range.end.col.type).toBe("infinity");
        if (node.range.end.col.type === "infinity") {
          expect(node.range.end.col.sign).toBe("positive");
        }
      }
    });

    test("should parse fully infinite range keys", () => {
      const key = "range:Workbook1:Sheet1:A1:INFINITY";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("range");
      if (node.type === "range") {
        expect(node.workbookName).toBe("Workbook1");
        expect(node.range.end.row.type).toBe("infinity");
        expect(node.range.end.col.type).toBe("infinity");
        if (node.range.end.row.type === "infinity") {
          expect(node.range.end.row.sign).toBe("positive");
        }
        if (node.range.end.col.type === "infinity") {
          expect(node.range.end.col.sign).toBe("positive");
        }
      }
    });
  });

  describe("Multi-spreadsheet range nodes", () => {
    test("should parse list-based multi-sheet range keys", () => {
      const key = "multi-range:list:Sheet1,Sheet2,Sheet3:A1:E10";
      const expected: DependencyNode = {
        type: "multi-spreadsheet-range",
        ranges: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 9 },
            col: { type: "number", value: 4 },
          },
        },
        sheetNames: {
          type: "list",
          list: ["Sheet1", "Sheet2", "Sheet3"],
        },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse range-based multi-sheet range keys", () => {
      const key = "multi-range:range:Q1:Q4:B2:F6";
      const expected: DependencyNode = {
        type: "multi-spreadsheet-range",
        ranges: {
          start: { row: 1, col: 1 },
          end: {
            row: { type: "number", value: 5 },
            col: { type: "number", value: 5 },
          },
        },
        sheetNames: {
          type: "range",
          startSpreadsheetName: "Q1",
          endSpreadsheetName: "Q4",
        },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse infinite ranges in multi-sheet scenarios", () => {
      const key = "multi-range:list:Sheet1:A1:INFINITY";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("multi-spreadsheet-range");
      if (node.type === "multi-spreadsheet-range") {
        expect(node.ranges.end.row.type).toBe("infinity");
        expect(node.ranges.end.col.type).toBe("infinity");
        if (node.ranges.end.row.type === "infinity") {
          expect(node.ranges.end.row.sign).toBe("positive");
        }
        if (node.ranges.end.col.type === "infinity") {
          expect(node.ranges.end.col.sign).toBe("positive");
        }
        expect(node.sheetNames.type).toBe("list");
        if (node.sheetNames.type === "list") {
          expect(node.sheetNames.list).toEqual(["Sheet1"]);
        }
      }
    });

    test("should handle empty sheet list", () => {
      const key = "multi-range:list::A1:E10";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("multi-spreadsheet-range");
      if (
        node.type === "multi-spreadsheet-range" &&
        node.sheetNames.type === "list"
      ) {
        expect(node.sheetNames.list).toEqual([]);
      }
    });
  });

  describe("Named expression nodes", () => {
    test("should parse global named expression keys", () => {
      const key = "named:global:SALES_TAX";
      const expected: DependencyNode = {
        type: "named-expression",
        name: "SALES_TAX",
        scope: { type: "global" },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse workbook-scoped named expression keys", () => {
      const key = "named:workbook:Workbook1:WORKBOOK_RATE";
      const expected: DependencyNode = {
        type: "named-expression",
        name: "WORKBOOK_RATE",
        scope: { type: "workbook", workbookName: "Workbook1" },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse sheet-scoped named expression keys", () => {
      const key = "named:sheet:Workbook1:Sheet1:LOCAL_RATE";
      const expected: DependencyNode = {
        type: "named-expression",
        name: "LOCAL_RATE",
        scope: { type: "sheet", workbookName: "Workbook1", sheetName: "Sheet1" },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

  });

  describe("Table nodes", () => {
    test("should parse table header area keys", () => {
      const key = "table:Workbook1:Sheet1:SalesData:headers";
      const expected: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: { kind: "Headers" },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse table totals area keys", () => {
      const key = "table:Workbook1:Sheet1:SalesData:AllData";
      const expected: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: { kind: "AllData" },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse table data area keys with columns", () => {
      const key = "table:Workbook1:Sheet1:SalesData:data:Product,Sales,Region";
      const expected: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: {
          kind: "Data",
          columns: ["Product", "Sales", "Region"],
          isCurrentRow: false,
        },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should parse table all area keys", () => {
      const key = "table:Workbook1:Sheet1:SalesData:all";
      const expected: DependencyNode = {
        type: "table",
        workbookName: "Workbook1",
        tableName: "SalesData",
        sheetName: "Sheet1",
        area: { kind: "All" },
      };

      expect(keyToDependencyNode(key)).toEqual(expected);
    });

    test("should handle empty columns list in data area", () => {
      const key = "table:Workbook1:Sheet1:SalesData:data:";
      const node = keyToDependencyNode(key);

      expect(node.type).toBe("table");
      if (node.type === "table" && node.area.kind === "Data") {
        expect(node.workbookName).toBe("Workbook1");
        expect(node.area.columns).toEqual([]);
      }
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
        "Unknown dependency node type"
      );
    });

    test("should throw error for invalid cell key parts", () => {
      expect(() => keyToDependencyNode("cell:Workbook1:Sheet1:A")).toThrow(
        "Invalid cell reference"
      );
      expect(() => keyToDependencyNode("cell:Workbook1:Sheet1:A1:extra")).toThrow(
        "Invalid cell key format"
      );
    });

    test("should throw error for invalid range key parts", () => {
      expect(() => keyToDependencyNode("range:Workbook1:Sheet1:A1:E10:extra")).toThrow(
        "Invalid range key format"
      );
    });

    test("should throw error for invalid multi-range key parts", () => {
      expect(() => keyToDependencyNode("multi-range:list:Sheet1:A1")).toThrow(
        "Invalid multi-range key format"
      );
      expect(() => keyToDependencyNode("multi-range:range:Q1:Q4:B2")).toThrow(
        "Invalid multi-range range key format"
      );
      expect(() =>
        keyToDependencyNode("multi-range:invalid:Q1:Q4:B2:F6")
      ).toThrow("Invalid multi-range sheet names type");
    });

    test("should throw error for invalid named expression key parts", () => {
      expect(() => keyToDependencyNode("named:global")).toThrow(
        "Invalid named expression key format"
      );
      expect(() => keyToDependencyNode("named:workbook:Workbook1")).toThrow(
        "Invalid workbook named expression key format"
      );
      expect(() => keyToDependencyNode("named:sheet:Workbook1:Sheet1")).toThrow(
        "Invalid sheet named expression key format"
      );
      expect(() => keyToDependencyNode("named:invalid:name")).toThrow(
        "Unknown named expression scope type"
      );
    });

    test("should throw error for invalid table key parts", () => {
      expect(() => keyToDependencyNode("table:Workbook1:Sheet1:Table1")).toThrow(
        "Invalid table key format"
      );
      expect(() => keyToDependencyNode("table:Workbook1:Sheet1:Table1:data")).toThrow(
        "Invalid table data key format"
      );
      expect(() => keyToDependencyNode("table:Workbook1:Sheet1:Table1:invalid")).toThrow(
        "Invalid table area type"
      );
    });
  });
});

describe("Roundtrip conversion (node -> key -> node)", () => {
  test("should preserve cell nodes through roundtrip", () => {
    const originalNode: DependencyNode = {
      type: "cell",
      workbookName: "Workbook1",
      sheetName: "Sheet1",
      address: { rowIndex: 5, colIndex: 10 },
    };

    const key = dependencyNodeToKey(originalNode);
    const roundtripNode = keyToDependencyNode(key);

    expect(roundtripNode).toEqual(originalNode);
  });

  test("should preserve range nodes through roundtrip", () => {
    const originalNode: DependencyNode = {
      type: "range",
      workbookName: "Workbook1",
      sheetName: "Sheet1",
              range: {
          start: { row: 0, col: 0 },
          end: {
            row: { type: "number", value: 9 },
            col: { type: "infinity", sign: "positive" },
          },
        },
    };

    const key = dependencyNodeToKey(originalNode);
    const roundtripNode = keyToDependencyNode(key);

    expect(roundtripNode).toEqual(originalNode);
  });

  test("should preserve multi-spreadsheet range nodes through roundtrip", () => {
    const originalNode: DependencyNode = {
      type: "multi-spreadsheet-range",
      ranges: {
        start: { row: 1, col: 1 },
        end: {
          row: { type: "number", value: 5 },
          col: { type: "number", value: 5 },
        },
      },
      sheetNames: {
        type: "range",
        startSpreadsheetName: "Q1",
        endSpreadsheetName: "Q4",
      },
    };

    const key = dependencyNodeToKey(originalNode);
    const roundtripNode = keyToDependencyNode(key);

    expect(roundtripNode).toEqual(originalNode);
  });

  test("should preserve named expression nodes through roundtrip", () => {
    const globalNode: DependencyNode = {
      type: "named-expression",
      name: "SALES_TAX",
      scope: { type: "global" },
    };

    const workbookNode: DependencyNode = {
      type: "named-expression",
      name: "WORKBOOK_RATE",
      scope: { type: "workbook", workbookName: "Workbook1" },
    };

    const sheetNode: DependencyNode = {
      type: "named-expression",
      name: "LOCAL_RATE",
      scope: { type: "sheet", workbookName: "Workbook1", sheetName: "Sheet1" },
    };

    expect(keyToDependencyNode(dependencyNodeToKey(globalNode))).toEqual(
      globalNode
    );
    expect(keyToDependencyNode(dependencyNodeToKey(workbookNode))).toEqual(
      workbookNode
    );
    expect(keyToDependencyNode(dependencyNodeToKey(sheetNode))).toEqual(
      sheetNode
    );
  });

  test("should preserve table nodes through roundtrip", () => {
    const dataNode: DependencyNode = {
      type: "table",
      workbookName: "Workbook1",
      tableName: "SalesData",
      sheetName: "Sheet1",
      area: {
        kind: "Data",
        columns: ["Product", "Sales", "Region"],
        isCurrentRow: false,
      },
    };

    const headerNode: DependencyNode = {
      type: "table",
      workbookName: "Workbook1",
      tableName: "SalesData",
      sheetName: "Sheet1",
      area: { kind: "Headers" },
    };

    expect(keyToDependencyNode(dependencyNodeToKey(dataNode))).toEqual(
      dataNode
    );
    expect(keyToDependencyNode(dependencyNodeToKey(headerNode))).toEqual(
      headerNode
    );
  });

  test("should preserve edge cases through roundtrip", () => {
    const emptySheetNode: DependencyNode = {
      type: "cell",
      workbookName: "Workbook1",
      sheetName: "",
      address: { rowIndex: 0, colIndex: 0 },
    };

    const specialCharNode: DependencyNode = {
      type: "named-expression",
      name: "MY_NAME!@#$%^&*()",
      scope: { type: "sheet", workbookName: "Workbook1", sheetName: "Sheet1" },
    };

    const largeNumberNode: DependencyNode = {
      type: "cell",
      workbookName: "Workbook1",
      sheetName: "Sheet1",
      address: { rowIndex: 999999, colIndex: 16383 },
    };

    expect(keyToDependencyNode(dependencyNodeToKey(emptySheetNode))).toEqual(
      emptySheetNode
    );
    expect(keyToDependencyNode(dependencyNodeToKey(specialCharNode))).toEqual(
      specialCharNode
    );
    expect(keyToDependencyNode(dependencyNodeToKey(largeNumberNode))).toEqual(
      largeNumberNode
    );
  });
});
