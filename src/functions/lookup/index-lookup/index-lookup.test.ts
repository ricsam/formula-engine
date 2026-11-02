import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import {
  FormulaError,
  type SerializedCellValue,
} from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("INDEX function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should return value from single cell array", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "=INDEX(A1, 1, 1)"],
        ])
      );

      expect(cell("B1")).toBe("Apple");
    });

    test("should return value from single cell array with default column", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "=INDEX(A1, 1)"],
        ])
      );

      expect(cell("B1")).toBe("Apple");
    });

    test("should return value from row array", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "Cherry"],
          ["D1", "=INDEX(A1:C1, 1, 2)"],
        ])
      );
      expect(cell("D1", true)).toBe("Banana");
    });

    test("should return value from column array", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", "=INDEX(A1:A3, 2, 1)");

      expect(cell("B1")).toBe("Banana");
    });

    test("should return value from 2D array", () => {
      // Set up 2x3 array
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "Cherry"],
          ["A2", 10],
          ["B2", 20],
          ["C2", 30],
        ])
      );
      setCellContent("D1", "=INDEX(A1:C2, 2, 3)");

      expect(cell("D1")).toBe(30);
    });

    test("should handle large array", () => {
      // Set up 3x3 array
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["B1", 2],
          ["C1", 3],
          ["A2", 4],
          ["B2", 5],
          ["C2", 6],
          ["A3", 7],
          ["B3", 8],
          ["C3", 9],
        ])
      );
      setCellContent("D1", "=INDEX(A1:C3, 3, 2)");

      expect(cell("D1")).toBe(8);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return #VALUE! for string row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=INDEX(A1:A2, "first", 1)');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for string column_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "Banana");
      setCellContent("C1", '=INDEX(A1:B1, 1, "second")');

      expect(cell("C1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for boolean row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, TRUE, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for infinity row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, INFINITY, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("bounds checking", () => {
    test("should return #REF! for row_num out of bounds (too high)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, 3, 1)");

      expect(cell("B1")).toBe(FormulaError.REF);
    });

    test("should return #REF! for column_num out of bounds (too high)", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "=INDEX(A1:B1, 1, 3)"],
        ])
      );

      expect(cell("C1")).toBe(FormulaError.REF);
    });

    test("should return the whole column for row_num 0", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, 0, 1)");

      expect(cell("B1")).toBe("Apple");
      expect(cell("B2")).toBe("Banana");
    });

    test("should return #VALUE! for negative row_num", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", "=INDEX(A1:A2, -1, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #REF! for single cell with wrong indices", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "=INDEX(A1, 2, 1)");

      expect(cell("B1")).toBe(FormulaError.REF);
    });
  });

  describe("error handling", () => {
    test("should return #VALUE! for wrong number of arguments", () => {
      setCellContent("B1", "=INDEX()");
      expect(cell("B1")).toBe(FormulaError.VALUE);

      setCellContent("B2", "=INDEX(A1)");
      expect(cell("B2")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for too many arguments", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "=INDEX(A1, 1, 1, 1)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should handle decimal row_num (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", "=INDEX(A1:A3, 2.9, 1)");

      expect(cell("B1")).toBe("Banana");
    });

    test("should handle decimal column_num (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", "Banana");
      setCellContent("C1", "Cherry");
      setCellContent("D1", "=INDEX(A1:C1, 1, 2.7)");

      expect(cell("D1")).toBe("Banana");
    });
  });

  describe("edge cases", () => {
    test("should work with mixed data types in array", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Text"],
          ["A2", 42],
          ["A3", true],
        ])
      );
      setCellContent("B1", "=INDEX(A1:A3, 2, 1)");
      setCellContent("B2", "=INDEX(A1:A3, 3, 1)");

      expect(cell("B1")).toBe(42);
      expect(cell("B2")).toBe(true);
    });

    test("should handle empty cells in array", () => {
      setCellContent("A1", "Apple");
      // A2 is empty
      setCellContent("A3", "Cherry");
      setCellContent("B1", "=INDEX(A1:A3, 2, 1)");

      expect(cell("B1")).toBe("");
    });
  });

  describe("whole rows and columns", () => {
    beforeEach(() => {
      // A 3x3 block in A1:C3: 1..9 row by row
      const data = new Map<string, SerializedCellValue>();
      ["A", "B", "C"].forEach((col, x) => {
        for (let y = 0; y < 3; y++) data.set(`${col}${y + 1}`, y * 3 + x + 1);
      });
      engine.setSheetContent(sheetAddress, data);
    });

    test("row_num 0 selects a column and column_num 0 a row", () => {
      setCellContent("E1", "=INDEX(A1:C3, 0, 2)");
      setCellContent("G1", "=INDEX(A1:C3, 3, 0)");

      expect(cell("E1")).toBe(2);
      expect(cell("E2")).toBe(5);
      expect(cell("E3")).toBe(8);
      expect(cell("E4")).toBe("");
      expect(cell("G1")).toBe(7);
      expect(cell("H1")).toBe(8);
      expect(cell("I1")).toBe(9);
      expect(cell("G2")).toBe("");
    });

    test("both 0 select the whole array", () => {
      setCellContent("E1", "=SUM(INDEX(A1:C3, 0, 0))");
      expect(cell("E1")).toBe(45);
    });

    test("an omitted row_num or column_num is 0", () => {
      setCellContent("E1", "=SUM(INDEX(A1:C3,, 2))");
      setCellContent("E2", "=SUM(INDEX(A1:C3, 2,))");
      setCellContent("E3", "=SUM(INDEX(A1:C3,,))");
      setCellContent("E4", "=INDEX(A1:A3, 2,)");
      setCellContent("E5", "=INDEX(A1:C1,, 3)");

      expect(cell("E1")).toBe(15);
      expect(cell("E2")).toBe(15);
      expect(cell("E3")).toBe(45);
      // One column or row of a single row or column is a single cell.
      expect(cell("E4")).toBe(4);
      expect(cell("E5")).toBe(3);
    });

    test("a selection from a range reads like that range", () => {
      setCellContent("E1", "=COUNT(INDEX(A1:C3,, 2))");
      setCellContent("E2", "=MAX(INDEX(A:C,, 2))");
      setCellContent("E3", "=SUM(INDEX(A1:C3, 2,) * 10)");

      expect(cell("E1")).toBe(3);
      expect(cell("E2")).toBe(8);
      expect(cell("E3")).toBe(150);
    });

    test("selects from computed arrays", () => {
      setCellContent("E1", "=INDEX(SEQUENCE(3, 3),, 2)");
      setCellContent("F1", "=SUM(INDEX(SEQUENCE(3, 3), 3,))");
      setCellContent("G1", "=INDEX({1,2;3,4}, 2,)");
      setCellContent("F2", "=SUM(-INDEX({1,2;3,4},, 1))");

      expect(cell("E1")).toBe(2);
      expect(cell("E2")).toBe(5);
      expect(cell("E3")).toBe(8);
      expect(cell("F1")).toBe(24);
      expect(cell("G1")).toBe(3);
      expect(cell("H1")).toBe(4);
      expect(cell("F2")).toBe(-4);
    });

    test("is read by ranges that overlap the spilled selection", () => {
      setCellContent("E1", "=INDEX(SEQUENCE(3, 3),, 2)");
      setCellContent("F1", "=SUM(E2:E3)");

      expect(cell("F1")).toBe(13);
    });

    test("rejects an index beyond the array", () => {
      setCellContent("E1", "=INDEX(A1:C3, 0, 4)");
      setCellContent("E2", "=INDEX(A1:C3, 4, 0)");

      expect(cell("E1")).toBe(FormulaError.REF);
      expect(cell("E2")).toBe(FormulaError.REF);
    });

    test("recalculates when the selected cells change", () => {
      setCellContent("E1", "=SUM(INDEX(A1:C3,, 2))");
      expect(cell("E1")).toBe(15);

      setCellContent("B2", 50);
      expect(cell("E1")).toBe(60);
    });
  });
});
