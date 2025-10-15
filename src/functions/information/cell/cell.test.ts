import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("CELL function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const address = (ref: string) => ({
    sheetName,
    workbookName,
    ...parseCellReference(ref),
  });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test('CELL("row") returns current cell row number', () => {
      setCellContent("B5", '=CELL("row")');
      expect(cell("B5")).toBe(5);
    });

    test('CELL("col") returns current cell column number', () => {
      setCellContent("C5", '=CELL("col")');
      expect(cell("C5")).toBe(3);
    });

    test("CELL handles case-insensitive info_type", () => {
      setCellContent("D10", '=CELL("ROW")');
      expect(cell("D10")).toBe(10);

      setCellContent("D10", '=CELL("COL")');
      expect(cell("D10")).toBe(4);
    });
  });

  describe("error handling", () => {
    test("CELL with no arguments returns error", () => {
      setCellContent("A1", "=CELL()");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("CELL with too many arguments returns error", () => {
      setCellContent("A1", '=CELL("row", A1, A2)');
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("CELL with non-string info_type returns error", () => {
      setCellContent("A1", "=CELL(123)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("CELL with unknown info_type returns error", () => {
      setCellContent("A1", '=CELL("unknown")');
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  describe("additional info_types", () => {
    test('CELL("address") should return cell address', () => {
      setCellContent("B5", '=CELL("address")');
      expect(cell("B5")).toBe("$B$5");
    });

    test('CELL("address", A1) should return cell address', () => {
      setCellContent("B1", '=CELL("address", A1)');
      expect(cell("B1")).toBe("$A$1");
    });

    test('CELL("address", Z10) should return cell address', () => {
      setCellContent("A1", '=CELL("address", Z10)');
      expect(cell("A1")).toBe("$Z$10");
    });

    test('CELL("contents") should return cell value', () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 42],
          ["B1", '=CELL("contents", A1)'],
        ])
      );
      expect(cell("B1")).toBe(42);
    });

    test('CELL("contents") for empty cell should return 0', () => {
      setCellContent("B1", '=CELL("contents", A1)');
      expect(cell("B1")).toBe(0);
    });

    test('CELL("type") should return cell type for number', () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 42],
          ["B1", '=CELL("type", A1)'],
        ])
      );
      expect(cell("B1")).toBe("v");
    });

    test('CELL("type") should return cell type for text', () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["B1", '=CELL("type", A1)'],
        ])
      );
      expect(cell("B1")).toBe("l");
    });

    test('CELL("type") should return cell type for empty cell', () => {
      setCellContent("B1", '=CELL("type", A1)');
      expect(cell("B1")).toBe("b");
    });

    test("CELL with range should use upper-left cell", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["B2", 10],
          ["B3", 20],
          ["C2", 30],
          ["A1", '=CELL("row", B2:C3)'],
        ])
      );
      expect(cell("A1")).toBe(2);
    });
  });
});
