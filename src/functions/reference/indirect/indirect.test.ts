import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("INDIRECT function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string | number) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("error handling", () => {
    test("INDIRECT with no arguments returns error", () => {
      setCellContent("A1", "=INDIRECT()");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("INDIRECT with too many arguments returns error", () => {
      setCellContent("A1", '=INDIRECT("A1", TRUE, "extra")');
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("INDIRECT with non-string ref_text returns error", () => {
      setCellContent("A1", "=INDIRECT(123)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  // These tests are now implemented
  describe("basic functionality", () => {
    test("INDIRECT with simple cell reference", () => {
      setCellContent("A1", 42);
      setCellContent("B1", '=INDIRECT("A1")');
      expect(cell("B1")).toBe(42);
    });

    test("INDIRECT with dynamic reference", () => {
      setCellContent("A1", 10);
      setCellContent("A2", 20);
      setCellContent("B1", "A1");
      setCellContent("C1", "=INDIRECT(B1)");
      expect(cell("C1")).toBe(10);
    });

    test("INDIRECT with concatenated reference", () => {
      setCellContent("A5", 100);
      setCellContent("B1", '=INDIRECT("A" & 5)');
      expect(cell("B1")).toBe(100);
    });

    test("INDIRECT with cross-sheet reference", () => {
      engine.addSheet({ workbookName, sheetName: "Sheet2" });
      setCellContent("A1", 50);
      engine.setCellContent(
        { sheetName: "Sheet2", workbookName, ...parseCellReference("B1") },
        '=INDIRECT("TestSheet!A1")'
      );
      expect(
        engine.getCellValue({
          sheetName: "Sheet2",
          workbookName,
          ...parseCellReference("B1"),
        })
      ).toBe(50);
    });

    test("INDIRECT with range reference should spill", () => {
      setCellContent("A1", 1);
      setCellContent("A2", 2);
      setCellContent("A3", 3);
      setCellContent("B1", '=INDIRECT("A1:A3")');
      // Should spill values 1, 2, 3 down from B1
      expect(cell("B1")).toBe(1);
      expect(cell("B2")).toBe(2);
      expect(cell("B3")).toBe(3);
    });
  });

  describe("R1C1 style (not yet implemented)", () => {
    test.skip("INDIRECT with R1C1 style reference", () => {
      setCellContent("C5", 99);
      setCellContent("A1", '=INDIRECT("R5C3", FALSE)');
      expect(cell("A1")).toBe(99);
    });
  });
});
