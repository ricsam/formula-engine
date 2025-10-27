import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ADDRESS function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("ADDRESS with row and column returns absolute reference", () => {
      setCellContent("A1", "=ADDRESS(2, 3)");
      expect(cell("A1")).toBe("$C$2");
    });

    test("ADDRESS works with different row and column numbers", () => {
      setCellContent("A1", "=ADDRESS(1, 1)");
      expect(cell("A1")).toBe("$A$1");
      
      setCellContent("A2", "=ADDRESS(10, 26)");
      expect(cell("A2")).toBe("$Z$10");
    });

    test("ADDRESS handles multi-letter columns", () => {
      setCellContent("A1", "=ADDRESS(1, 27)");
      expect(cell("A1")).toBe("$AA$1");
      
      setCellContent("A2", "=ADDRESS(5, 52)");
      expect(cell("A2")).toBe("$AZ$5");
    });
  });

  describe("reference type (abs_num parameter)", () => {
    test("ADDRESS with abs_num=1 returns absolute reference", () => {
      setCellContent("A1", "=ADDRESS(2, 3, 1)");
      expect(cell("A1")).toBe("$C$2");
    });

    test("ADDRESS with abs_num=2 returns absolute row, relative column", () => {
      setCellContent("A1", "=ADDRESS(2, 3, 2)");
      expect(cell("A1")).toBe("C$2");
    });

    test("ADDRESS with abs_num=3 returns relative row, absolute column", () => {
      setCellContent("A1", "=ADDRESS(2, 3, 3)");
      expect(cell("A1")).toBe("$C2");
    });

    test("ADDRESS with abs_num=4 returns relative reference", () => {
      setCellContent("A1", "=ADDRESS(2, 3, 4)");
      expect(cell("A1")).toBe("C2");
    });
  });

  describe("sheet reference", () => {
    test("ADDRESS with sheet_text includes sheet name", () => {
      setCellContent("A1", '=ADDRESS(2, 3, 1, TRUE, "Sheet2")');
      expect(cell("A1")).toBe("Sheet2!$C$2");
    });

    test("ADDRESS with sheet and different abs_num", () => {
      setCellContent("A1", '=ADDRESS(5, 10, 4, TRUE, "Data")');
      expect(cell("A1")).toBe("Data!J5");
    });
  });

  describe("error handling", () => {
    test("ADDRESS with too few arguments returns error", () => {
      setCellContent("A1", "=ADDRESS(1)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("ADDRESS with too many arguments returns error", () => {
      setCellContent("A1", "=ADDRESS(1, 2, 3, TRUE, 'Sheet', 6)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("ADDRESS with non-numeric row returns error", () => {
      setCellContent("A1", '=ADDRESS("text", 2)');
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("ADDRESS with non-numeric column returns error", () => {
      setCellContent("A1", '=ADDRESS(2, "text")');
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });

    test("ADDRESS with zero or negative row returns error", () => {
      setCellContent("A1", "=ADDRESS(0, 2)");
      expect(cell("A1")).toBe(FormulaError.VALUE);

      setCellContent("A2", "=ADDRESS(-1, 2)");
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("ADDRESS with zero or negative column returns error", () => {
      setCellContent("A1", "=ADDRESS(2, 0)");
      expect(cell("A1")).toBe(FormulaError.VALUE);

      setCellContent("A2", "=ADDRESS(2, -1)");
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });

    test("ADDRESS with invalid abs_num returns error", () => {
      setCellContent("A1", "=ADDRESS(2, 3, 5)");
      expect(cell("A1")).toBe(FormulaError.VALUE);
    });
  });

  describe("R1C1 style (not yet implemented)", () => {
    test.skip("ADDRESS with a1=FALSE returns R1C1 style", () => {
      setCellContent("A1", "=ADDRESS(2, 3, 1, FALSE)");
      expect(cell("A1")).toBe("R2C3");
    });
  });
});
