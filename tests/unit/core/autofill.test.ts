import { describe, it, expect, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";
import { getCellReference, parseCellReference } from "../../../src/core/utils";
import type { SpreadsheetRange, SerializedCellValue } from "../../../src/core/types";

describe("AutoFill and ClearSpreadsheetRange", () => {
  let engine: FormulaEngine;
  const sheetName = "Sheet1";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("clearSpreadsheetRange", () => {
    it("should clear all cells in a range", () => {
      // Set up some test data
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "A1"],
          ["B1", "B1"],
          ["A2", "A2"],
          ["B2", "B2"],
        ])
      );

      const range: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
      };

      engine.clearSpreadsheetRange(sheetAddress, range);

      // Check that all cells are cleared (engine returns "" for empty cells)
      expect(cell("A1")).toBe("");
      expect(cell("B1")).toBe("");
      expect(cell("A2")).toBe("");
      expect(cell("B2")).toBe("");
    });

    it("should throw error for infinite ranges", () => {
      const infiniteRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "number", value: 1 } }
      };

      expect(() => engine.clearSpreadsheetRange(sheetAddress, infiniteRange))
        .toThrow("Clearing infinite ranges is not supported");
    });

    it("should throw error for non-existent sheet", () => {
      const range: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
      };

      expect(() => engine.clearSpreadsheetRange({ workbookName, sheetName: "NonExistentSheet" }, range))
        .toThrow('Sheet not found');
    });
  });

  describe("autoFill - Single Cell", () => {
    it("should copy number values", () => {
      setCellContent("A1", 42);

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 0 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      expect(cell("B1")).toBe(42);
      expect(cell("C1")).toBe(42);
      expect(cell("D1")).toBe(42);
    });

    it("should copy text values", () => {
      setCellContent("A1", "Apple");

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 1 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 2 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "down");

      expect(cell("A2")).toBe("Apple");
      expect(cell("A3")).toBe("Apple");
    });

    it("should clear cells when seed is blank", () => {
      // Set up some existing data
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["B1", "Existing"],
          ["C1", "Data"],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 0 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      expect(cell("B1")).toBe("");
      expect(cell("C1")).toBe("");
    });

    it("should adjust relative references in formulas", () => {
      setCellContent("A1", "=A2+B2");

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 2 }, row: { type: "number", value: 0 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      // Check raw formula content, not evaluated value
      const sheetContent = engine.getSheetSerialized(sheetAddress);
      expect(sheetContent.get("B1")).toBe("=B2+C2");
      expect(sheetContent.get("C1")).toBe("=C2+D2");
    });

    it("should preserve absolute references in formulas", () => {
      setCellContent("A1", "=$A$1+B2");

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      // Check raw formula content, not evaluated value
      const sheetContent = engine.getSheetSerialized(sheetAddress);
      expect(sheetContent.get("B1")).toBe("=$A$1+C2");
    });

    it("should handle mixed absolute references", () => {
      setCellContent("A1", "=A$1+$B2");

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 1 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "down");

      // Check raw formula content, not evaluated value
      const sheetContent = engine.getSheetSerialized(sheetAddress);
      expect(sheetContent.get("A2")).toBe("=A$1+$B3");
    });
  });

  describe("autoFill - Multi-Cell Linear Progression", () => {
    it("should infer linear step for numbers going down", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 2],
          ["A2", 4],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 4 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "down");

      expect(cell("A3")).toBe(6);
      expect(cell("A4")).toBe(8);
      expect(cell("A5")).toBe(10);
    });

    it("should infer linear step for numbers going right", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["B1", 3],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 2, row: 0 },
        end: { col: { type: "number", value: 4 }, row: { type: "number", value: 0 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      expect(cell("C1")).toBe(5);
      expect(cell("D1")).toBe(7);
      expect(cell("E1")).toBe(9);
    });

    it("should infer linear step for string numbers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "10"],
          ["A2", "15"],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 3 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "down");

      expect(cell("A3")).toBe("20");
      expect(cell("A4")).toBe("25");
    });

    it("should handle negative steps", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 7],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 3 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "down");

      expect(cell("A3")).toBe(4);
      expect(cell("A4")).toBe(1);
    });
  });

  describe("autoFill - Multi-Cell Pattern Repetition", () => {
    it("should repeat pattern when no linear step is found", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "A"],
          ["A2", "B"],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 5 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "down");

      expect(cell("A3")).toBe("A");
      expect(cell("A4")).toBe("B");
      expect(cell("A5")).toBe("A");
      expect(cell("A6")).toBe("B");
    });

    it("should repeat 2x2 block pattern", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "A1"],
          ["B1", "B1"],
          ["A2", "A2"],
          ["B2", "B2"],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 2, row: 0 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 1 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      expect(cell("C1")).toBe("A1");
      expect(cell("D1")).toBe("B1");
      expect(cell("C2")).toBe("A2");
      expect(cell("D2")).toBe("B2");
    });

    it("should adjust formulas in repeated patterns", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=A3"],
          ["A2", "=A4"],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "right");

      // Check raw formula content, not evaluated value
      const sheetContent = engine.getSheetSerialized(sheetAddress);
      expect(sheetContent.get("B1")).toBe("=B3");
      expect(sheetContent.get("B2")).toBe("=B4");
    });
  });

  describe("autoFill - Error Cases", () => {
    it("should throw error for infinite ranges", () => {
      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const infiniteFillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "infinity", sign: "positive" }, row: { type: "number", value: 0 } }
      };

      expect(() => engine.autoFill(sheetAddress, seedRange, infiniteFillRange, "right"))
        .toThrow("AutoFill with infinite ranges is not supported");
    });

    it("should throw error for non-existent sheet", () => {
      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 1, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } }
      };

      expect(() => engine.autoFill({ workbookName, sheetName: "NonExistentSheet" }, seedRange, fillRange, "right"))
        .toThrow('Sheet not found');
    });
  });

  describe("autoFill - Direction Handling", () => {
    it("should handle filling up direction", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A3", 10],
          ["A4", 15],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 0, row: 2 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 3 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 0 }, row: { type: "number", value: 1 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "up");

      expect(cell("A2")).toBe(5);
      expect(cell("A1")).toBe(0);
    });

    it("should handle filling left direction", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["C1", 20],
          ["D1", 25],
        ])
      );

      const seedRange: SpreadsheetRange = {
        start: { col: 2, row: 0 },
        end: { col: { type: "number", value: 3 }, row: { type: "number", value: 0 } }
      };

      const fillRange: SpreadsheetRange = {
        start: { col: 0, row: 0 },
        end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } }
      };

      engine.autoFill(sheetAddress, seedRange, fillRange, "left");

      expect(cell("B1")).toBe(15);
      expect(cell("A1")).toBe(10);
    });
  });
});
