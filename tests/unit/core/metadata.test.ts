import { describe, it, expect, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../src/core/engine";

// Example metadata interface for testing
interface TestMetadata {
  richText?: {
    content: string;
    formatting: Array<{ start: number; end: number; bold?: boolean }>;
  };
  link?: {
    url: string;
    title?: string;
  };
  comment?: string;
  customData?: Record<string, unknown>;
}

describe("Cell Metadata", () => {
  let engine: FormulaEngine<{ cell: TestMetadata }>;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty<{ cell: TestMetadata }>();
    engine.addWorkbook("wb1");
    engine.addSheet({ workbookName: "wb1", sheetName: "sheet1" });
  });

  describe("Basic Operations", () => {
    it("should set and get cell metadata", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const metadata: TestMetadata = {
        richText: {
          content: "Hello World",
          formatting: [{ start: 0, end: 5, bold: true }],
        },
      };

      engine.setCellMetadata(addr, metadata);
      const retrieved = engine.getCellMetadata(addr);

      expect(retrieved).toEqual(metadata);
    });

    it("should return undefined for cell without metadata", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const retrieved = engine.getCellMetadata(addr);

      expect(retrieved).toBeUndefined();
    });

    it("should update existing metadata", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellMetadata(addr, { comment: "First comment" });
      expect(engine.getCellMetadata(addr)).toEqual({
        comment: "First comment",
      });

      engine.setCellMetadata(addr, { comment: "Updated comment" });
      expect(engine.getCellMetadata(addr)).toEqual({
        comment: "Updated comment",
      });
    });

    it("should store complex nested metadata", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const metadata: TestMetadata = {
        richText: {
          content: "Complex text",
          formatting: [
            { start: 0, end: 7, bold: true },
            { start: 8, end: 12 },
          ],
        },
        link: {
          url: "https://example.com",
          title: "Example",
        },
        comment: "This is a comment",
        customData: {
          nested: {
            deeply: {
              value: 42,
            },
          },
        },
      };

      engine.setCellMetadata(addr, metadata);
      const retrieved = engine.getCellMetadata(addr);

      expect(retrieved).toEqual(metadata);
    });

    it("should handle metadata for multiple cells", () => {
      const addr1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const addr2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };
      const addr3 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 1,
      };

      engine.setCellMetadata(addr1, { comment: "Cell A1" });
      engine.setCellMetadata(addr2, { comment: "Cell B1" });
      engine.setCellMetadata(addr3, { link: { url: "https://example.com" } });

      expect(engine.getCellMetadata(addr1)).toEqual({ comment: "Cell A1" });
      expect(engine.getCellMetadata(addr2)).toEqual({ comment: "Cell B1" });
      expect(engine.getCellMetadata(addr3)).toEqual({
        link: { url: "https://example.com" },
      });
    });
  });

  describe("Paste Operations", () => {
    it("should copy metadata when pasting cells", () => {
      const source = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const target = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };

      engine.setCellContent(source, "Hello");
      engine.setCellMetadata(source, {
        richText: {
          content: "Hello",
          formatting: [{ start: 0, end: 5, bold: true }],
        },
      });

      engine.pasteCells([source], target, { include: "all" });

      expect(engine.getCellValue(target)).toBe("Hello");
      expect(engine.getCellMetadata(target)).toEqual({
        richText: {
          content: "Hello",
          formatting: [{ start: 0, end: 5, bold: true }],
        },
      });
    });

    it('should not copy metadata when target is "content"', () => {
      const source = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const target = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };

      engine.setCellContent(source, "Hello");
      engine.setCellMetadata(source, { comment: "Source comment" });

      engine.pasteCells([source], target, { include: ["content"] });

      expect(engine.getCellValue(target)).toBe("Hello");
      expect(engine.getCellMetadata(target)).toBeUndefined();
    });

    it("should copy metadata for multi-cell paste", () => {
      const source1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const source2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };

      engine.setCellContent(source1, "A");
      engine.setCellContent(source2, "B");
      engine.setCellMetadata(source1, { comment: "Comment A" });
      engine.setCellMetadata(source2, { comment: "Comment B" });

      const targetStart = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 1,
      };
      engine.pasteCells([source1, source2], targetStart, { include: "all" });

      const target1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 1,
      };
      const target2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 1,
      };

      expect(engine.getCellMetadata(target1)).toEqual({ comment: "Comment A" });
      expect(engine.getCellMetadata(target2)).toEqual({ comment: "Comment B" });
    });
  });

  describe("AutoFill Operations", () => {
    it("should copy metadata when autofilling", () => {
      const source = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(source, "1");
      engine.setCellMetadata(source, { comment: "Source cell" });

      engine.autoFill(
        { workbookName: "wb1", sheetName: "sheet1" },
        {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 0 },
            row: { type: "number", value: 0 },
          },
        },
        [
          {
            start: { col: 0, row: 1 },
            end: {
              col: { type: "number", value: 0 },
              row: { type: "number", value: 2 },
            },
          },
        ],
        "down"
      );

      const target1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 1,
      };
      const target2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 2,
      };

      expect(engine.getCellMetadata(target1)).toEqual({
        comment: "Source cell",
      });
      expect(engine.getCellMetadata(target2)).toEqual({
        comment: "Source cell",
      });
    });
  });

  describe("Clear Operations", () => {
    it("should clear metadata when clearing range", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(addr, "Hello");
      engine.setCellMetadata(addr, { comment: "Test comment" });

      engine.clearSpreadsheetRange({
        workbookName: "wb1",
        sheetName: "sheet1",
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 0 },
            row: { type: "number", value: 0 },
          },
        },
      });

      // After clearing, cell value should be empty or undefined
      const value = engine.getCellValue(addr);
      expect(value === undefined || value === "").toBe(true);
      expect(engine.getCellMetadata(addr)).toBeUndefined();
    });

    it("should clear metadata for all cells in range", () => {
      const addr1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const addr2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };
      const addr3 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 1,
      };

      engine.setCellContent(addr1, "A");
      engine.setCellContent(addr2, "B");
      engine.setCellContent(addr3, "C");
      engine.setCellMetadata(addr1, { comment: "A" });
      engine.setCellMetadata(addr2, { comment: "B" });
      engine.setCellMetadata(addr3, { comment: "C" });

      engine.clearSpreadsheetRange({
        workbookName: "wb1",
        sheetName: "sheet1",
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "number", value: 1 },
          },
        },
      });

      expect(engine.getCellMetadata(addr1)).toBeUndefined();
      expect(engine.getCellMetadata(addr2)).toBeUndefined();
      expect(engine.getCellMetadata(addr3)).toBeUndefined();
    });
  });

  describe("Serialization", () => {
    it("should serialize and deserialize metadata", () => {
      const addr1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const addr2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };

      const metadata1: TestMetadata = {
        richText: {
          content: "Rich Text",
          formatting: [{ start: 0, end: 4, bold: true }],
        },
        link: { url: "https://example.com" },
      };

      const metadata2: TestMetadata = {
        comment: "A comment",
      };

      engine.setCellContent(addr1, "Cell 1");
      engine.setCellContent(addr2, "Cell 2");
      engine.setCellMetadata(addr1, metadata1);
      engine.setCellMetadata(addr2, metadata2);

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty<{ cell: TestMetadata }>();
      newEngine.resetToSerializedEngine(serialized);

      expect(newEngine.getCellValue(addr1)).toBe("Cell 1");
      expect(newEngine.getCellValue(addr2)).toBe("Cell 2");
      expect(newEngine.getCellMetadata(addr1)).toEqual(metadata1);
      expect(newEngine.getCellMetadata(addr2)).toEqual(metadata2);
    });

    it("should handle empty metadata in serialization", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      engine.setCellContent(addr, "No metadata");

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty<{ cell: TestMetadata }>();
      newEngine.resetToSerializedEngine(serialized);

      expect(newEngine.getCellValue(addr)).toBe("No metadata");
      expect(newEngine.getCellMetadata(addr)).toBeUndefined();
    });
  });

  describe("Workbook Cloning", () => {
    it("should clone metadata when cloning workbook", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(addr, "Original");
      engine.setCellMetadata(addr, {
        richText: {
          content: "Original",
          formatting: [{ start: 0, end: 8, bold: true }],
        },
        comment: "Original comment",
      });

      engine.cloneWorkbook("wb1", "wb2");

      const clonedAddr = {
        workbookName: "wb2",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      expect(engine.getCellValue(clonedAddr)).toBe("Original");
      expect(engine.getCellMetadata(clonedAddr)).toEqual({
        richText: {
          content: "Original",
          formatting: [{ start: 0, end: 8, bold: true }],
        },
        comment: "Original comment",
      });
    });

    it("should clone all metadata in workbook", () => {
      const addr1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const addr2 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };

      engine.setCellMetadata(addr1, { comment: "First" });
      engine.setCellMetadata(addr2, { comment: "Second" });

      engine.cloneWorkbook("wb1", "wb2");

      const cloned1 = {
        workbookName: "wb2",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const cloned2 = {
        workbookName: "wb2",
        sheetName: "sheet1",
        colIndex: 1,
        rowIndex: 0,
      };

      expect(engine.getCellMetadata(cloned1)).toEqual({ comment: "First" });
      expect(engine.getCellMetadata(cloned2)).toEqual({ comment: "Second" });
    });

    it("should keep original and cloned metadata independent", () => {
      const originalAddr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      engine.setCellMetadata(originalAddr, { comment: "Original" });

      engine.cloneWorkbook("wb1", "wb2");

      const clonedAddr = {
        workbookName: "wb2",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      // Modify original
      engine.setCellMetadata(originalAddr, { comment: "Modified" });

      // Cloned should remain unchanged
      expect(engine.getCellMetadata(clonedAddr)).toEqual({
        comment: "Original",
      });
    });
  });

  describe("Multi-Sheet Operations", () => {
    beforeEach(() => {
      engine.addSheet({ workbookName: "wb1", sheetName: "sheet2" });
    });

    it("should handle metadata across multiple sheets", () => {
      const addr1 = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const addr2 = {
        workbookName: "wb1",
        sheetName: "sheet2",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellMetadata(addr1, { comment: "Sheet 1" });
      engine.setCellMetadata(addr2, { comment: "Sheet 2" });

      expect(engine.getCellMetadata(addr1)).toEqual({ comment: "Sheet 1" });
      expect(engine.getCellMetadata(addr2)).toEqual({ comment: "Sheet 2" });
    });

    it("should copy metadata between sheets", () => {
      const source = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      const target = {
        workbookName: "wb1",
        sheetName: "sheet2",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(source, "Data");
      engine.setCellMetadata(source, { link: { url: "https://example.com" } });

      engine.pasteCells([source], target, { include: "all" });

      expect(engine.getCellMetadata(target)).toEqual({
        link: { url: "https://example.com" },
      });
    });
  });

  describe("Edge Cases", () => {
    it("should handle null and undefined in metadata", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellMetadata(addr, {
        customData: {
          nullValue: null,
          undefinedValue: undefined,
        },
      });

      const retrieved = engine.getCellMetadata(addr);
      expect(retrieved?.customData?.nullValue).toBeNull();
      // Note: undefined values may be lost in JSON serialization
    });

    it("should handle very large metadata objects", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      const largeMetadata: TestMetadata = {
        customData: {
          array: Array(1000)
            .fill(0)
            .map((_, i) => ({ id: i, value: `Item ${i}` })),
        },
      };

      engine.setCellMetadata(addr, largeMetadata);
      const retrieved = engine.getCellMetadata(addr);

      expect(retrieved?.customData?.array).toHaveLength(1000);
    });

    it("should handle metadata for cells with formulas", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(addr, "=1+1");
      engine.setCellMetadata(addr, { comment: "This is a formula" });

      expect(engine.getCellValue(addr)).toBe(2);
      expect(engine.getCellMetadata(addr)).toEqual({
        comment: "This is a formula",
      });
    });
  });

  describe("Integration with Content", () => {
    it("should preserve metadata when updating cell content", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(addr, "Original");
      engine.setCellMetadata(addr, {
        richText: {
          content: "Original",
          formatting: [{ start: 0, end: 8, bold: true }],
        },
      });

      engine.setCellContent(addr, "Updated");

      // Metadata should remain
      expect(engine.getCellMetadata(addr)).toEqual({
        richText: {
          content: "Original",
          formatting: [{ start: 0, end: 8, bold: true }],
        },
      });
    });

    it("should allow independent content and metadata updates", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      engine.setCellContent(addr, "Content 1");
      engine.setCellMetadata(addr, { comment: "Metadata 1" });

      engine.setCellContent(addr, "Content 2");
      expect(engine.getCellMetadata(addr)).toEqual({ comment: "Metadata 1" });

      engine.setCellMetadata(addr, { comment: "Metadata 2" });
      expect(engine.getCellValue(addr)).toBe("Content 2");
      expect(engine.getCellMetadata(addr)).toEqual({ comment: "Metadata 2" });
    });
  });

  describe("Sheet Metadata", () => {
    let engine: FormulaEngine<{
      cell: TestMetadata;
      sheet: { textBoxes?: string[]; frozen?: boolean };
    }>;
    beforeEach(() => {
      engine = FormulaEngine.buildEmpty<{
        cell: TestMetadata;
        sheet: { textBoxes?: string[]; frozen?: boolean };
      }>();
      engine.addWorkbook("wb1");
      engine.addSheet({ workbookName: "wb1", sheetName: "sheet1" });
    });

    it("should set and get sheet metadata", () => {
      const sheetMeta = { textBoxes: ["box1", "box2"], frozen: true };

      engine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        sheetMeta
      );
      const retrieved = engine.getSheetMetadata({
        workbookName: "wb1",
        sheetName: "sheet1",
      });

      expect(retrieved).toEqual(sheetMeta);
    });

    it("should handle undefined sheet metadata", () => {
      const retrieved = engine.getSheetMetadata({
        workbookName: "wb1",
        sheetName: "sheet1",
      });
      expect(retrieved).toBeUndefined();
    });

    it("should preserve sheet metadata when renaming sheet", () => {
      engine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["box1"] }
      );

      engine.renameSheet({
        workbookName: "wb1",
        sheetName: "sheet1",
        newSheetName: "renamed",
      });

      const retrieved = engine.getSheetMetadata({
        workbookName: "wb1",
        sheetName: "renamed",
      });
      expect(retrieved).toEqual({ textBoxes: ["box1"] });
    });

    it("should clone sheet metadata when cloning workbook", () => {
      engine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        {
          textBoxes: ["box1", "box2"],
          frozen: true,
        }
      );

      engine.cloneWorkbook("wb1", "wb2");

      const clonedMeta = engine.getSheetMetadata({
        workbookName: "wb2",
        sheetName: "sheet1",
      });
      expect(clonedMeta).toEqual({ textBoxes: ["box1", "box2"], frozen: true });
    });

    it("should serialize and deserialize sheet metadata", () => {
      engine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        {
          textBoxes: ["text box 1"],
        }
      );

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty<{
        cell: TestMetadata;
        sheet: { textBoxes?: string[]; frozen?: boolean };
      }>();
      newEngine.resetToSerializedEngine(serialized);

      const retrieved = newEngine.getSheetMetadata({
        workbookName: "wb1",
        sheetName: "sheet1",
      });
      expect(retrieved).toEqual({ textBoxes: ["text box 1"] });
    });

    it("should handle multiple sheets with different metadata", () => {
      engine.addSheet({ workbookName: "wb1", sheetName: "sheet2" });

      engine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["box1"] }
      );
      engine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet2" },
        { frozen: true }
      );

      expect(
        engine.getSheetMetadata({ workbookName: "wb1", sheetName: "sheet1" })
      ).toEqual({ textBoxes: ["box1"] });
      expect(
        engine.getSheetMetadata({ workbookName: "wb1", sheetName: "sheet2" })
      ).toEqual({ frozen: true });
    });
  });

  describe("Workbook Metadata", () => {
    let engine: FormulaEngine<{
      cell: TestMetadata;
      workbook: { theme?: string; author?: string };
    }>;
    beforeEach(() => {
      engine = FormulaEngine.buildEmpty<{
        cell: TestMetadata;
        workbook: { theme?: string; author?: string };
      }>();
      engine.addWorkbook("wb1");
      engine.addSheet({ workbookName: "wb1", sheetName: "sheet1" });
    });

    it("should set and get workbook metadata", () => {
      const workbookMeta = { theme: "dark", author: "John Doe" };

      engine.setWorkbookMetadata("wb1", workbookMeta);
      const retrieved = engine.getWorkbookMetadata("wb1");

      expect(retrieved).toEqual(workbookMeta);
    });

    it("should handle undefined workbook metadata", () => {
      const retrieved = engine.getWorkbookMetadata("wb1");
      expect(retrieved).toBeUndefined();
    });

    it("should preserve workbook metadata when renaming workbook", () => {
      engine.setWorkbookMetadata("wb1", { theme: "light", author: "Jane" });

      engine.renameWorkbook({
        workbookName: "wb1",
        newWorkbookName: "renamed",
      });

      const retrieved = engine.getWorkbookMetadata("renamed");
      expect(retrieved).toEqual({ theme: "light", author: "Jane" });
    });

    it("should clone workbook metadata when cloning workbook", () => {
      engine.setWorkbookMetadata("wb1", { theme: "blue", author: "Bob" });

      engine.cloneWorkbook("wb1", "wb2");

      const clonedMeta = engine.getWorkbookMetadata("wb2");
      expect(clonedMeta).toEqual({ theme: "blue", author: "Bob" });
    });

    it("should serialize and deserialize workbook metadata", () => {
      engine.setWorkbookMetadata("wb1", { theme: "green" });

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty<{
        cell: TestMetadata;
        workbook: { theme?: string; author?: string };
      }>();
      newEngine.resetToSerializedEngine(serialized);

      const retrieved = newEngine.getWorkbookMetadata("wb1");
      expect(retrieved).toEqual({ theme: "green" });
    });
  });

  describe("Combined Metadata", () => {
    let combinedEngine: FormulaEngine<{
      cell: TestMetadata;
      sheet: { textBoxes?: string[] };
      workbook: { theme?: string };
    }>;

    beforeEach(() => {
      combinedEngine = FormulaEngine.buildEmpty<{
        cell: TestMetadata;
        sheet: { textBoxes?: string[] };
        workbook: { theme?: string };
      }>();
      combinedEngine.addWorkbook("wb1");
      combinedEngine.addSheet({ workbookName: "wb1", sheetName: "sheet1" });
    });

    it("should handle cell, sheet, and workbook metadata together", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      combinedEngine.setCellContent(addr, "Test");
      combinedEngine.setCellMetadata(addr, { comment: "formatted" });
      combinedEngine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["box1"] }
      );
      combinedEngine.setWorkbookMetadata("wb1", { theme: "dark" });

      expect(combinedEngine.getCellMetadata(addr)).toEqual({
        comment: "formatted",
      });
      expect(
        combinedEngine.getSheetMetadata({
          workbookName: "wb1",
          sheetName: "sheet1",
        })
      ).toEqual({ textBoxes: ["box1"] });
      expect(combinedEngine.getWorkbookMetadata("wb1")).toEqual({
        theme: "dark",
      });
    });

    it("should clone all three metadata types correctly", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      combinedEngine.setCellContent(addr, "Test");
      combinedEngine.setCellMetadata(addr, { comment: "formatted" });
      combinedEngine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["box1", "box2"] }
      );
      combinedEngine.setWorkbookMetadata("wb1", { theme: "light" });

      combinedEngine.cloneWorkbook("wb1", "wb2");

      const clonedAddr = {
        workbookName: "wb2",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      expect(combinedEngine.getCellMetadata(clonedAddr)).toEqual({
        comment: "formatted",
      });
      expect(
        combinedEngine.getSheetMetadata({
          workbookName: "wb2",
          sheetName: "sheet1",
        })
      ).toEqual({ textBoxes: ["box1", "box2"] });
      expect(combinedEngine.getWorkbookMetadata("wb2")).toEqual({
        theme: "light",
      });
    });

    it("should serialize all three metadata types", () => {
      const addr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };

      combinedEngine.setCellContent(addr, "Data");
      combinedEngine.setCellMetadata(addr, { comment: "bold" });
      combinedEngine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["tb1"] }
      );
      combinedEngine.setWorkbookMetadata("wb1", { theme: "ocean" });

      const serialized = combinedEngine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty<{
        cell: TestMetadata;
        sheet: { textBoxes?: string[] };
        workbook: { theme?: string };
      }>();
      newEngine.resetToSerializedEngine(serialized);

      const restoredAddr = {
        workbookName: "wb1",
        sheetName: "sheet1",
        colIndex: 0,
        rowIndex: 0,
      };
      expect(newEngine.getCellValue(restoredAddr)).toBe("Data");
      expect(newEngine.getCellMetadata(restoredAddr)).toEqual({
        comment: "bold",
      });
      expect(
        newEngine.getSheetMetadata({ workbookName: "wb1", sheetName: "sheet1" })
      ).toEqual({ textBoxes: ["tb1"] });
      expect(newEngine.getWorkbookMetadata("wb1")).toEqual({ theme: "ocean" });
    });

    it("should keep cloned metadata independent from original", () => {
      combinedEngine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["box1"] }
      );
      combinedEngine.setWorkbookMetadata("wb1", { theme: "original" });

      combinedEngine.cloneWorkbook("wb1", "wb2");

      // Modify original
      combinedEngine.setSheetMetadata(
        { workbookName: "wb1", sheetName: "sheet1" },
        { textBoxes: ["modified"] }
      );
      combinedEngine.setWorkbookMetadata("wb1", { theme: "changed" });

      // Cloned should remain unchanged
      expect(
        combinedEngine.getSheetMetadata({
          workbookName: "wb2",
          sheetName: "sheet1",
        })
      ).toEqual({ textBoxes: ["box1"] });
      expect(combinedEngine.getWorkbookMetadata("wb2")).toEqual({
        theme: "original",
      });
    });
  });
});
