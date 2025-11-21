import { describe, expect, test, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../../src/core/engine";
import type { CellAddress } from "../../../../src/core/types";

describe("CopyManager", () => {
  let engine: FormulaEngine;
  const workbookName = "TestWorkbook";
  const sheetName = "Sheet1";

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("copyCells - basic functionality", () => {
    test("copies literal values", () => {
      // Set up source cells
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Hello"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        42
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Check target cells
      const targetA3 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      });
      const targetB3 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 2,
      });

      expect(targetA3).toBe("Hello");
      expect(targetB3).toBe(42);

      // Source cells should still exist
      const sourceA1 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      expect(sourceA1).toBe("Hello");
    });

    test("copies formulas and adjusts references", () => {
      // Set up source cell with formula
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=B1+C1"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Check that formula was adjusted (A1 -> A3, so B1+C1 -> B3+C3)
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("A3")).toBe("=B3+C3");
    });

    test("preserves absolute references in formulas", () => {
      // Set up source cell with absolute references
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=$B$1+C1"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Check that absolute reference is preserved, relative is adjusted
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("A3")).toBe("=$B$1+C3");
    });

    test("copies formulas as values when type is 'value'", () => {
      // Set up source cells
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        10
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        20
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 },
        "=A1+B1"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 2,
        rowIndex: 2,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "value",
        formatting: false,
      });

      // Check that target has the evaluated value, not the formula
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("C3")).toBe(30); // Should be the value, not "=A1+B1"
    });
  });

  describe("copyCells - cut behavior", () => {
    test("clears source cells when cut is true", () => {
      // Set up source cells
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Hello"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        42
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      engine.copyCells(source, target, {
        cut: true,
        type: "formula",
        formatting: false,
      });

      // Check target cells
      const targetA3 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      });
      const targetB3 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 2,
      });

      expect(targetA3).toBe("Hello");
      expect(targetB3).toBe(42);

      // Source cells should be cleared
      const sourceA1 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      const sourceB1 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 0,
      });

      expect(sourceA1).toBe("");
      expect(sourceB1).toBe("");
    });
  });

  describe("copyCells - formatting", () => {
    test("copies cell styles when formatting is true", () => {
      // Add a cell style to source range
      engine.addCellStyle({
        area: {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } },
          },
        },
        style: {
          backgroundColor: "#FF0000",
          color: "#FFFFFF",
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: true,
      });

      // Check that styling was copied to target range
      const targetStyle = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      });

      expect(targetStyle).toBeDefined();
      expect(targetStyle?.backgroundColor).toBe("#FF0000");
      expect(targetStyle?.color).toBe("#FFFFFF");
    });

    test("copies conditional styles when formatting is true", () => {
      // Add a conditional style to source range
      engine.addConditionalStyle({
        area: {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 1 } },
          },
        },
        condition: {
          type: "formula",
          formula: "ROW() > 5",
          color: { l: 50, c: 80, h: 0 },
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 10,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: true,
      });

      // Check that conditional style was copied
      const styles = engine.getConditionalStyles(workbookName);
      expect(styles.length).toBeGreaterThan(1); // Should have original + copied style

      // Find the copied style (should be at row 10-11)
      const copiedStyle = styles.find(
        (s) =>
          s.area.range.start.row === 10 &&
          s.area.range.end.row.type === "number" &&
          s.area.range.end.row.value === 11
      );
      expect(copiedStyle).toBeDefined();
      expect(copiedStyle?.condition.type).toBe("formula");
    });

    test("does not copy formatting when formatting is false", () => {
      // Add a cell style to source range
      engine.addCellStyle({
        area: {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } },
          },
        },
        style: {
          backgroundColor: "#FF0000",
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      };

      const stylesBefore = engine.getCellStyles(workbookName).length;

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Formatting should not be copied
      const stylesAfter = engine.getCellStyles(workbookName).length;
      expect(stylesAfter).toBe(stylesBefore); // No new styles added
    });
  });

  describe("copyCells - non-contiguous cells", () => {
    test("handles non-contiguous source cells", () => {
      // Set up non-contiguous source cells
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "A"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 2, rowIndex: 2 },
        "B"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { workbookName, sheetName, colIndex: 2, rowIndex: 2 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Top-left is A1 (0,0), so offsets are maintained
      // A1 (0,0) -> F6 (5,5) - offset is (5,5)
      // C3 (2,2) -> H8 (7,7) - offset is (5,5)
      const targetF6 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      const targetH8 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 7,
        rowIndex: 7,
      });

      expect(targetF6).toBe("A");
      expect(targetH8).toBe("B");
    });
  });

  describe("copyCells - edge cases", () => {
    test("handles empty source cells", () => {
      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      // Should not throw
      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Target should remain empty
      const targetA3 = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      });
      expect(targetA3).toBe("");
    });

    test("handles empty source array", () => {
      const source: CellAddress[] = [];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      // Should not throw
      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });
    });

    test("copies error values as values when type is 'value'", () => {
      // Set up source cell with error formula
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=1/0"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "value",
        formatting: false,
      });

      // Should copy the formula as-is since evaluation resulted in error
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("A3")).toBe("=1/0");
    });
  });

  describe("copyCells - cross-sheet", () => {
    test("copies cells to different sheet", () => {
      const targetSheetName = "Sheet2";
      engine.addSheet({ workbookName, sheetName: targetSheetName });

      // Set up source cell
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Hello"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName: targetSheetName,
        colIndex: 0,
        rowIndex: 0,
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: false,
      });

      // Check target cell in different sheet
      const targetValue = engine.getCellValue({
        workbookName,
        sheetName: targetSheetName,
        colIndex: 0,
        rowIndex: 0,
      });

      expect(targetValue).toBe("Hello");
    });
  });

  describe("copyCells - combined operations", () => {
    test("cut with formatting copies styles and clears source", () => {
      // Set up source cell with content and style
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Test"
      );

      engine.addCellStyle({
        area: {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } },
          },
        },
        style: {
          backgroundColor: "#00FF00",
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      };

      engine.copyCells(source, target, {
        cut: true,
        type: "formula",
        formatting: true,
      });

      // Target should have content and style
      const targetValue = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      });
      expect(targetValue).toBe("Test");

      const targetStyle = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      });
      expect(targetStyle?.backgroundColor).toBe("#00FF00");

      // Source should be cleared
      const sourceValue = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      expect(sourceValue).toBe("");
    });

    test("copies only selected cell styling, not entire styled range", () => {
      // Style the entire A column (A:A or A1:A with infinite rows)
      engine.addCellStyle({
        area: {
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { 
              col: { type: "number", value: 0 }, 
              row: { type: "infinity", sign: "positive" } 
            },
          },
        },
        style: {
          backgroundColor: "#FF0000",
          fontSize: 16,
        },
      });

      // Copy only cell A5 to B2
      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 4 }, // A5
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1, // B
        rowIndex: 1, // 2
      };

      engine.copyCells(source, target, {
        cut: false,
        type: "formula",
        formatting: true,
      });

      // Only B2 should be styled, not the entire B column
      const b2Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 1,
      });
      expect(b2Style).toBeDefined();
      expect(b2Style?.backgroundColor).toBe("#FF0000");
      expect(b2Style?.fontSize).toBe(16);

      // B1 should NOT be styled
      const b1Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 0,
      });
      expect(b1Style).toBeUndefined();

      // B10 should NOT be styled
      const b10Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 9,
      });
      expect(b10Style).toBeUndefined();

      // Verify we didn't create a B:B style - check the cellStyles array
      const cellStyles = engine.getCellStyles(workbookName);
      // Should have original A:A + new B2:B2 (single cell)
      expect(cellStyles).toHaveLength(2);
      
      const b2StyleRule = cellStyles.find(s => 
        s.area.range.start.col === 1 && 
        s.area.range.start.row === 1
      );
      expect(b2StyleRule).toBeDefined();
      expect(b2StyleRule?.area.range.end.col).toEqual({ type: "number", value: 1 });
      expect(b2StyleRule?.area.range.end.row).toEqual({ type: "number", value: 1 });
    });
  });
});

