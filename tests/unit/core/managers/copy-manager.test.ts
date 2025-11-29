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

  describe("pasteCells - basic functionality", () => {
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "value",
        include: ["content"],
      });

      // Check that target has the evaluated value, not the formula
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("C3")).toBe(30); // Should be the value, not "=A1+B1"
    });
  });

  describe("pasteCells - cut behavior", () => {
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

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
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

  describe("pasteCells - formatting", () => {
    test("copies cell styles when target is 'all'", () => {
      // Add a cell style to source range
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 1 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: "all",
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

    test("copies conditional styles when target is 'all'", () => {
      // Add a conditional style to source range
      engine.addConditionalStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 1 },
                row: { type: "number", value: 1 },
              },
            },
          },
        ],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: "all",
      });

      // Check that conditional style was copied
      const styles = engine.getConditionalStylesIntersectingWithRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "infinity", sign: "positive" },
            row: { type: "infinity", sign: "positive" },
          },
        },
      });
      expect(styles.length).toBeGreaterThan(1); // Should have original + copied style

      // Find the copied style (should be at row 10-11)
      const copiedStyle = styles.find(
        (s) =>
          s.areas[0]!.range.start.row === 10 &&
          s.areas[0]!.range.end.row.type === "number" &&
          s.areas[0]!.range.end.row.value === 11
      );
      expect(copiedStyle).toBeDefined();
      expect(copiedStyle?.condition.type).toBe("formula");
    });

    test("does not copy formatting when target is 'content'", () => {
      // Add a cell style to source range
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 1 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
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

      const stylesBefore = engine.getStylesIntersectingWithRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "number", value: 0 },
          },
        },
      }).length;

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      // Formatting should not be copied
      const stylesAfter = engine.getStylesIntersectingWithRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 1 },
            row: { type: "number", value: 0 },
          },
        },
      }).length;
      expect(stylesAfter).toBe(stylesBefore); // No new styles added
    });
  });

  describe("pasteCells - non-contiguous cells", () => {
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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

  describe("pasteCells - edge cases", () => {
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
      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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
      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "value",
        include: ["content"],
      });

      // Should copy the formula as-is since evaluation resulted in error
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("A3")).toBe("=1/0");
    });
  });

  describe("pasteCells - cross-sheet", () => {
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
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

  describe("pasteCells - combined operations", () => {
    test("cut with all target pastes styles and clears source", () => {
      // Set up source cell with content and style
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Test"
      );

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
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

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: "all",
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
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "infinity", sign: "positive" },
              },
            },
          },
        ],
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

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: "all",
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
      const cellStyles = engine.getStylesIntersectingWithRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "infinity", sign: "positive" },
            row: { type: "infinity", sign: "positive" },
          },
        },
      });
      // Should have original A:A + new B2:B2 (single cell)
      expect(cellStyles).toHaveLength(2);

      const b2StyleRule = cellStyles.find(
        (s) =>
          s.areas[0]!.range.start.col === 1 && s.areas[0]!.range.start.row === 1
      );
      expect(b2StyleRule).toBeDefined();
      expect(b2StyleRule?.areas[0]!.range.end.col).toEqual({
        type: "number",
        value: 1,
      });
      expect(b2StyleRule?.areas[0]!.range.end.row).toEqual({
        type: "number",
        value: 1,
      });
    });
  });

  describe("pasteCells - style target mode", () => {
    test("copies only styles when target is 'style'", () => {
      // Set up source cell with content and style
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Content"
      );

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        style: {
          backgroundColor: "#0000FF",
          color: "#FFFFFF",
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 2,
        rowIndex: 2,
      };

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["style"],
      });

      // Target should have the style
      const targetStyle = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 2,
        rowIndex: 2,
      });
      expect(targetStyle).toBeDefined();
      expect(targetStyle?.backgroundColor).toBe("#0000FF");
      expect(targetStyle?.color).toBe("#FFFFFF");

      // Target should NOT have the content
      const targetValue = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 2,
        rowIndex: 2,
      });
      expect(targetValue).toBe(""); // Should be empty
    });

    test("ignores type property when target is 'style'", () => {
      // Set up source cell with formula
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=10+20"
      );

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        style: {
          bold: true,
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      };

      // Use type: "value" but target: "style" - type should be ignored
      engine.pasteCells(source, target, {
        cut: false,
        type: "value",
        include: ["style"],
      });

      // Target should have the style
      const targetStyle = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(targetStyle?.bold).toBe(true);

      // Target should NOT have any content (neither formula nor value)
      const targetValue = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(targetValue).toBe("");
    });

    test("copies conditional styles when target is 'style'", () => {
      // Set up source cell with content and conditional style
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Data"
      );

      engine.addConditionalStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 60, c: 70, h: 120 },
        },
      });

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];
      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 3,
        rowIndex: 3,
      };

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["style"],
      });

      // Check that conditional style was copied
      const styles = engine.getConditionalStylesIntersectingWithRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 3, row: 3 },
          end: {
            col: { type: "number", value: 3 },
            row: { type: "number", value: 3 },
          },
        },
      });
      expect(styles.length).toBeGreaterThan(0);

      // Target should NOT have the content
      const targetValue = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 3,
        rowIndex: 3,
      });
      expect(targetValue).toBe("");
    });
  });

  describe("fillAreas", () => {
    test("fills single area with literal value", () => {
      // Set up template cell
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Hello"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 7 },
              row: { type: "number" as const, value: 7 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: "all",
      });

      // Check that all cells in the area have the value
      for (let row = 5; row <= 7; row++) {
        for (let col = 5; col <= 7; col++) {
          const value = engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: col,
            rowIndex: row,
          });
          expect(value).toBe("Hello");
        }
      }
    });

    test("fills multiple areas with literal value", () => {
      // Set up template cell
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        42
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 6 },
              row: { type: "number" as const, value: 6 },
            },
          },
        },
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 10, row: 10 },
            end: {
              col: { type: "number" as const, value: 11 },
              row: { type: "number" as const, value: 11 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: "all",
      });

      // Check first area
      for (let row = 5; row <= 6; row++) {
        for (let col = 5; col <= 6; col++) {
          const value = engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: col,
            rowIndex: row,
          });
          expect(value).toBe(42);
        }
      }

      // Check second area
      for (let row = 10; row <= 11; row++) {
        for (let col = 10; col <= 11; col++) {
          const value = engine.getCellValue({
            workbookName,
            sheetName,
            colIndex: col,
            rowIndex: row,
          });
          expect(value).toBe(42);
        }
      }
    });

    test("fills with formulas and adjusts relative references", () => {
      // Set up template cell with formula at A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=B1+C1"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 6 },
              row: { type: "number" as const, value: 6 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });

      // F6 (5,5) should be =G6+H6 (offset +5,+5 from A1)
      expect(sheetContent.get("F6")).toBe("=G6+H6");
      // G6 (6,5) should be =H6+I6 (offset +6,+5 from A1)
      expect(sheetContent.get("G6")).toBe("=H6+I6");
      // F7 (5,6) should be =G7+H7 (offset +5,+6 from A1)
      expect(sheetContent.get("F7")).toBe("=G7+H7");
      // G7 (6,6) should be =H7+I7 (offset +6,+6 from A1)
      expect(sheetContent.get("G7")).toBe("=H7+I7");
    });

    test("fills with formulas and preserves absolute references", () => {
      // Set up template cell with absolute formula
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=$B$1"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 6 },
              row: { type: "number" as const, value: 6 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });

      // All cells should have the same absolute reference
      expect(sheetContent.get("F6")).toBe("=$B$1");
      expect(sheetContent.get("G6")).toBe("=$B$1");
      expect(sheetContent.get("F7")).toBe("=$B$1");
      expect(sheetContent.get("G7")).toBe("=$B$1");
    });

    test("fills with mixed references", () => {
      // Set up template cell with mixed references
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=$A1+B$1"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 5 },
              row: { type: "number" as const, value: 6 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });

      // F6 (5,5): =$A6+G$1 (col A stays, row adjusts; row 1 stays, col adjusts)
      expect(sheetContent.get("F6")).toBe("=$A6+G$1");
      // F7 (5,6): =$A7+G$1
      expect(sheetContent.get("F7")).toBe("=$A7+G$1");
    });

    test("fills only styles when target is 'style'", () => {
      // Set up template cell with content and style
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Template"
      );

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        style: {
          backgroundColor: "#FF0000",
        },
      });

      // Add existing content to target cells
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 5, rowIndex: 5 },
        "Existing"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 6 },
              row: { type: "number" as const, value: 6 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["style"],
      });

      // Check that styles were copied
      const f6Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Style?.backgroundColor).toBe("#FF0000");

      // Check that content was NOT changed
      const f6Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Value).toBe("Existing");

      // Check that F7 has style but no content
      const f7Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 6,
      });
      expect(f7Style?.backgroundColor).toBe("#FF0000");

      const f7Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 6,
      });
      expect(f7Value).toBe("");
    });

    test("fills only content when target is 'content'", () => {
      // Set up template cell with content and style
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Content"
      );

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        style: {
          backgroundColor: "#00FF00",
        },
      });

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 5 },
              row: { type: "number" as const, value: 5 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      // Check that content was copied
      const f6Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Value).toBe("Content");

      // Check that style was NOT copied
      const f6Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Style).toBeUndefined();
    });

    test("clears seed range when cut is true", () => {
      // Set up template cell
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "Template"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 5 },
              row: { type: "number" as const, value: 5 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: true,
        type: "formula",
        include: "all",
      });

      // Check that target was filled
      const f6Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Value).toBe("Template");

      // Check that seed range was cleared
      const a1Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      expect(a1Value).toBe("");
    });

    test("fills empty seed range clears target areas", () => {
      // Set up target cells with existing content
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 5, rowIndex: 5 },
        "Old"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 6, rowIndex: 5 },
        "Data"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      }; // Empty cell
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 6 },
              row: { type: "number" as const, value: 5 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      // Check that target cells were cleared
      const f6Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Value).toBe("");

      const g6Value = engine.getCellValue({
        workbookName,
        sheetName,
        colIndex: 6,
        rowIndex: 5,
      });
      expect(g6Value).toBe("");
    });

    test("fills with type='value' to copy evaluated values", () => {
      // Set up source cells for formula
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        10
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 },
        20
      );

      // Template cell with formula
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=B1+C1"
      );

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 0 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 5 },
              row: { type: "number" as const, value: 5 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "value",
        include: ["content"],
      });

      // Check that the evaluated value was copied, not the formula
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("F6")).toBe(30); // Should be the value, not "=B1+C1"
    });

    test("fills with 2x2 seed pattern (column-first strategy)", () => {
      // Set up 2x2 seed pattern
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        10
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        20
      ); // B1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        11
      ); // A2
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
        21
      ); // B2

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 1 },
            row: { type: "number" as const, value: 1 },
          },
        },
      };

      // Fill into 5x5 target range
      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 9 },
              row: { type: "number" as const, value: 9 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      // Check the pattern fills correctly
      // Column-first: first fill down to height 5, then replicate right
      // F6,F7,F8,F9,F10 should be 10,11,10,11,10 (pattern repeats vertically)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 5,
        })
      ).toBe(10); // F6
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 6,
        })
      ).toBe(11); // F7
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 7,
        })
      ).toBe(10); // F8
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 8,
        })
      ).toBe(11); // F9
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 9,
        })
      ).toBe(10); // F10

      // G column should be 20,21,20,21,20
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 6,
          rowIndex: 5,
        })
      ).toBe(20); // G6
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 6,
          rowIndex: 6,
        })
      ).toBe(21); // G7

      // H column should replicate F column (10,11,10,11,10)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 7,
          rowIndex: 5,
        })
      ).toBe(10); // H6
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 7,
          rowIndex: 6,
        })
      ).toBe(11); // H7

      // I column should replicate G column (20,21,20,21,20)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 8,
          rowIndex: 5,
        })
      ).toBe(20); // I6
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 8,
          rowIndex: 6,
        })
      ).toBe(21); // I7

      // J column should replicate F column again (partial - 1 col)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 9,
          rowIndex: 5,
        })
      ).toBe(10); // J6
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 9,
          rowIndex: 6,
        })
      ).toBe(11); // J7
    });

    test("fills with 2x2 seed containing formulas", () => {
      // Set up 2x2 seed with formulas
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=ROW()"
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=COLUMN()"
      ); // B1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        "=ROW()+10"
      ); // A2
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
        "=COLUMN()+10"
      ); // B2

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 1 },
            row: { type: "number" as const, value: 1 },
          },
        },
      };

      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 7 },
              row: { type: "number" as const, value: 7 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      // Check formulas are adjusted correctly
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });

      // F6 should have =ROW() (from A1, adjusted +5 rows, +5 cols)
      expect(sheetContent.get("F6")).toBe("=ROW()");
      // G6 should have =COLUMN() (from B1)
      expect(sheetContent.get("G6")).toBe("=COLUMN()");
      // F7 should have =ROW()+10 (from A2)
      expect(sheetContent.get("F7")).toBe("=ROW()+10");
      // G7 should have =COLUMN()+10 (from B2)
      expect(sheetContent.get("G7")).toBe("=COLUMN()+10");

      // Check pattern repeats
      // F8 should be =ROW() again (pattern repeats)
      expect(sheetContent.get("F8")).toBe("=ROW()");

      // H6 should replicate F6
      expect(sheetContent.get("H6")).toBe("=ROW()");
    });

    test("fills with multi-cell seed and styles", () => {
      // Set up 2x1 seed with different styles
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "A"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "B"
      );

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 0, row: 0 },
              end: {
                col: { type: "number", value: 0 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        style: { backgroundColor: "#FF0000" },
      });

      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 1, row: 0 },
              end: {
                col: { type: "number", value: 1 },
                row: { type: "number", value: 0 },
              },
            },
          },
        ],
        style: { backgroundColor: "#00FF00" },
      });

      const seedRange = {
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number" as const, value: 1 },
            row: { type: "number" as const, value: 0 },
          },
        },
      };

      const targetRanges = [
        {
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: {
              col: { type: "number" as const, value: 8 },
              row: { type: "number" as const, value: 7 },
            },
          },
        },
      ];

      engine.fillAreas(seedRange, targetRanges, {
        cut: false,
        type: "formula",
        include: "all",
      });

      // Check content
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 5,
          rowIndex: 5,
        })
      ).toBe("A");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 6,
          rowIndex: 5,
        })
      ).toBe("B");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 7,
          rowIndex: 5,
        })
      ).toBe("A"); // Replicates
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 8,
          rowIndex: 5,
        })
      ).toBe("B");

      // Check styles
      const f6Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      });
      expect(f6Style?.backgroundColor).toBe("#FF0000");

      const g6Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 6,
        rowIndex: 5,
      });
      expect(g6Style?.backgroundColor).toBe("#00FF00");

      const h6Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 7,
        rowIndex: 5,
      });
      expect(h6Style?.backgroundColor).toBe("#FF0000"); // Replicates F6's style

      const i6Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 8,
        rowIndex: 5,
      });
      expect(i6Style?.backgroundColor).toBe("#00FF00"); // Replicates G6's style
    });
  });

  describe("pasteCells - overlapping ranges", () => {
    test("should handle overlapping copy (A1:D5 to B1:E5)", () => {
      // Set up source data
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        10
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        20
      ); // B1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 },
        30
      ); // C1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 3, rowIndex: 0 },
        40
      ); // D1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        50
      ); // A2

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 }, // B1
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 }, // C1
        { workbookName, sheetName, colIndex: 3, rowIndex: 0 }, // D1
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 }, // A2
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1, // B1
        rowIndex: 0,
      };

      engine.pasteCells(source, target, {
        cut: false,
        type: "formula",
        include: ["content"],
      });

      // Should snapshot before copying, so B1 should become 10 (original A1 value)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(10);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 0,
        })
      ).toBe(20);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 3,
          rowIndex: 0,
        })
      ).toBe(30);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 4,
          rowIndex: 0,
        })
      ).toBe(40);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 1,
        })
      ).toBe(50);
    });

    test("should handle overlapping cut (A1:B4 to B1:C4) and clear source cells", () => {
      // Set up: A1 has a formula, A2-A4 and B1-B4 have values
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "=SUM(A1:A4)"
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        10
      ); // A2
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 2 },
        20
      ); // A3
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 3 },
        30
      ); // A4
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        100
      ); // B1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 },
        200
      ); // B2
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 2 },
        300
      ); // B3
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 3 },
        400
      ); // B4

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 }, // A2
        { workbookName, sheetName, colIndex: 0, rowIndex: 2 }, // A3
        { workbookName, sheetName, colIndex: 0, rowIndex: 3 }, // A4
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 }, // B1
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 }, // B2
        { workbookName, sheetName, colIndex: 1, rowIndex: 2 }, // B3
        { workbookName, sheetName, colIndex: 1, rowIndex: 3 }, // B4
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1, // B1
        rowIndex: 0,
      };

      // Cut A1:B4 and paste to B1:C4
      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // A1-A4 should be cleared (they were in source but not in target area)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe("");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 1,
        })
      ).toBe("");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 2,
        })
      ).toBe("");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 3,
        })
      ).toBe("");

      // B1-B4 should have the values from A1-A4 (snapshot)
      const sheetContent = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent.get("B1")).toBe("=SUM(B1:B4)"); // Formula adjusted
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 1,
        })
      ).toBe(10);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 2,
        })
      ).toBe(20);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 3,
        })
      ).toBe(30);

      // C1-C4 should have the values from B1-B4 (snapshot)
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 0,
        })
      ).toBe(100);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 1,
        })
      ).toBe(200);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 2,
        })
      ).toBe(300);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 3,
        })
      ).toBe(400);
    });

    test("should handle overlapping cut (A1:B2 to B1:C1) using pasteCells", () => {
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=SUM(A1:A2)"
      ); // B1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        1
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        2
      ); // A2

      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(1); // A1
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 1,
        })
      ).toBe(2); // A2
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(3); // B1
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 0,
        })
      ).toBe(""); // C1

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 }, // A2
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 }, // B1
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 }, // B2
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1, // B1
        rowIndex: 0,
      };

      // Cut A1:B2 and paste to B1:C1
      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(""); // A1
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 1,
        })
      ).toBe(""); // A2

      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(1); // B1
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 1,
        })
      ).toBe(2); // B2
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 0,
        })
      ).toBe(3); // C1
    });

    test("should handle overlapping cut (A1:B2 to B1) using smartPaste", () => {
      // Set up exact scenario: A1=1, A2=2, B1==SUM(A1:A2)
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        1
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        2
      ); // A2
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=SUM(A1:A2)"
      ); // B1

      // Verify initial state
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 0, rowIndex: 0 })).toBe(1);
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 0, rowIndex: 1 })).toBe(2);
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })).toBe(3);

      // Cut A1:B2 using smartPaste with single cell selection at B1
      const sourceCells = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 }, // B1
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 }, // A2
        { workbookName, sheetName, colIndex: 1, rowIndex: 1 }, // B2
      ];

      engine.smartPaste(
        sourceCells,
        {
          workbookName,
          sheetName,
          areas: [
            {
              start: { col: 1, row: 0 },
              end: {
                col: { type: "number", value: 1 },
                row: { type: "number", value: 0 },
              },
            },
          ],
        },
        {
          cut: true,
          type: "formula",
          include: "all",
        }
      );

      // A1 and A2 should be cleared
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 0, rowIndex: 0 })).toBe("");
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 0, rowIndex: 1 })).toBe("");

      // B1 should have 1 (from A1)
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 1, rowIndex: 0 })).toBe(1);

      // B2 should have 2 (from A2)
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 1, rowIndex: 1 })).toBe(2);

      // C1 should have the formula =SUM(B1:B2) (from B1, adjusted)
      const sheetContent = engine.getSheetSerialized({ workbookName, sheetName });
      expect(sheetContent.get("C1")).toBe("=SUM(B1:B2)");
      expect(engine.getCellValue({ workbookName, sheetName, colIndex: 2, rowIndex: 0 })).toBe(3);
    });

    test("should punch holes in style areas when cutting cells", () => {
      // Add a style covering J3:N20
      engine.addCellStyle({
        areas: [
          {
            workbookName,
            sheetName,
            range: {
              start: { col: 9, row: 2 }, // J3
              end: {
                col: { type: "number", value: 13 }, // N20
                row: { type: "number", value: 19 },
              },
            },
          },
        ],
        style: {
          backgroundColor: "#FF0000",
        },
      });

      // Verify style applies
      const j3Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 9,
        rowIndex: 2,
      });
      expect(j3Style?.backgroundColor).toBe("#FF0000");

      // Cut L11 (colIndex: 11, rowIndex: 10)
      engine.moveCell(
        { workbookName, sheetName, colIndex: 11, rowIndex: 10 },
        { workbookName, sheetName, colIndex: 20, rowIndex: 20 } // Move to U21
      );

      // L11 should now have no style (moved away)
      const l11Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 11,
        rowIndex: 10,
      });
      expect(l11Style).toBeUndefined();

      // J3 should still have style (not affected)
      const j3StyleAfter = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 9,
        rowIndex: 2,
      });
      expect(j3StyleAfter?.backgroundColor).toBe("#FF0000");

      // K11 should still have style (left of cut cell)
      const k11Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 10,
        rowIndex: 10,
      });
      expect(k11Style?.backgroundColor).toBe("#FF0000");

      // M11 should still have style (right of cut cell)
      const m11Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 12,
        rowIndex: 10,
      });
      expect(m11Style?.backgroundColor).toBe("#FF0000");

      // U21 should have the style (moved cell)
      const u21Style = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 20,
        rowIndex: 20,
      });
      expect(u21Style?.backgroundColor).toBe("#FF0000");

      // Verify the style has multiple areas now (hole punched)
      const allStyles = engine.getAllCellStyles();
      const redStyle = allStyles.find(
        (s) => s.style.backgroundColor === "#FF0000"
      );
      expect(redStyle).toBeDefined();
      expect(redStyle!.areas.length).toBeGreaterThan(1); // Should have multiple areas after hole punching
    });

    test("should handle overlapping cut with reference update", () => {
      // Set up: A1=10, E6=A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        10
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 4, rowIndex: 5 },
        "=A1"
      ); // E6

      // Verify initial state
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 4,
          rowIndex: 5,
        })
      ).toBe(10);

      // Cut A1 to B1 (overlapping scenario where we're moving within same row)
      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1, // B1
        rowIndex: 0,
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // B1 should now have 10
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(10);

      // A1 should be empty
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe("");

      // E6 should now reference B1 instead of A1
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 4,
          rowIndex: 5,
        })
      ).toBe(10);
      const sheetContent1 = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent1.get("E6")).toBe("=B1");
    });
  });

  describe("pasteCells - cut with reference updates", () => {
    test("should update simple cell reference when cell is cut", () => {
      // B1 contains =A1, cut A1 to D5
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        42
      ); // A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=A1"
      ); // B1

      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(42);

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 }, // A1
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 3, // D5
        rowIndex: 4,
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // A1 should be empty
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe("");

      // D5 should have 42
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 3,
          rowIndex: 4,
        })
      ).toBe(42);

      // B1 should now reference D5 and still evaluate to 42
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(42);
      const sheetContent2 = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent2.get("B1")).toBe("=D5");
    });

    test("should update absolute references when cell is cut", () => {
      // B1 contains =$A$1, cut A1 to D5
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        100
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=$A$1"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 3,
        rowIndex: 4,
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // B1 should now reference $D$5
      const sheetContent3 = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent3.get("B1")).toBe("=$D$5");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(100);
    });

    test("should update range reference when entire range is cut", () => {
      // E1 contains =SUM(A1:D5), cut A1:D5 to F10:I14
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        1
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 3, rowIndex: 4 },
        5
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 4, rowIndex: 0 },
        "=SUM(A1:D5)"
      );

      // Build source cells for A1:D5
      const source: CellAddress[] = [];
      for (let row = 0; row <= 4; row++) {
        for (let col = 0; col <= 3; col++) {
          source.push({
            workbookName,
            sheetName,
            colIndex: col,
            rowIndex: row,
          });
        }
      }

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 5, // F10
        rowIndex: 9,
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // E1 should now reference F10:I14
      const sheetContent4 = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent4.get("E1")).toBe("=SUM(F10:I14)");
    });

    test("should NOT update range reference when only partial range is cut", () => {
      // E1 contains =SUM(A1:D5), cut only A1:B5
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        1
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 4, rowIndex: 0 },
        "=SUM(A1:D5)"
      );

      // Build source cells for A1:B5 (partial range)
      const source: CellAddress[] = [];
      for (let row = 0; row <= 4; row++) {
        for (let col = 0; col <= 1; col++) {
          source.push({
            workbookName,
            sheetName,
            colIndex: col,
            rowIndex: row,
          });
        }
      }

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 5, // F10
        rowIndex: 9,
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // E1 should still reference A1:D5 (not updated because partial move)
      const sheetContent5 = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent5.get("E1")).toBe("=SUM(A1:D5)");
    });

    test("should update multiple cells referencing the same cut cell", () => {
      // B1, C1, D1 all reference A1
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        50
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 1, rowIndex: 0 },
        "=A1"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 2, rowIndex: 0 },
        "=A1*2"
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 3, rowIndex: 0 },
        "=A1+A1"
      );

      const source: CellAddress[] = [
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
      ];

      const target: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5, // A6
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // All references should be updated
      const sheetContent6 = engine.getSheetSerialized({
        workbookName,
        sheetName,
      });
      expect(sheetContent6.get("B1")).toBe("=A6");
      expect(sheetContent6.get("C1")).toBe("=A6*2");
      expect(sheetContent6.get("D1")).toBe("=A6+A6");

      // Values should still be correct
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 1,
          rowIndex: 0,
        })
      ).toBe(50);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 2,
          rowIndex: 0,
        })
      ).toBe(100);
      expect(
        engine.getCellValue({
          workbookName,
          sheetName,
          colIndex: 3,
          rowIndex: 0,
        })
      ).toBe(100);
    });

    test("should update cross-sheet references when cell is cut", () => {
      // Add another sheet
      engine.addSheet({ workbookName, sheetName: "Sheet2" });

      // Sheet1!A1 contains value, Sheet2!A1 references it
      engine.setCellContent(
        { workbookName, sheetName: "Sheet1", colIndex: 0, rowIndex: 0 },
        123
      );
      engine.setCellContent(
        { workbookName, sheetName: "Sheet2", colIndex: 0, rowIndex: 0 },
        "=Sheet1!A1"
      );

      expect(
        engine.getCellValue({
          workbookName,
          sheetName: "Sheet2",
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(123);

      // Cut Sheet1!A1 to Sheet1!D5
      const source: CellAddress[] = [
        { workbookName, sheetName: "Sheet1", colIndex: 0, rowIndex: 0 },
      ];

      const target: CellAddress = {
        workbookName,
        sheetName: "Sheet1",
        colIndex: 3,
        rowIndex: 4,
      };

      engine.pasteCells(source, target, {
        cut: true,
        type: "formula",
        include: ["content"],
      });

      // Sheet2!A1 should now reference Sheet1!D5
      const sheet2Content = engine.getSheetSerialized({
        workbookName,
        sheetName: "Sheet2",
      });
      expect(sheet2Content.get("A1")).toBe("=Sheet1!D5");
      expect(
        engine.getCellValue({
          workbookName,
          sheetName: "Sheet2",
          colIndex: 0,
          rowIndex: 0,
        })
      ).toBe(123);
    });
  });
});
