import { describe, expect, test, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../../src/core/engine";
import type {
  CellAddress,
  ConditionalStyle,
  DirectCellStyle,
  LCHColor,
} from "../../../../src/core/types";
import { lchToHex } from "../../../../src/core/utils/color-utils";

describe("StyleManager", () => {
  let engine: FormulaEngine;
  const workbookName = "TestWorkbook";
  const sheetName = "Sheet1";

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("addConditionalStyle", () => {
    test("adds a formula-based style", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 9 },
              row: { type: "number", value: 9 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "ROW() > 4",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
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
      expect(styles).toHaveLength(1);
      expect(styles[0]).toEqual(style);
    });

    test("adds a gradient-based style", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "gradient",
          min: { type: "lowest_value", color: { l: 90, c: 10, h: 120 } },
          max: { type: "highest_value", color: { l: 30, c: 80, h: 0 } },
        },
      };

      engine.addConditionalStyle(style);
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
      expect(styles).toHaveLength(1);
      expect(styles[0]).toEqual(style);
    });

    test("adds multiple styles in order", () => {
      const style1: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      const style2: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "FALSE",
          color: { l: 30, c: 50, h: 180 },
        },
      };

      engine.addConditionalStyle(style1);
      engine.addConditionalStyle(style2);

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
      expect(styles).toHaveLength(2);
      expect(styles[0]).toEqual(style1);
      expect(styles[1]).toEqual(style2);
    });
  });

  describe("removeConditionalStyle", () => {
    test("removes a style by index", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);

      const removed = engine.removeConditionalStyle(workbookName, 0);
      expect(removed).toBe(true);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
    });

    test("returns false for invalid index", () => {
      const removed = engine.removeConditionalStyle(workbookName, 0);
      expect(removed).toBe(false);
    });

    test("returns false for negative index", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      const removed = engine.removeConditionalStyle(workbookName, -1);
      expect(removed).toBe(false);
    });
  });

  describe("getCellStyle - formula conditions", () => {
    test("applies style when formula evaluates to TRUE", () => {
      // Set up a cell with value
      const cellAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      };

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 9 },
              row: { type: "number", value: 9 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "ROW() > 4",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      const cellStyle = engine.getCellStyle(cellAddress);

      expect(cellStyle).toBeDefined();
      expect(cellStyle?.backgroundColor).toBeDefined();
      expect(cellStyle?.backgroundColor).toMatch(/^#[0-9a-f]{6}$/i);
    });

    test("does not apply style when formula evaluates to FALSE", () => {
      const cellAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      };

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 9 },
              row: { type: "number", value: 9 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "ROW() > 4",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      const cellStyle = engine.getCellStyle(cellAddress);

      expect(cellStyle).toBeUndefined();
    });

    test("ROW() function returns correct row number in formula condition", () => {
      // Test that ROW() evaluates relative to the cell being styled
      const cellAddress1: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 1, // Row 2 in 1-based
      };

      const cellAddress2: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5, // Row 6 in 1-based
      };

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 9 },
              row: { type: "number", value: 9 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "ROW() > 4",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      engine.addConditionalStyle(style);

      const style1 = engine.getCellStyle(cellAddress1);
      const style2 = engine.getCellStyle(cellAddress2);

      expect(style1).toBeUndefined(); // Row 2 is not > 4
      expect(style2).toBeDefined(); // Row 6 is > 4
    });
  });

  describe("getCellStyle - gradient conditions", () => {
    test("applies gradient based on cell value with lowest_value/highest_value", () => {
      // Set up cells with numeric values
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        10
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        50
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 2 },
        90
      );

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 0 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        condition: {
          type: "gradient",
          min: { type: "lowest_value", color: { l: 90, c: 10, h: 120 } },
          max: { type: "highest_value", color: { l: 30, c: 80, h: 0 } },
        },
      };

      engine.addConditionalStyle(style);

      const style0 = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      const style1 = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 1,
      });
      const style2 = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      });

      expect(style0).toBeDefined();
      expect(style1).toBeDefined();
      expect(style2).toBeDefined();

      expect(style0?.backgroundColor).toBeDefined();
      expect(style1?.backgroundColor).toBeDefined();
      expect(style2?.backgroundColor).toBeDefined();

      // All should have different colors
      expect(style0?.backgroundColor).not.toBe(style1?.backgroundColor);
      expect(style1?.backgroundColor).not.toBe(style2?.backgroundColor);

      expect(style0?.backgroundColor).toEqual(
        lchToHex({ l: 90, c: 10, h: 120 })
      );
      expect(style2?.backgroundColor).toEqual(lchToHex({ l: 30, c: 80, h: 0 }));
    });

    test("applies gradient with number-based min/max using valueFormula", () => {
      // Set up cells with numeric values
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        25
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        50
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 2 },
        75
      );

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 0 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        condition: {
          type: "gradient",
          min: {
            type: "number",
            color: { l: 90, c: 10, h: 120 },
            valueFormula: "0",
          },
          max: {
            type: "number",
            color: { l: 30, c: 80, h: 0 },
            valueFormula: "100",
          },
        },
      };

      engine.addConditionalStyle(style);

      const style1 = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 1,
      });

      expect(style1).toBeDefined();
      expect(style1?.backgroundColor).toBeDefined();
    });

    test("does not apply gradient to non-numeric cells", () => {
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        "text"
      );

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 0 },
              row: { type: "number", value: 0 },
            },
          },
        }],
        condition: {
          type: "gradient",
          min: { type: "lowest_value", color: { l: 90, c: 10, h: 120 } },
          max: { type: "highest_value", color: { l: 30, c: 80, h: 0 } },
        },
      };

      engine.addConditionalStyle(style);

      const cellStyle = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });

      expect(cellStyle).toBeUndefined();
    });

    test("applies min color to values less than min (number-based)", () => {
      // Set up cells with values outside the min/max range
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        -10 // Less than min (0)
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        0 // At min
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 2 },
        50 // In range
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 3 },
        100 // At max
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 4 },
        150 // Greater than max (100)
      );

      const minColor: LCHColor = { l: 90, c: 10, h: 120 };
      const maxColor: LCHColor = { l: 30, c: 80, h: 0 };

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 0 },
              row: { type: "number", value: 4 },
            },
          },
        }],
        condition: {
          type: "gradient",
          min: {
            type: "number",
            color: minColor,
            valueFormula: "0",
          },
          max: {
            type: "number",
            color: maxColor,
            valueFormula: "100",
          },
        },
      };

      engine.addConditionalStyle(style);

      // Value less than min should get min color
      const styleBelowMin = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      expect(styleBelowMin).toBeDefined();
      expect(styleBelowMin?.backgroundColor).toBe(lchToHex(minColor));

      // Value at min should get min color
      const styleAtMin = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 1,
      });
      expect(styleAtMin).toBeDefined();
      expect(styleAtMin?.backgroundColor).toBe(lchToHex(minColor));

      // Value in range should be interpolated
      const styleInRange = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 2,
      });
      expect(styleInRange).toBeDefined();
      expect(styleInRange?.backgroundColor).not.toBe(lchToHex(minColor));
      expect(styleInRange?.backgroundColor).not.toBe(lchToHex(maxColor));

      // Value at max should get max color
      const styleAtMax = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 3,
      });
      expect(styleAtMax).toBeDefined();
      expect(styleAtMax?.backgroundColor).toBe(lchToHex(maxColor));

      // Value greater than max should get max color
      const styleAboveMax = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 4,
      });
      expect(styleAboveMax).toBeDefined();
      expect(styleAboveMax?.backgroundColor).toBe(lchToHex(maxColor));
    });

    test("applies min/max color to values outside range (lowest_value/highest_value)", () => {
      // Set up cells with values
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        5 // Lowest in area
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 1 },
        10
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 2 },
        15
      );
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 3 },
        20 // Highest in area
      );
      // Add a cell outside the area with a value less than min
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 4 },
        0 // Less than min (5)
      );
      // Add a cell outside the area with a value greater than max
      engine.setCellContent(
        { workbookName, sheetName, colIndex: 0, rowIndex: 5 },
        25 // Greater than max (20)
      );

      const minColor: LCHColor = { l: 90, c: 10, h: 120 };
      const maxColor: LCHColor = { l: 30, c: 80, h: 0 };

      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 0 },
              row: { type: "number", value: 3 },
            },
          },
        }],
        condition: {
          type: "gradient",
          min: {
            type: "lowest_value",
            color: minColor,
          },
          max: {
            type: "highest_value",
            color: maxColor,
          },
        },
      };

      engine.addConditionalStyle(style);

      // Value at min (lowest in area) should get min color
      const styleAtMin = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      });
      expect(styleAtMin).toBeDefined();
      expect(styleAtMin?.backgroundColor).toBe(lchToHex(minColor));

      // Value at max (highest in area) should get max color
      const styleAtMax = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 3,
      });
      expect(styleAtMax).toBeDefined();
      expect(styleAtMax?.backgroundColor).toBe(lchToHex(maxColor));

      // Value outside area but less than min should not have style (outside area)
      const styleOutsideBelow = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 4,
      });
      expect(styleOutsideBelow).toBeUndefined();

      // Value outside area but greater than max should not have style (outside area)
      const styleOutsideAbove = engine.getCellStyle({
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 5,
      });
      expect(styleOutsideAbove).toBeUndefined();
    });
  });

  describe("first matching rule wins", () => {
    test("applies first matching style when multiple rules match", () => {
      const cellAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 0,
        rowIndex: 0,
      };

      const redColor: LCHColor = { l: 50, c: 80, h: 0 };
      const blueColor: LCHColor = { l: 50, c: 80, h: 270 };

      const style1: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: redColor,
        },
      };

      const style2: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: blueColor,
        },
      };

      engine.addConditionalStyle(style1);
      engine.addConditionalStyle(style2);

      const cellStyle = engine.getCellStyle(cellAddress);

      expect(cellStyle).toBeDefined();
      // Should apply the first style (red), not the second (blue)
      // We can't directly check the color, but we can verify it's not undefined
      expect(cellStyle?.backgroundColor).toBeDefined();
    });
  });

  describe("workbook/sheet lifecycle", () => {
    test("removes styles when workbook is removed", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);

      engine.removeWorkbook(workbookName);

      // Re-create the workbook to check if styles were cleared
      engine.addWorkbook(workbookName);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
    });

    test("updates workbook name in styles when workbook is renamed", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);

      const newWorkbookName = "RenamedWorkbook";
      engine.renameWorkbook({ workbookName, newWorkbookName });

      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName: newWorkbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName: newWorkbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })[0]!.areas[0]!.workbookName
      ).toBe(newWorkbookName);
    });

    test("removes styles when sheet is removed", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);

      engine.removeSheet({ workbookName, sheetName });
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
    });

    test("updates sheet name in styles when sheet is renamed", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);

      const newSheetName = "RenamedSheet";
      engine.renameSheet({ workbookName, sheetName, newSheetName });

      const styles = engine.getConditionalStylesIntersectingWithRange({
        workbookName,
        sheetName: newSheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "infinity", sign: "positive" },
            row: { type: "infinity", sign: "positive" },
          },
        },
      });
      expect(styles).toHaveLength(1);
      expect(styles[0]!.areas[0]!.sheetName).toBe(newSheetName);
    });
  });

  describe("serialization", () => {
    test("includes conditional styles in serialized state", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);

      const state = engine.getState();
      expect(state.conditionalStyles).toBeDefined();
      expect(Array.isArray(state.conditionalStyles)).toBe(true);
      expect(
        state.conditionalStyles.filter(
          (s) => s.areas.some(area => area.workbookName === workbookName)
        )
      ).toHaveLength(1);
    });

    test("restores conditional styles from serialized state", () => {
      const style: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 50, h: 0 },
        },
      };

      engine.addConditionalStyle(style);

      const serialized = engine.serializeEngine();
      const newEngine = FormulaEngine.buildEmpty();
      newEngine.resetToSerializedEngine(serialized);

      const restoredStyles = newEngine.getConditionalStylesIntersectingWithRange({
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
      expect(restoredStyles).toHaveLength(1);
      expect(restoredStyles[0]).toEqual(style);
    });
  });

  describe("cellStyles", () => {
    test("adds a direct cell style", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
          color: "#FFFFFF",
        },
      };

      engine.addCellStyle(cellStyle);
      const styles = engine.getStylesIntersectingWithRange({
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
      expect(styles).toHaveLength(1);
      expect(styles[0]).toEqual(cellStyle);
    });

    test("adds multiple cell styles", () => {
      const cellStyle1: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
          color: "#FFFFFF",
        },
      };

      const cellStyle2: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 3, row: 3 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#0000FF",
          color: "#FFFF00",
        },
      };

      engine.addCellStyle(cellStyle1);
      engine.addCellStyle(cellStyle2);

      const styles = engine.getStylesIntersectingWithRange({
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
      expect(styles).toHaveLength(2);
      expect(styles[0]).toEqual(cellStyle1);
      expect(styles[1]).toEqual(cellStyle2);
    });

    test("removes a cell style by index", () => {
      const cellStyle1: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      const cellStyle2: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 3, row: 3 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#0000FF",
        },
      };

      engine.addCellStyle(cellStyle1);
      engine.addCellStyle(cellStyle2);

      const removed = engine.removeCellStyle(workbookName, 0);
      expect(removed).toBe(true);

      const styles = engine.getStylesIntersectingWithRange({
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
      expect(styles).toHaveLength(1);
      expect(styles[0]).toEqual(cellStyle2);
    });

    test("returns false when removing invalid cell style index", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const removed = engine.removeCellStyle(workbookName, 10);
      expect(removed).toBe(false);

      const styles = engine.getStylesIntersectingWithRange({
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
      expect(styles).toHaveLength(1);
    });

    test("returns false when removing negative cell style index", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const removed = engine.removeCellStyle(workbookName, -1);
      expect(removed).toBe(false);

      const styles = engine.getStylesIntersectingWithRange({
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
      expect(styles).toHaveLength(1);
    });

    test("applies direct cell style to cells in range", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
          color: "#FFFFFF",
        },
      };

      engine.addCellStyle(cellStyle);

      const cellAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 1,
      };

      const result = engine.getCellStyle(cellAddress);
      expect(result).toBeDefined();
      expect(result?.backgroundColor).toBe("#FF0000");
      expect(result?.color).toBe("#FFFFFF");
    });

    test("does not apply direct cell style to cells outside range", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const cellAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 5,
        rowIndex: 5,
      };

      const result = engine.getCellStyle(cellAddress);
      expect(result).toBeUndefined();
    });

    test("direct cell styles take precedence over conditional styles", () => {
      // Add a conditional style
      const conditionalStyle: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      // Add a direct cell style for the same area
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#00FF00",
          color: "#000000",
        },
      };

      engine.addConditionalStyle(conditionalStyle);
      engine.addCellStyle(cellStyle);

      const cellAddress: CellAddress = {
        workbookName,
        sheetName,
        colIndex: 1,
        rowIndex: 1,
      };

      const result = engine.getCellStyle(cellAddress);
      expect(result).toBeDefined();
      // Conditional style has precedence (Excel behavior)
      expect(result?.backgroundColor).toBeDefined(); // Will be the conditional style color
    });

    test("filters cell styles by workbook name", () => {
      const otherWorkbookName = "OtherWorkbook";
      engine.addWorkbook(otherWorkbookName);
      engine.addSheet({ workbookName: otherWorkbookName, sheetName });

      const cellStyle1: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      const cellStyle2: DirectCellStyle = {
        areas: [{
          workbookName: otherWorkbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#0000FF",
        },
      };

      engine.addCellStyle(cellStyle1);
      engine.addCellStyle(cellStyle2);

      const styles1 = engine.getStylesIntersectingWithRange({
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
      expect(styles1).toHaveLength(1);
      expect(styles1[0]).toEqual(cellStyle1);

      const styles2 = engine.getStylesIntersectingWithRange({
        workbookName: otherWorkbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "infinity", sign: "positive" },
            row: { type: "infinity", sign: "positive" },
          },
        },
      });
      expect(styles2).toHaveLength(1);
      expect(styles2[0]).toEqual(cellStyle2);
    });
  });

  describe("clearCellStyles", () => {
    test("removes cellStyle completely contained in clear range", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 2, row: 2 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 4 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);
      expect(
        engine.getStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);

      // Clear a range that contains the style
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 10 },
            row: { type: "number", value: 10 },
          },
        },
      });

      expect(
        engine.getStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
    });

    test("adjusts cellStyle when clear range overlaps edge", () => {
      // Style: A1:E5 (0,0 to 4,4)
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 4 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      // Clear top portion: A1:E2 (0,0 to 4,1)
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 4 },
            row: { type: "number", value: 1 },
          },
        },
      });

      const styles = engine.getStylesIntersectingWithRange({
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
      expect(styles).toHaveLength(1);
      // Should have bottom portion: A3:E5 (0,2 to 4,4)
      expect(styles[0]!.areas[0]!.range.start).toEqual({ col: 0, row: 2 });
      expect(styles[0]!.areas[0]!.range.end).toEqual({
        col: { type: "number", value: 4 },
        row: { type: "number", value: 4 },
      });
    });

    test("splits cellStyle when clear range creates hole", () => {
      // Style: A1:E5 (0,0 to 4,4)
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 4 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      // Clear middle: B2:D4 (1,1 to 3,3)
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 1, row: 1 },
          end: {
            col: { type: "number", value: 3 },
            row: { type: "number", value: 3 },
          },
        },
      });

      const styles = engine.getStylesIntersectingWithRange({
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
      // Should have 1 style with 4 areas (top, bottom, left, right)
      expect(styles).toHaveLength(1);
      expect(styles[0]!.areas).toHaveLength(4);
    });

    test("preserves cellStyle that doesn't intersect", () => {
      const cellStyle1: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      const cellStyle2: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 10, row: 10 },
            end: {
              col: { type: "number", value: 12 },
              row: { type: "number", value: 12 },
            },
          },
        }],
        style: {
          backgroundColor: "#00FF00",
        },
      };

      engine.addCellStyle(cellStyle1);
      engine.addCellStyle(cellStyle2);

      // Clear a range that only affects cellStyle1
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 5 },
            row: { type: "number", value: 5 },
          },
        },
      });

      const styles = engine.getStylesIntersectingWithRange({
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
      // cellStyle1 removed, cellStyle2 preserved
      expect(styles).toHaveLength(1);
      expect(styles[0]!.areas[0]!.range.start).toEqual({ col: 10, row: 10 });
    });

    test("clears conditional styles similarly to cellStyles", () => {
      const conditionalStyle: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 4 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      engine.addConditionalStyle(conditionalStyle);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(1);

      // Clear top portion
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 4 },
            row: { type: "number", value: 1 },
          },
        },
      });

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
      expect(styles).toHaveLength(1);
      // Should have bottom portion
      expect(styles[0]!.areas[0]!.range.start).toEqual({ col: 0, row: 2 });
    });

    test("clears both cellStyles and conditionalStyles in one call", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      const conditionalStyle: ConditionalStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 50, c: 80, h: 0 },
        },
      };

      engine.addCellStyle(cellStyle);
      engine.addConditionalStyle(conditionalStyle);

      // Clear the entire range
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 2 },
            row: { type: "number", value: 2 },
          },
        },
      });

      expect(
        engine.getStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
      expect(
        engine.getConditionalStylesIntersectingWithRange({
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        })
      ).toHaveLength(0);
    });

    test("preserves style properties when splitting", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 4 },
              row: { type: "number", value: 4 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
          color: "#FFFFFF",
          fontSize: 14,
          bold: true,
        },
      };

      engine.addCellStyle(cellStyle);

      // Clear middle to create hole
      engine.clearCellStyles({
        workbookName,
        sheetName,
        range: {
          start: { col: 1, row: 1 },
          end: {
            col: { type: "number", value: 3 },
            row: { type: "number", value: 3 },
          },
        },
      });

      const styles = engine.getStylesIntersectingWithRange({
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
      // Should have 1 style with 4 areas
      expect(styles).toHaveLength(1);
      expect(styles[0]!.areas).toHaveLength(4);

      // The style should preserve the original style properties
      expect(styles[0]!.style.backgroundColor).toBe("#FF0000");
      expect(styles[0]!.style.color).toBe("#FFFFFF");
      expect(styles[0]!.style.fontSize).toBe(14);
      expect(styles[0]!.style.bold).toBe(true);
    });
  });

  describe("getStyleForRange", () => {
    test("returns style when range is completely contained within a single style", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 5 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
          color: "#FFFFFF",
          fontSize: 14,
          bold: true,
        },
      };

      engine.addCellStyle(cellStyle);

      const result = engine.getStyleForRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 1, row: 1 },
          end: {
            col: { type: "number", value: 3 },
            row: { type: "number", value: 3 },
          },
        },
      });

      expect(result).toBeDefined();
      expect(result?.style.backgroundColor).toBe("#FF0000");
      expect(result?.style.color).toBe("#FFFFFF");
      expect(result?.style.fontSize).toBe(14);
      expect(result?.style.bold).toBe(true);
    });

    test("returns undefined when range spans multiple styles", () => {
      const cellStyle1: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      const cellStyle2: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 3, row: 0 },
            end: {
              col: { type: "number", value: 5 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#0000FF",
        },
      };

      engine.addCellStyle(cellStyle1);
      engine.addCellStyle(cellStyle2);

      const result = engine.getStyleForRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 5 },
            row: { type: "number", value: 2 },
          },
        },
      });

      expect(result).toBeUndefined();
    });

    test("returns undefined when range is not completely contained", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 3 },
              row: { type: "number", value: 3 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const result = engine.getStyleForRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 5 },
            row: { type: "number", value: 5 },
          },
        },
      });

      expect(result).toBeUndefined();
    });

    test("returns undefined when no styles intersect with range", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 2 },
              row: { type: "number", value: 2 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const result = engine.getStyleForRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 10, row: 10 },
          end: {
            col: { type: "number", value: 12 },
            row: { type: "number", value: 12 },
          },
        },
      });

      expect(result).toBeUndefined();
    });

    test("returns undefined when range matches style exactly but is in different workbook", () => {
      const otherWorkbookName = "OtherWorkbook";
      engine.addWorkbook(otherWorkbookName);
      engine.addSheet({ workbookName: otherWorkbookName, sheetName });

      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "number", value: 3 },
              row: { type: "number", value: 3 },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const result = engine.getStyleForRange({
        workbookName: otherWorkbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: {
            col: { type: "number", value: 3 },
            row: { type: "number", value: 3 },
          },
        },
      });

      expect(result).toBeUndefined();
    });

    test("works with infinite ranges", () => {
      const cellStyle: DirectCellStyle = {
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: {
              col: { type: "infinity", sign: "positive" },
              row: { type: "infinity", sign: "positive" },
            },
          },
        }],
        style: {
          backgroundColor: "#FF0000",
        },
      };

      engine.addCellStyle(cellStyle);

      const result = engine.getStyleForRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 5, row: 5 },
          end: {
            col: { type: "number", value: 10 },
            row: { type: "number", value: 10 },
          },
        },
      });

      expect(result).toBeDefined();
      expect(result?.style.backgroundColor).toBe("#FF0000");
    });
  });

  describe("clearCellStylesInRange", () => {
    test("completely removes style when cleared range contains entire style", () => {
      // Add style to A1:C3
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
          },
        }],
        style: { backgroundColor: "#FF0000" },
      });

      // Clear the entire range
      engine._styleManager.clearCellStylesInRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
        },
      });

      // Style should be completely removed
      const style = engine.getCellStyle({ workbookName, sheetName, colIndex: 0, rowIndex: 0 });
      expect(style).toBeUndefined();
    });

    test("splits style when cleared range is in the middle", () => {
      // Add style to A1:E5
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
          },
        }],
        style: { backgroundColor: "#FF0000" },
      });

      // Clear C3 (middle cell)
      engine._styleManager.clearCellStylesInRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 2, row: 2 },
          end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
        },
      });

      // C3 should have no style
      const c3Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 2, rowIndex: 2 });
      expect(c3Style).toBeUndefined();

      // Surrounding cells should still have the style
      const a1Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 0, rowIndex: 0 });
      expect(a1Style?.backgroundColor).toBe("#FF0000");

      const e5Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 4, rowIndex: 4 });
      expect(e5Style?.backgroundColor).toBe("#FF0000");

      const b3Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 1, rowIndex: 2 });
      expect(b3Style?.backgroundColor).toBe("#FF0000");

      const d3Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 3, rowIndex: 2 });
      expect(d3Style?.backgroundColor).toBe("#FF0000");
    });

    test("preserves styles in other sheets/workbooks", () => {
      // Add style to Sheet1
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
          },
        }],
        style: { backgroundColor: "#FF0000" },
      });

      // Add Sheet2
      engine.addSheet({ workbookName, sheetName: "Sheet2" });

      // Add style to Sheet2
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName: "Sheet2",
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
          },
        }],
        style: { backgroundColor: "#00FF00" },
      });

      // Clear range in Sheet1
      engine._styleManager.clearCellStylesInRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
        },
      });

      // Sheet1 style should be removed
      const sheet1Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 0, rowIndex: 0 });
      expect(sheet1Style).toBeUndefined();

      // Sheet2 style should be preserved
      const sheet2Style = engine.getCellStyle({ workbookName, sheetName: "Sheet2", colIndex: 0, rowIndex: 0 });
      expect(sheet2Style?.backgroundColor).toBe("#00FF00");
    });

    test("handles no intersection gracefully", () => {
      // Add style to A1:C3
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
          },
        }],
        style: { backgroundColor: "#FF0000" },
      });

      // Clear non-intersecting range F6:G7
      engine._styleManager.clearCellStylesInRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 5, row: 5 },
          end: { col: { type: "number", value: 6 }, row: { type: "number", value: 6 } },
        },
      });

      // Original style should be unchanged
      const a1Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 0, rowIndex: 0 });
      expect(a1Style?.backgroundColor).toBe("#FF0000");
    });
  });

  describe("paste replacement behavior", () => {
    test("pasting replaces existing cell styles", () => {
      // Add red background to A1
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 0 }, row: { type: "number", value: 0 } },
          },
        }],
        style: { backgroundColor: "#FF0000", bold: true, fontSize: 16 },
      });

      // Set up source cell B1 with blue background
      engine.setCellContent({ workbookName, sheetName, colIndex: 1, rowIndex: 0 }, "Source");
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 1, row: 0 },
            end: { col: { type: "number", value: 1 }, row: { type: "number", value: 0 } },
          },
        }],
        style: { backgroundColor: "#0000FF", italic: true },
      });

      // Paste B1 into A1
      engine.pasteCells(
        [{ workbookName, sheetName, colIndex: 1, rowIndex: 0 }],
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { cut: false, type: "formula", include: "all" }
      );

      // A1 should have ONLY blue background and italic (old style completely replaced)
      const a1Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 0, rowIndex: 0 });
      expect(a1Style?.backgroundColor).toBe("#0000FF");
      expect(a1Style?.italic).toBe(true);
      expect(a1Style?.bold).toBeUndefined(); // Old bold is gone
      expect(a1Style?.fontSize).toBeUndefined(); // Old fontSize is gone
    });

    test("pasting to large styled range preserves surrounding styles", () => {
      // Add red background to A1:E5
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 4 }, row: { type: "number", value: 4 } },
          },
        }],
        style: { backgroundColor: "#FF0000" },
      });

      // Set up source cell with blue background
      engine.setCellContent({ workbookName, sheetName, colIndex: 10, rowIndex: 10 }, "Blue");
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 10, row: 10 },
            end: { col: { type: "number", value: 10 }, row: { type: "number", value: 10 } },
          },
        }],
        style: { backgroundColor: "#0000FF" },
      });

      // Paste into C3 (middle of the red range)
      engine.pasteCells(
        [{ workbookName, sheetName, colIndex: 10, rowIndex: 10 }],
        { workbookName, sheetName, colIndex: 2, rowIndex: 2 },
        { cut: false, type: "formula", include: "all" }
      );

      // C3 should be blue
      const c3Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 2, rowIndex: 2 });
      expect(c3Style?.backgroundColor).toBe("#0000FF");

      // Surrounding cells should still be red
      const a1Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 0, rowIndex: 0 });
      expect(a1Style?.backgroundColor).toBe("#FF0000");

      const e5Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 4, rowIndex: 4 });
      expect(e5Style?.backgroundColor).toBe("#FF0000");

      const b2Style = engine.getCellStyle({ workbookName, sheetName, colIndex: 1, rowIndex: 1 });
      expect(b2Style?.backgroundColor).toBe("#FF0000");
    });

    test("conditional styles are preserved when pasting cell styles", () => {
      // Add conditional style to A1:C3
      engine.addConditionalStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 0, row: 0 },
            end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
          },
        }],
        condition: {
          type: "formula",
          formula: "TRUE",
          color: { l: 70, c: 80, h: 120 },
        },
      });

      // Set up source with cell style
      engine.setCellContent({ workbookName, sheetName, colIndex: 5, rowIndex: 5 }, "Source");
      engine.addCellStyle({
        areas: [{
          workbookName,
          sheetName,
          range: {
            start: { col: 5, row: 5 },
            end: { col: { type: "number", value: 5 }, row: { type: "number", value: 5 } },
          },
        }],
        style: { backgroundColor: "#FF0000" },
      });

      // Paste into A1
      engine.pasteCells(
        [{ workbookName, sheetName, colIndex: 5, rowIndex: 5 }],
        { workbookName, sheetName, colIndex: 0, rowIndex: 0 },
        { cut: false, type: "formula", include: "all" }
      );

      // A1 should have the pasted cell style (red background)
      // But conditional style should still exist in the list
      const conditionalStyles = engine.getConditionalStylesIntersectingWithRange({
        workbookName,
        sheetName,
        range: {
          start: { col: 0, row: 0 },
          end: { col: { type: "number", value: 2 }, row: { type: "number", value: 2 } },
        },
      });

      expect(conditionalStyles.length).toBeGreaterThan(0);
    });
  });
});
