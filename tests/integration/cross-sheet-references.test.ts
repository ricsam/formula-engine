import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../src/core/engine';
import type { CellValue } from '../../src/core/types';

describe('Cross-Sheet References Integration Tests', () => {
  let engine: FormulaEngine;
  let sheet1Id: number;
  let sheet2Id: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheet1Name = engine.addSheet('Sheet1');
    const sheet2Name = engine.addSheet('Sheet2');
    sheet1Id = engine.getSheetId(sheet1Name);
    sheet2Id = engine.getSheetId(sheet2Name);
  });

  describe('Basic Cross-Sheet References', () => {
    test('should handle simple cross-sheet cell reference', () => {
      // Set up data in Sheet1
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 42);
      
      // Reference Sheet1.A1 from Sheet2
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1');
      
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe(42);
    });

    test('should handle quoted sheet names in cross-sheet references', () => {
      // Create sheet with spaces in name
      const sheetWithSpacesName = engine.addSheet('My Data Sheet');
      const sheetWithSpacesId = engine.getSheetId(sheetWithSpacesName);
      
      // Set up data
      engine.setCellContent({ sheet: sheetWithSpacesId, col: 0, row: 0 }, 100);
      
      // Reference using quoted sheet name
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, "='My Data Sheet'!A1");
      
      const result = engine.getCellValue({ sheet: sheet1Id, col: 0, row: 0 });
      expect(result).toBe(100);
    });

    test('should handle cross-sheet range references', () => {
      // Set up data in Sheet1
      const sheet1Data = new Map([
        ['A1', 10],
        ['A2', 20],
        ['A3', 30]
      ]);
      engine.setSheetContent(sheet1Id, sheet1Data);
      
      // SUM a range from Sheet1 in Sheet2
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=SUM(Sheet1!A1:A3)');
      
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe(60);
    });

    test('should handle absolute references in cross-sheet formulas', () => {
      // Set up data
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 50);
      
      // Reference with absolute addressing
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!$A$1 * 2');
      
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe(100);
    });
  });

  describe('Cross-Sheet Formula Combinations', () => {
    test('should handle arithmetic operations with cross-sheet references', () => {
      // Set up data in both sheets
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 15);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, 25);
      
      // Combine values from both sheets
      engine.setCellContent({ sheet: sheet1Id, col: 1, row: 0 }, '=A1 + Sheet2!A1');
      
      const result = engine.getCellValue({ sheet: sheet1Id, col: 1, row: 0 });
      expect(result).toBe(40);
    });

    test('should handle function calls with cross-sheet arguments', () => {
      // Set up data
      const sheet1Data = new Map([
        ['A1', 5],
        ['A2', 10],
        ['A3', 15]
      ]);
      engine.setSheetContent(sheet1Id, sheet1Data);
      
      const sheet2Data = new Map([
        ['A1', 20],
        ['A2', 25]
      ]);
      engine.setSheetContent(sheet2Id, sheet2Data);
      
      // MAX function with ranges from different sheets
      engine.setCellContent({ sheet: sheet1Id, col: 1, row: 0 }, '=MAX(A1:A3, Sheet2!A1:A2)');
      
      const result = engine.getCellValue({ sheet: sheet1Id, col: 1, row: 0 });
      expect(result).toBe(25);
    });

    test('should handle nested cross-sheet references', () => {
      // Create a chain: Sheet1 -> Sheet2 -> Sheet1
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 10);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1 * 2');
      engine.setCellContent({ sheet: sheet1Id, col: 1, row: 0 }, '=Sheet2!A1 + 5');
      
      const result = engine.getCellValue({ sheet: sheet1Id, col: 1, row: 0 });
      expect(result).toBe(25); // (10 * 2) + 5
    });
  });

  describe('Cross-Sheet Dependency Tracking', () => {
    test('should track dependencies across sheets', () => {
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 100);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1');
      
      // Check precedents
      const precedents = engine.getCellPrecedents({ sheet: sheet2Id, col: 0, row: 0 });
      expect(precedents).toHaveLength(1);
      expect(precedents[0]).toEqual({ sheet: sheet1Id, col: 0, row: 0 });
      
      // Check dependents
      const dependents = engine.getCellDependents({ sheet: sheet1Id, col: 0, row: 0 });
      expect(dependents).toHaveLength(1);
      expect(dependents[0]).toEqual({ sheet: sheet2Id, col: 0, row: 0 });
    });

    test('should update cross-sheet dependencies when values change', () => {
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 10);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1 * 3');
      
      // Initial value
      expect(engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 })).toBe(30);
      
      // Update source value
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 20);
      
      // Check that dependent value updated
      expect(engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 })).toBe(60);
    });
  });

  describe('Error Handling', () => {
    test('should handle references to non-existent sheets', () => {
      // Reference a sheet that doesn't exist
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, '=NonExistentSheet!A1');
      
      const result = engine.getCellValue({ sheet: sheet1Id, col: 0, row: 0 });
      expect(result).toBe('#REF!');
    });

    test('should handle circular references across sheets', () => {
      // Create circular reference: Sheet1.A1 -> Sheet2.A1 -> Sheet1.A1
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, '=Sheet2!A1');
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1');
      
      const result1 = engine.getCellValue({ sheet: sheet1Id, col: 0, row: 0 });
      const result2 = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      
      expect(result1).toBe('#CYCLE!');
      expect(result2).toBe('#CYCLE!');
    });

    test('should handle references to empty cells in other sheets', () => {
      // Reference an empty cell from another sheet
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, '=Sheet2!A1');
      
      const result = engine.getCellValue({ sheet: sheet1Id, col: 0, row: 0 });
      expect(result).toBeUndefined(); // Empty cell
    });
  });

  describe('Sheet Operations Impact', () => {
    test('should handle sheet renaming with cross-sheet references', () => {
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 42);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1');
      
      // Rename Sheet1
      engine.renameSheet(sheet1Id, 'DataSheet');
      
      // The reference should still work (though formula text might not update automatically)
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe(42);
    });

    test('should handle sheet deletion with cross-sheet references', () => {
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 42);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=Sheet1!A1');
      
      // Verify initial state
      expect(engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 })).toBe(42);
      
      // Delete the referenced sheet
      engine.removeSheet(sheet1Id);
      
      // The reference should now return an error
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe('#REF!');
    });
  });

  describe('Complex Cross-Sheet Scenarios', () => {
    test('should handle VLOOKUP across sheets', () => {
      // Set up lookup table in Sheet1
      const lookupData = new Map<string, any>([
        ['A1', 'Product'],
        ['B1', 'Price'],
        ['A2', 'Apple'],
        ['B2', 100],
        ['A3', 'Banana'],
        ['B3', 150],
        ['A4', 'Cherry'],
        ['B4', 200]
      ]);
      engine.setSheetContent(sheet1Id, lookupData);
      
      // Perform VLOOKUP from Sheet2
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=VLOOKUP("Banana", Sheet1!A1:B4, 2, FALSE())');
      
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe(150);
    });

    test('should handle conditional logic with cross-sheet data', () => {
      engine.setCellContent({ sheet: sheet1Id, col: 0, row: 0 }, 75);
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=IF(Sheet1!A1 > 50, "High", "Low")');
      
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe('High');
    });

    test('should handle array formulas with cross-sheet ranges', () => {
      // Set up data for filtering
      const filterData = new Map([
        ['A1', 10],
        ['A2', 25],
        ['A3', 15],
        ['A4', 30],
        ['A5', 5]
      ]);
      engine.setSheetContent(sheet1Id, filterData);
      
      // Filter values > 20 from another sheet
      engine.setCellContent({ sheet: sheet2Id, col: 0, row: 0 }, '=FILTER(Sheet1!A1:A5, Sheet1!A1:A5 > 20)');
      
      const result = engine.getCellValue({ sheet: sheet2Id, col: 0, row: 0 });
      expect(result).toBe(25); // First filtered value
    });
  });
});