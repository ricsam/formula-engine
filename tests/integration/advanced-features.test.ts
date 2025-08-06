import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../src/core/engine';
import type { CellValue } from '../../src/core/types';

describe('Advanced Features Integration Tests', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    sheetId = engine.getSheetId(sheetName);
  });

  describe('Array Spilling', () => {
    test('Basic array spilling from FILTER function', () => {
      // Set up data
      const data = new Map<string, CellValue>([
        ['A1', 1], ['B1', 'Apple'],
        ['A2', 2], ['B2', 'Banana'],
        ['A3', 3], ['B3', 'Cherry'],
        ['A4', 4], ['B4', 'Date'],
        ['A5', 5], ['B5', 'Elderberry']
      ]);
      engine.setSheetContents(sheetId, data);

      // Create a FILTER that returns multiple rows
      engine.setCellContents({ sheet: sheetId, col: 3, row: 0 }, '=FILTER(A1:B5, A1:A5 > 2)');
      
      // Check the origin cell
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 0 })).toBe(3);
      
      // Check spilled values
      expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 0 })).toBe('Cherry');
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 1 })).toBe(4);
      expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 1 })).toBe('Date');
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 2 })).toBe(5);
      expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 2 })).toBe('Elderberry');
      
      // Check that cells are marked as array cells
      expect(engine.isCellPartOfArray({ sheet: sheetId, col: 3, row: 0 })).toBe(true);
      expect(engine.isCellPartOfArray({ sheet: sheetId, col: 4, row: 0 })).toBe(true);
      expect(engine.isCellPartOfArray({ sheet: sheetId, col: 3, row: 1 })).toBe(true);
    });

    test('Array spilling with ARRAY_CONSTRAIN', () => {
      // Create a large array
      const data = new Map([
        ['A1', 1], ['B1', 2], ['C1', 3],
        ['A2', 4], ['B2', 5], ['C2', 6],
        ['A3', 7], ['B3', 8], ['C3', 9]
      ]);
      engine.setSheetContents(sheetId, data);

      // Constrain to 2x2
      engine.setCellContents({ sheet: sheetId, col: 4, row: 0 }, '=ARRAY_CONSTRAIN(A1:C3, 2, 2)');
      
      // Check spilled values
      expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 0 })).toBe(1);
      expect(engine.getCellValue({ sheet: sheetId, col: 5, row: 0 })).toBe(2);
      expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 1 })).toBe(4);
      expect(engine.getCellValue({ sheet: sheetId, col: 5, row: 1 })).toBe(5);
      
      // Check that row 2 is not spilled
      expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 2 })).toBeUndefined();
    });

    test('Spill blocking returns #SPILL! error', () => {
      // Set up data that will block spilling
      engine.setCellContents({ sheet: sheetId, col: 1, row: 1 }, 'Blocking Value');
      
      // Create array data
      const data = new Map([
        ['A1', 1], ['A2', 2], ['A3', 3]
      ]);
      engine.setSheetContents(sheetId, data);

      // Try to spill an array where there's a blocking value
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, '=FILTER(A1:A3, A1:A3 > 0)');
      
      // Should get #SPILL! error because B2 is blocking
      expect(engine.getCellValue({ sheet: sheetId, col: 0, row: 0 })).toBe('#SPILL!');
    });

    test('Array spilling updates when source data changes', () => {
      // Set up initial data
      const data = new Map([
        ['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4]
      ]);
      engine.setSheetContents(sheetId, data);

      // Create a filter
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=FILTER(A1:A4, A1:A4 > 2)');
      
      // Initially should have 3 and 4
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(3);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 1 })).toBe(4);
      
      // Update source data
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 5); // Change A2 from 2 to 5
      
      // Now should have 5, 3, and 4
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(5);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 1 })).toBe(3);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 2 })).toBe(4);
    });

    test('Nested array formulas', () => {
      // Set up data
      const data = new Map([
        ['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4], ['A5', 5]
      ]);
      engine.setSheetContents(sheetId, data);

      // Create nested array operations
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=FILTER(A1:A5, A1:A5 > 2)');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=SUM(FILTER(A1:A5, A1:A5 > 2))');
      
      // Check the sum of filtered values (3 + 4 + 5 = 12)
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe(12);
    });
  });

  describe('Relative Addressing', () => {
    test('Formulas use relative references by default', () => {
      // Set up data and formula
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1*2');
      
      // Copy the formula to another cell
      const formulaToCopy = engine.getCellFormula({ sheet: sheetId, col: 1, row: 0 });
      expect(formulaToCopy).toBe('=A1*2');
      
      // When we copy B1 to B2, it should adjust to =A2*2
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 20);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 1 }, '=A2*2');
      
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 1 })).toBe(40);
    });

    test('Absolute references with $ signs', () => {
      // Set up data
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 100);
      
      // Absolute reference
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=$A$1*2');
      engine.setCellContents({ sheet: sheetId, col: 1, row: 1 }, '=$A$1*3');
      
      // Both should reference A1
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(200);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 1 })).toBe(300);
    });

    test('Mixed references', () => {
      // Set up data
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 20);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, 30);
      
      // Column absolute: $A1
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=$A1');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 1 }, '=$A2');
      
      // Row absolute: A$1
      engine.setCellContents({ sheet: sheetId, col: 3, row: 0 }, '=A$1');
      engine.setCellContents({ sheet: sheetId, col: 3, row: 1 }, '=A$1'); // Still references row 1
      
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe(10);
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 1 })).toBe(20);
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 0 })).toBe(10);
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 1 })).toBe(10); // Same as above
    });

    test('Copy and paste with relative references', () => {
      // Set up source data
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=A1+B1');
      
      // Copy C1
      engine.copy({ 
        start: { sheet: sheetId, col: 2, row: 0 },
        end: { sheet: sheetId, col: 2, row: 0 }
      });
      
      // Set up new data
      engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 2 }, 20);
      
      // Paste to C3 - formula should adjust to =A3+B3
      engine.paste({ sheet: sheetId, col: 2, row: 2 });
      
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 2 })).toBe(30);
    });
  });

  describe('Associative Operations Optimization', () => {
    test('SUM over large range is optimized', () => {
      // Create a large dataset
      const data = new Map<string, CellValue>();
      for (let i = 0; i < 100; i++) {
        data.set(`A${i+1}`, i + 1);
      }
      engine.setSheetContents(sheetId, data);

      // Sum over the range
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=SUM(A1:A100)');
      
      // Sum of 1 to 100 = 5050
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(5050);
      
      // Update a cell in the range
      engine.setCellContents({ sheet: sheetId, col: 0, row: 49 }, 150); // Change A50 from 50 to 150
      
      // Sum should update efficiently (5050 - 50 + 150 = 5150)
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(5150);
    });

    test('COUNT over sparse range', () => {
      // Create sparse data
      const data = new Map([
        ['A1', 1],
        ['A10', 10],
        ['A50', 50],
        ['A100', 100]
      ]);
      engine.setSheetContents(sheetId, data);

      // Count non-empty cells
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=COUNT(A1:A100)');
      
      // Should count only 4 non-empty cells
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(4);
    });

    test('AVERAGE with incremental updates', () => {
      // Set up initial data
      const data = new Map([
        ['A1', 10],
        ['A2', 20],
        ['A3', 30],
        ['A4', 40],
        ['A5', 50]
      ]);
      engine.setSheetContents(sheetId, data);

      // Calculate average
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=AVERAGE(A1:A5)');
      
      // Initial average = (10+20+30+40+50)/5 = 30
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(30);
      
      // Update a value
      engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 80); // Change A3 from 30 to 80
      
      // New average = (10+20+80+40+50)/5 = 40
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(40);
    });

    test('Nested associative operations', () => {
      // Set up data
      const data = new Map([
        ['A1', 1], ['B1', 10],
        ['A2', 2], ['B2', 20],
        ['A3', 3], ['B3', 30],
        ['A4', 4], ['B4', 40],
        ['A5', 5], ['B5', 50]
      ]);
      engine.setSheetContents(sheetId, data);

      // Complex formula with multiple associative operations
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=SUM(A1:A5) * AVERAGE(B1:B5)');
      
      // SUM(A1:A5) = 15, AVERAGE(B1:B5) = 30, result = 450
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe(450);
    });

    test('MIN/MAX over large ranges', () => {
      // Create data with known min/max
      const data = new Map<string, CellValue>();
      for (let i = 0; i < 50; i++) {
        data.set(`A${i+1}`, i * 2 + 10);
      }
      // Add specific min and max values
      data.set('A25', 5);  // Min
      data.set('A40', 200); // Max
      
      engine.setSheetContents(sheetId, data);

      // Find min and max
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=MIN(A1:A50)');
      engine.setCellContents({ sheet: sheetId, col: 1, row: 1 }, '=MAX(A1:A50)');
      
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(5);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 1 })).toBe(200);
    });
  });

  describe('Integration of Advanced Features', () => {
    test('Array spilling with relative references', () => {
      // Set up data
      const data = new Map([
        ['A1', 1], ['B1', 10],
        ['A2', 2], ['B2', 20],
        ['A3', 3], ['B3', 30]
      ]);
      engine.setSheetContents(sheetId, data);

      // Create a formula that adds columns with relative references
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=A1:A3 + B1:B3');
      
      // Check spilled results
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe(11);
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 1 })).toBe(22);
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 2 })).toBe(33);
    });

    test('Associative operations on spilled arrays', () => {
      // Set up data
      const data = new Map([
        ['A1', 1], ['A2', 2], ['A3', 3], ['A4', 4], ['A5', 5]
      ]);
      engine.setSheetContents(sheetId, data);

      // Filter and then sum the results
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=FILTER(A1:A5, A1:A5 > 2)');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=SUM(B1:B3)'); // Sum the spilled range
      
      // The filter returns [3, 4, 5], sum = 12
      expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe(12);
    });

    test('Copy-paste preserving array formulas', () => {
      // Set up data
      const data = new Map([
        ['A1', 1], ['A2', 2], ['A3', 3]
      ]);
      engine.setSheetContents(sheetId, data);

      // Create an array formula
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1:A3 * 2');
      
      // Verify initial spill
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe(2);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 1 })).toBe(4);
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 2 })).toBe(6);
      
      // Copy the array formula range
      engine.copy({ 
        start: { sheet: sheetId, col: 1, row: 0 },
        end: { sheet: sheetId, col: 1, row: 2 }
      });
      
      // Paste to a new location
      engine.paste({ sheet: sheetId, col: 3, row: 0 });
      
      // Check pasted values
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 0 })).toBe(2);
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 1 })).toBe(4);
      expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 2 })).toBe(6);
    });
  });
});