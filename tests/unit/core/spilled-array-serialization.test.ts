import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../../src/core/engine';

describe('Spilled Array Serialization', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    sheetId = engine.getSheetId(sheetName);
  });

  test('should serialize spilled cells as empty (Excel compatibility)', () => {
    // Set up data that will create a spilled array
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
    engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
    engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 3);

    // Create a FILTER formula that will spill into multiple cells
    engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=FILTER(A1:A3, A1:A3>1)');

    // Verify the array spilled correctly (values are visible)
    expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe(2); // Origin cell
    expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 1 })).toBe(3); // Spilled cell
    
    // Verify array cell types
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 2, row: 0 })).toBe(true);
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 2, row: 1 })).toBe(true);

    // Get serialized content
    const serialized = engine.getSheetSerialized(sheetId);
    
    // Origin cell should contain the formula
    expect(serialized.get('C1')).toBe('=FILTER(A1:A3, A1:A3>1)');
    
    // Spilled cell should NOT be present in serialized output (Excel compatibility)
    expect(serialized.has('C2')).toBe(false);
    expect(serialized.get('C2')).toBeUndefined();
    
    // But the formula should be empty for the spilled cell
    expect(engine.getCellFormula({ sheet: sheetId, col: 2, row: 1 })).toBe('');
  });

  test('should serialize individual cell formulas correctly in presence of spilled arrays', () => {
    // Set up data
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
    engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
    
    // Create a spilled array
    engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=FILTER(A1:A2, A1:A2>0)');
    
    // Add a regular formula in a different location
    engine.setCellContents({ sheet: sheetId, col: 4, row: 0 }, '=SUM(A1:A2)');
    
    const serialized = engine.getSheetSerialized(sheetId);
    
    // Array origin should serialize with formula
    expect(serialized.get('C1')).toBe('=FILTER(A1:A2, A1:A2>0)');
    
    // Array spilled cell should not be in serialized output
    expect(serialized.has('C2')).toBe(false);
    
    // Regular formula should serialize normally
    expect(serialized.get('E1')).toBe('=SUM(A1:A2)');
    
    // Source data should serialize as values
    expect(serialized.get('A1')).toBe(1);
    expect(serialized.get('A2')).toBe(2);
  });

  test('should handle 2D spilled arrays correctly', () => {
    // Create a 2x2 array that will spill
    const data = new Map([
      ['A1', 1], ['B1', 2],
      ['A2', 3], ['B2', 4]
    ]);
    engine.setSheetContents(sheetId, data);
    
    // Use ARRAY_CONSTRAIN to create a 2x2 spilled array
    engine.setCellContents({ sheet: sheetId, col: 3, row: 0 }, '=ARRAY_CONSTRAIN(A1:B2, 2, 2)');
    
    // Verify values are visible
    expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 0 })).toBe(1); // D1 (origin)
    expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 0 })).toBe(2); // E1 (spilled)
    expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 1 })).toBe(3); // D2 (spilled)
    expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 1 })).toBe(4); // E2 (spilled)
    
    const serialized = engine.getSheetSerialized(sheetId);
    
    // Only origin cell should be in serialized output
    expect(serialized.get('D1')).toBe('=ARRAY_CONSTRAIN(A1:B2, 2, 2)');
    
    // All spilled cells should be absent from serialized output
    expect(serialized.has('E1')).toBe(false);
    expect(serialized.has('D2')).toBe(false);
    expect(serialized.has('E2')).toBe(false);
  });

  test('should serialize getCellSerialized correctly for individual spilled cells', () => {
    // Set up spilled array
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
    engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
    engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=FILTER(A1:A2, A1:A2>0)');
    
    // Test individual cell serialization
    const originSerialized = engine.getCellSerialized({ sheet: sheetId, col: 2, row: 0 });
    const spilledSerialized = engine.getCellSerialized({ sheet: sheetId, col: 2, row: 1 });
    
    expect(originSerialized).toBe('=FILTER(A1:A2, A1:A2>0)');
    expect(spilledSerialized).toBeUndefined(); // Spilled cell should serialize as undefined
  });
});
