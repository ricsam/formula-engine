import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../src/core/engine';

describe('FILTER with Infinite Ranges - Spilling Bug', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    sheetId = engine.getSheetId(sheetName);
  });

  test('should spill FILTER results correctly with infinite ranges in high columns', () => {
    // Set up data exactly as user described - fruit on col P, flag on col Q, filter on col R
    const data = new Map<string, any>([
      // Column P (15) - Fruit
      ['P1', 'Apple'],
      ['P2', 'Banana'], 
      ['P3', 'Cherry'],
      ['P4', 'Date'],
      ['P5', 'Elderberry'],
      // Column Q (16) - Flag
      ['Q1', 'Yes'],
      ['Q2', 'No'],
      ['Q3', 'Yes'],
      ['Q4', 'Yes'],
      ['Q5', 'No'],
      // Column R (17) - Filter formula
      ['R1', '=FILTER(P:P, Q:Q="Yes")']
    ]);

    engine.setSheetContent(sheetId, data);

    // Verify input data
    expect(engine.getCellValue({ sheet: sheetId, col: 15, row: 0 })).toBe('Apple');
    expect(engine.getCellValue({ sheet: sheetId, col: 15, row: 2 })).toBe('Cherry');
    expect(engine.getCellValue({ sheet: sheetId, col: 15, row: 3 })).toBe('Date');
    
    expect(engine.getCellValue({ sheet: sheetId, col: 16, row: 0 })).toBe('Yes');
    expect(engine.getCellValue({ sheet: sheetId, col: 16, row: 2 })).toBe('Yes');
    expect(engine.getCellValue({ sheet: sheetId, col: 16, row: 3 })).toBe('Yes');

    // Check formula is set correctly
    expect(engine.getCellFormula({ sheet: sheetId, col: 17, row: 0 })).toBe('=FILTER(P:P, Q:Q="Yes")');

    // Verify spilled array results
    expect(engine.getCellValue({ sheet: sheetId, col: 17, row: 0 })).toBe('Apple');  // R1 - first match
    expect(engine.getCellValue({ sheet: sheetId, col: 17, row: 1 })).toBe('Cherry'); // R2 - second match  
    expect(engine.getCellValue({ sheet: sheetId, col: 17, row: 2 })).toBe('Date');   // R3 - third match

    // Verify array cell status
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 17, row: 0 })).toBe(true);
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 17, row: 1 })).toBe(true);
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 17, row: 2 })).toBe(true);
  });

  test('should also work with lower column letters for comparison', () => {
    // Test the same thing but with columns A, B, C to see if it's column-specific
    const data = new Map<string, any>([
      ['A1', 'Apple'],
      ['A2', 'Banana'], 
      ['A3', 'Cherry'],
      ['A4', 'Date'],
      ['A5', 'Elderberry'],
      ['B1', 'Yes'],
      ['B2', 'No'],
      ['B3', 'Yes'],
      ['B4', 'Yes'],
      ['B5', 'No'],
      ['C1', '=FILTER(A:A, B:B="Yes")']
    ]);

    engine.setSheetContent(sheetId, data);

    // This should work correctly
    expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 0 })).toBe('Apple');
    expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 1 })).toBe('Cherry');
    expect(engine.getCellValue({ sheet: sheetId, col: 2, row: 2 })).toBe('Date');

    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 2, row: 0 })).toBe(true);
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 2, row: 1 })).toBe(true);
    expect(engine.isCellPartOfArray({ sheet: sheetId, col: 2, row: 2 })).toBe(true);
  });
});
