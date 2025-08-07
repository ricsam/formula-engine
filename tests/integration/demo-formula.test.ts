import { test, expect, describe } from "bun:test";
import { FormulaEngine } from '../../src/core/engine';

describe('Demo Formula Integration Test', () => {
  test('should handle the exact formula from the demo: =IF(COUNTIF(A2:A5,"Laptop")>0,"Yes","No")', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    const sheetId = engine.getSheetId(sheetName);

    // Set up data exactly like in the demo
    const testData = new Map<string, any>([
      ['A2', 'Laptop'],
      ['A3', 'Mouse'], 
      ['A4', 'Keyboard'],
      ['A5', 'Monitor']
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    // Test the exact formula that was failing
    const address = { sheet: sheetId, col: 1, row: 13 }; // B14 position from demo
    engine.setCellContents(address, '=IF(COUNTIF(A2:A5,"Laptop")>0,"Yes","No")');
    
    const result = engine.getCellValue(address);
    expect(result).toBe('Yes');
  });

  test('should handle COUNTIF with case-insensitive matching', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    const sheetId = engine.getSheetId(sheetName);

    const testData = new Map<string, any>([
      ['A1', 'laptop'],
      ['A2', 'LAPTOP'],
      ['A3', 'Laptop'],
      ['A4', 'mouse']
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    const address = { sheet: sheetId, col: 1, row: 0 };
    engine.setCellContents(address, '=COUNTIF(A1:A4,"Laptop")');
    
    const result = engine.getCellValue(address);
    expect(result).toBe(3); // Should match all 3 variations of "laptop"
  });

  test('should work with SUMIF function', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    const sheetId = engine.getSheetId(sheetName);

    const testData = new Map<string, any>([
      ['A1', 'Laptop'],
      ['A2', 'Mouse'],
      ['A3', 'Laptop'],
      ['A4', 'Keyboard'],
      ['B1', 1000],
      ['B2', 50],
      ['B3', 1200],
      ['B4', 100]
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    // Sum prices where product is "Laptop"
    const address = { sheet: sheetId, col: 2, row: 0 };
    engine.setCellContents(address, '=SUMIF(A1:A4,"Laptop",B1:B4)');
    
    const result = engine.getCellValue(address);
    expect(result).toBe(2200); // 1000 + 1200
  });
});