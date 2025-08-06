import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../../../src/core/engine';

describe('COUNTIF Function Tests', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    sheetId = engine.getSheetId(sheetName);
  });

  test('should work with COUNTIF function', () => {
    // Set up data similar to the demo
    const testData = new Map([
      ['A2', 'Laptop'],
      ['A3', 'Mouse'], 
      ['A4', 'Keyboard'],
      ['A5', 'Monitor']
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    // This should now work because COUNTIF is implemented
    const address = { sheet: sheetId, col: 1, row: 1 };
    engine.setCellContents(address, '=COUNTIF(A2:A5,"Laptop")');
    
    const result = engine.getCellValue(address);
    expect(result).toBe(1);
  });

  test('should support nested COUNTIF in IF function', () => {
    // Set up data similar to the demo
    const testData = new Map([
      ['A2', 'Laptop'],
      ['A3', 'Mouse'], 
      ['A4', 'Keyboard'],
      ['A5', 'Monitor']
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    // This should now work because COUNTIF is implemented
    const address = { sheet: sheetId, col: 1, row: 1 };
    engine.setCellContents(address, '=IF(COUNTIF(A2:A5,"Laptop")>0,"Yes","No")');
    
    const result = engine.getCellValue(address);
    expect(result).toBe('Yes');
  });

  test('COUNTIF basic functionality', () => {
    const testData = new Map([
      ['A1', 1],
      ['A2', 2],
      ['A3', 1],
      ['A4', 3],
      ['A5', 1]
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    const address = { sheet: sheetId, col: 1, row: 0 };
    engine.setCellContents(address, '=COUNTIF(A1:A5,1)');
    
    const result = engine.getCellValue(address);
    expect(result).toBe(3);
  });

  test('COUNTIF with text criteria', () => {
    const testData = new Map([
      ['A1', 'Apple'],
      ['A2', 'Banana'],
      ['A3', 'Apple'],
      ['A4', 'Orange'],
      ['A5', 'Apple']
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    const address = { sheet: sheetId, col: 1, row: 0 };
    engine.setCellContents(address, '=COUNTIF(A1:A5,"Apple")');
    
    const result = engine.getCellValue(address);
    expect(result).toBe(3);
  });

  test('COUNTIF with comparison operators', () => {
    const testData = new Map([
      ['A1', 10],
      ['A2', 20],
      ['A3', 15],
      ['A4', 25],
      ['A5', 5]
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    // Test greater than
    const address1 = { sheet: sheetId, col: 1, row: 0 };
    engine.setCellContents(address1, '=COUNTIF(A1:A5,">15")');
    const result1 = engine.getCellValue(address1);
    expect(result1).toBe(2); // 20 and 25

    // Test less than or equal
    const address2 = { sheet: sheetId, col: 1, row: 1 };
    engine.setCellContents(address2, '=COUNTIF(A1:A5,"<=15")');
    const result2 = engine.getCellValue(address2);
    expect(result2).toBe(3); // 10, 15, and 5
  });

  test('COUNTIF edge cases', () => {
    const testData = new Map([
      ['A1', 0],
      ['A2', ''],
      ['A3', false],
      ['A4', true],
      ['A5', '#DIV/0!']
    ]);
    
    engine.setSheetContents(sheetId, testData);
    
    // Count empty cells
    const address1 = { sheet: sheetId, col: 1, row: 0 };
    engine.setCellContents(address1, '=COUNTIF(A1:A5,"")');
    const result1 = engine.getCellValue(address1);
    expect(result1).toBe(1); // Only A2 is empty string

    // Count boolean values
    const address2 = { sheet: sheetId, col: 1, row: 1 };
    engine.setCellContents(address2, '=COUNTIF(A1:A5,FALSE)');
    const result2 = engine.getCellValue(address2);
    expect(result2).toBe(1); // Only A3 is false
  });
});