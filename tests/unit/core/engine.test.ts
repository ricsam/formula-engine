import { test, expect, describe } from "bun:test";
import { FormulaEngine } from '../../../src/core/engine';

describe('FormulaEngine', () => {
  test('should create an empty engine', () => {
    const engine = FormulaEngine.buildEmpty();
    expect(engine).toBeDefined();
    expect(engine.countSheets()).toBe(0);
  });

  test('should add a sheet', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    
    expect(sheetName).toBe('TestSheet');
    expect(engine.countSheets()).toBe(1);
    expect(engine.doesSheetExist('TestSheet')).toBe(true);
    expect(engine.getSheetId('TestSheet')).toBe(0);
  });

  test('should set and get cell values', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    // Set a single value
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 42);
    
    const value = engine.getCellValue({ sheet: sheetId, col: 0, row: 0 });
    expect(value).toBe(42);
  });

  test('should handle empty cells', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    const value = engine.getCellValue({ sheet: sheetId, col: 5, row: 5 });
    expect(value).toBeUndefined();
    expect(engine.isCellEmpty({ sheet: sheetId, col: 5, row: 5 })).toBe(true);
  });

  test('should set multiple values with 2D array', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    const data = [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9]
    ];
    
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, data);
    
    // Verify all values
    for (let row = 0; row < 3; row++) {
      for (let col = 0; col < 3; col++) {
        const value = engine.getCellValue({ sheet: sheetId, col, row });
        const expectedValue = data[row]![col]!; // We know these exist from our test data
        expect(value).toBe(expectedValue);
      }
    }
  });

  test('should get range values', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    // Set some values
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, [
      [1, 2, 3],
      [4, 5, 6]
    ]);
    
    const range = {
      start: { sheet: sheetId, col: 0, row: 0 },
      end: { sheet: sheetId, col: 2, row: 1 }
    };
    
    const values = engine.getRangeValues(range);
    expect(values).toEqual([
      [1, 2, 3],
      [4, 5, 6]
    ]);
  });

  test('should handle formulas (stored as strings for now)', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, '=A2+B2');
    
    const formula = engine.getCellFormula({ sheet: sheetId, col: 0, row: 0 });
    expect(formula).toBe('A2+B2');
    
    const serialized = engine.getCellSerialized({ sheet: sheetId, col: 0, row: 0 });
    expect(serialized).toBe('=A2+B2');
  });

  test('should parse cell addresses', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    const addr1 = engine.simpleCellAddressFromString('A1', sheetId);
    expect(addr1).toEqual({ sheet: sheetId, col: 0, row: 0 });
    
    const addr2 = engine.simpleCellAddressFromString('B3', sheetId);
    expect(addr2).toEqual({ sheet: sheetId, col: 1, row: 2 });
    
    const addr3 = engine.simpleCellAddressFromString('AA10', sheetId);
    expect(addr3).toEqual({ sheet: sheetId, col: 26, row: 9 });
  });

  test('should parse cell ranges', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    const range = engine.simpleCellRangeFromString('A1:C3', sheetId);
    expect(range).toEqual({
      start: { sheet: sheetId, col: 0, row: 0 },
      end: { sheet: sheetId, col: 2, row: 2 }
    });
  });

  test('should manage sheets', () => {
    const engine = FormulaEngine.buildEmpty();
    
    // Add multiple sheets
    engine.addSheet('Sheet1');
    engine.addSheet('Sheet2');
    engine.addSheet('Sheet3');
    
    expect(engine.countSheets()).toBe(3);
    expect(engine.getSheetNames()).toEqual(['Sheet1', 'Sheet2', 'Sheet3']);
    
    // Rename a sheet
    const sheet2Id = engine.getSheetId('Sheet2');
    engine.renameSheet(sheet2Id, 'DataSheet');
    expect(engine.getSheetName(sheet2Id)).toBe('DataSheet');
    
    // Remove a sheet
    engine.removeSheet(sheet2Id);
    expect(engine.countSheets()).toBe(2);
    expect(engine.doesSheetExist('DataSheet')).toBe(false);
  });

  test('should get bounding rectangle', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    // Empty sheet has no bounding rect
    let bounds = engine.getSheetBoundingRect(sheetId);
    expect(bounds).toBeUndefined();
    
    // Add some data
    engine.setCellContents({ sheet: sheetId, col: 1, row: 2 }, 'A');
    engine.setCellContents({ sheet: sheetId, col: 5, row: 8 }, 'B');
    
    bounds = engine.getSheetBoundingRect(sheetId);
    expect(bounds).toEqual({
      minCol: 1,
      maxCol: 5,
      minRow: 2,
      maxRow: 8,
      width: 5,
      height: 7
    });
  });

  test('should handle named expressions', () => {
    const engine = FormulaEngine.buildEmpty();
    
    // Add global named expression
    engine.addNamedExpression('PI', 3.14159);
    expect(engine.getNamedExpressionFormula('PI')).toBe('3.14159');
    
    // List named expressions
    const names = engine.listNamedExpressions();
    expect(names).toContain('PI');
    
    // Change named expression
    engine.changeNamedExpression('PI', 3.14159265);
    expect(engine.getNamedExpressionFormula('PI')).toBe('3.14159265');
    
    // Remove named expression
    engine.removeNamedExpression('PI');
    expect(engine.listNamedExpressions()).not.toContain('PI');
  });

  test('should handle copy and paste', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    // Set source data
    engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, [
      [1, 2],
      [3, 4]
    ]);
    
    // Copy
    const range = {
      start: { sheet: sheetId, col: 0, row: 0 },
      end: { sheet: sheetId, col: 1, row: 1 }
    };
    const copied = engine.copy(range);
    expect(copied).toEqual([[1, 2], [3, 4]]);
    expect(engine.isClipboardEmpty()).toBe(false);
    
    // Paste
    engine.paste({ sheet: sheetId, col: 3, row: 3 });
    
    // Verify pasted data
    expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 3 })).toBe(1);
    expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 3 })).toBe(2);
    expect(engine.getCellValue({ sheet: sheetId, col: 3, row: 4 })).toBe(3);
    expect(engine.getCellValue({ sheet: sheetId, col: 4, row: 4 })).toBe(4);
  });

  test('should suspend and resume evaluation', () => {
    const engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('Sheet1');
    const sheetId = engine.getSheetId(sheetName);
    
    // Suspend evaluation
    engine.suspendEvaluation();
    expect(engine.isEvaluationSuspended()).toBe(true);
    
    // Changes should not return anything while suspended
    const changes1 = engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 42);
    expect(changes1).toEqual([]);
    
    // Resume evaluation
    const changes2 = engine.resumeEvaluation();
    expect(engine.isEvaluationSuspended()).toBe(false);
    // In future, this would return the pending changes
  });
});