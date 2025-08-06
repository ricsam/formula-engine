import { describe, test, expect, beforeEach } from 'bun:test';
import { FormulaEngine } from '../../src/core/engine';
import type { SimpleCellAddress, SimpleCellRange } from '../../src/core/types';

describe('Dependency Tracking Integration Tests', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    sheetId = engine.getSheetId(sheetName);
  });

  describe('getCellPrecedents', () => {
    test('should track single cell reference', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1');
      
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 1, row: 0 });
      expect(precedents).toHaveLength(1);
      expect(precedents[0]).toEqual({ sheet: sheetId, col: 0, row: 0 });
    });

    test('should track multiple cell references', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, 20);
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=A1+B1');
      
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 2, row: 0 });
      expect(precedents).toHaveLength(2);
      expect(precedents).toContainEqual({ sheet: sheetId, col: 0, row: 0 });
      expect(precedents).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
    });

    test('should track range references', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 3);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=SUM(A1:A3)');
      
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 1, row: 0 });
      expect(precedents).toHaveLength(1);
      expect(precedents[0]).toEqual({
        start: { sheet: sheetId, col: 0, row: 0 },
        end: { sheet: sheetId, col: 0, row: 2 }
      });
    });

    test('should track indirect references through named expressions', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 100);
      engine.addNamedExpression('MyValue', '=A1');
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=MyValue * 2');
      
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 1, row: 0 });
      
      // Should include A1 as a transitive dependency through the named expression
      expect(precedents.length).toBeGreaterThanOrEqual(1);
      
      // Verify that A1 is included in the precedents (transitive dependency)
      const hasA1 = precedents.some(prec => {
        return 'col' in prec && 'row' in prec && 'sheet' in prec &&
               prec.col === 0 && prec.row === 0 && prec.sheet === sheetId;
      });
      expect(hasA1).toBe(true);
      
      // Also verify the formula evaluation works correctly
      const result = engine.getCellValue({ sheet: sheetId, col: 1, row: 0 });
      expect(result).toBe(200);
    });

    test('should return empty array for cells without formulas', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 42);
      
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 0, row: 0 });
      expect(precedents).toEqual([]);
    });

    test('should track precedents in complex formulas', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, 3);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, '=IF(A1>0, B1, C1)');
      
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 0, row: 1 });
      expect(precedents).toHaveLength(3);
      expect(precedents).toContainEqual({ sheet: sheetId, col: 0, row: 0 });
      expect(precedents).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(precedents).toContainEqual({ sheet: sheetId, col: 2, row: 0 });
    });
  });

  describe('getCellDependents', () => {
    test('should track single dependent', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1');
      
      const dependents = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependents).toHaveLength(1);
      expect(dependents[0]).toEqual({ sheet: sheetId, col: 1, row: 0 });
    });

    test('should track multiple dependents', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1*2');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=A1+5');
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, '=A1/2');
      
      const dependents = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependents).toHaveLength(3);
      expect(dependents).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(dependents).toContainEqual({ sheet: sheetId, col: 2, row: 0 });
      expect(dependents).toContainEqual({ sheet: sheetId, col: 0, row: 1 });
    });

    test('should track dependents through ranges', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 3);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=SUM(A1:A3)');
      
      // Each cell in the range should show B1 as a dependent
      const dependentsA1 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      const dependentsA2 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 1 });
      const dependentsA3 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 2 });
      
      expect(dependentsA1).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(dependentsA2).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(dependentsA3).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
    });

    test('should return empty array for cells with no dependents', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 42);
      
      const dependents = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependents).toEqual([]);
    });

    test('should track transitive dependents', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1*2');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=B1+5');
      
      // A1 should show B1 as a direct dependent
      const directDependents = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(directDependents).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      
      // B1 should show C1 as a dependent
      const indirectDependents = engine.getCellDependents({ sheet: sheetId, col: 1, row: 0 });
      expect(indirectDependents).toContainEqual({ sheet: sheetId, col: 2, row: 0 });
    });
  });

  describe('Range precedents and dependents', () => {
    test('should get precedents for a range', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1');
      engine.setCellContents({ sheet: sheetId, col: 1, row: 1 }, '=A2');
      
      const range: SimpleCellRange = {
        start: { sheet: sheetId, col: 1, row: 0 },
        end: { sheet: sheetId, col: 1, row: 1 }
      };
      
      const precedents = engine.getCellPrecedents(range);
      expect(precedents).toHaveLength(2);
      expect(precedents).toContainEqual({ sheet: sheetId, col: 0, row: 0 });
      expect(precedents).toContainEqual({ sheet: sheetId, col: 0, row: 1 });
    });

    test('should get dependents for a range', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1:A2');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=SUM(A1:A2)');
      
      const range: SimpleCellRange = {
        start: { sheet: sheetId, col: 0, row: 0 },
        end: { sheet: sheetId, col: 0, row: 1 }
      };
      
      const dependents = engine.getCellDependents(range);
      expect(dependents.length).toBeGreaterThanOrEqual(1);
      // Should include cells that depend on any cell in the range
    });
  });

  describe('Dependency updates', () => {
    test('should update dependencies when formula changes', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, 20);
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=A1');
      
      // Initially C1 depends on A1
      let dependentsA1 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependentsA1).toContainEqual({ sheet: sheetId, col: 2, row: 0 });
      
      // Change C1 to depend on B1 instead
      engine.setCellContents({ sheet: sheetId, col: 2, row: 0 }, '=B1');
      
      // Now A1 should have no dependents
      dependentsA1 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependentsA1).toEqual([]);
      
      // And B1 should have C1 as a dependent
      const dependentsB1 = engine.getCellDependents({ sheet: sheetId, col: 1, row: 0 });
      expect(dependentsB1).toContainEqual({ sheet: sheetId, col: 2, row: 0 });
    });

    test('should remove dependencies when cell is cleared', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 10);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1');
      
      // Initially A1 has B1 as a dependent
      let dependents = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependents).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      
      // Clear B1
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '');
      
      // Now A1 should have no dependents
      dependents = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      expect(dependents).toEqual([]);
    });
  });

  describe('Circular dependencies', () => {
    test('should handle circular dependencies gracefully', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, '=B1');
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1');
      
      // Both cells should show the circular reference error
      expect(engine.getCellValue({ sheet: sheetId, col: 0, row: 0 })).toBe('#CYCLE!');
      expect(engine.getCellValue({ sheet: sheetId, col: 1, row: 0 })).toBe('#CYCLE!');
      
      // Dependencies should still be tracked
      const precedentsA1 = engine.getCellPrecedents({ sheet: sheetId, col: 0, row: 0 });
      const precedentsB1 = engine.getCellPrecedents({ sheet: sheetId, col: 1, row: 0 });
      
      expect(precedentsA1).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(precedentsB1).toContainEqual({ sheet: sheetId, col: 0, row: 0 });
    });
  });

  describe('Array formula dependencies', () => {
    test('should track dependencies for array formulas', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 3);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1:A3 * 2');
      
      // The array formula should depend on A1:A3
      const precedents = engine.getCellPrecedents({ sheet: sheetId, col: 1, row: 0 });
      expect(precedents).toContainEqual({
        start: { sheet: sheetId, col: 0, row: 0 },
        end: { sheet: sheetId, col: 0, row: 2 }
      });
      
      // Each cell in A1:A3 should show B1 as a dependent (the array origin)
      const dependentsA1 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 0 });
      const dependentsA2 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 1 });
      const dependentsA3 = engine.getCellDependents({ sheet: sheetId, col: 0, row: 2 });
      
      expect(dependentsA1).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(dependentsA2).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
      expect(dependentsA3).toContainEqual({ sheet: sheetId, col: 1, row: 0 });
    });

    test('should track dependencies for spilled cells', () => {
      engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 1);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 1 }, 2);
      engine.setCellContents({ sheet: sheetId, col: 0, row: 2 }, 3);
      engine.setCellContents({ sheet: sheetId, col: 1, row: 0 }, '=A1:A3 * 2');
      engine.setCellContents({ sheet: sheetId, col: 2, row: 1 }, '=B2 + 10');
      
      // C2 depends on B2 (which is a spilled cell)
      const precedentsC2 = engine.getCellPrecedents({ sheet: sheetId, col: 2, row: 1 });
      expect(precedentsC2).toContainEqual({ sheet: sheetId, col: 1, row: 1 });
      
      // B2 (spilled cell) should show C2 as a dependent
      const dependentsB2 = engine.getCellDependents({ sheet: sheetId, col: 1, row: 1 });
      expect(dependentsB2).toContainEqual({ sheet: sheetId, col: 2, row: 1 });
    });
  });
});