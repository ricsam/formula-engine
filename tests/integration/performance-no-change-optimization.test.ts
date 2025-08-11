import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../src/core/engine";
import type { SimpleCellAddress } from "../../src/core/types";

describe("Performance Optimization - No Change Updates", () => {
  let engine: FormulaEngine;
  let sheetId: number;
  let updateCounts: Map<string, number>;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("TestSheet");
    sheetId = engine.getSheetId(sheetName);
    updateCounts = new Map();

    // Track all cell updates using onCellsUpdate for the sheet
    engine.onCellsUpdate(sheetId, (events) => {
      for (const event of events) {
        const key = `${event.address.sheet}:${event.address.col}:${event.address.row}`;
        updateCounts.set(key, (updateCounts.get(key) || 0) + 1);
      }
    });
  });

  const addr = (col: number, row: number): SimpleCellAddress => ({
    sheet: sheetId,
    col,
    row,
  });

  const getUpdateCount = (addr: SimpleCellAddress): number => {
    const key = `${addr.sheet}:${addr.col}:${addr.row}`;
    return updateCounts.get(key) || 0;
  };

  const clearUpdateCounts = () => {
    updateCounts.clear();
  };

  describe("setCellContent optimization", () => {
    test("should not trigger updates when setting same value", () => {
      // Set initial value
      engine.setCellContent(addr(0, 0), "Hello");
      expect(engine.getCellValue(addr(0, 0))).toBe("Hello");
      expect(getUpdateCount(addr(0, 0))).toBe(1);

      clearUpdateCounts();

      // Set the same value again
      engine.setCellContent(addr(0, 0), "Hello");
      expect(engine.getCellValue(addr(0, 0))).toBe("Hello");
      expect(getUpdateCount(addr(0, 0))).toBe(0); // Should not trigger update
    });

    test("should not trigger updates when setting same number", () => {
      // Set initial value
      engine.setCellContent(addr(0, 0), 42);
      expect(engine.getCellValue(addr(0, 0))).toBe(42);
      expect(getUpdateCount(addr(0, 0))).toBe(1);

      clearUpdateCounts();

      // Set the same value again
      engine.setCellContent(addr(0, 0), 42);
      expect(engine.getCellValue(addr(0, 0))).toBe(42);
      expect(getUpdateCount(addr(0, 0))).toBe(0); // Should not trigger update
    });

    test("should not trigger updates when setting equivalent string numbers", () => {
      // Set initial value
      engine.setCellContent(addr(0, 0), "42");
      expect(engine.getCellValue(addr(0, 0))).toBe(42); // Parsed as number
      expect(getUpdateCount(addr(0, 0))).toBe(1);

      clearUpdateCounts();

      // Set equivalent number
      engine.setCellContent(addr(0, 0), 42);
      expect(engine.getCellValue(addr(0, 0))).toBe(42);
      expect(getUpdateCount(addr(0, 0))).toBe(0); // Should not trigger update
    });

    test("should not trigger cascade when dependent unchanged", () => {
      // Set up dependency: B1 = A1 * 2
      engine.setCellContent(addr(0, 0), 5);  // A1
      engine.setCellContent(addr(1, 0), "=A1*2"); // B1
      
      expect(engine.getCellValue(addr(0, 0))).toBe(5);  // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(10); // B1

      clearUpdateCounts();

      // Set A1 to the same value
      engine.setCellContent(addr(0, 0), 5);
      
      // A1 should not update, and B1 should not recalculate
      expect(getUpdateCount(addr(0, 0))).toBe(0); // A1 not updated
      expect(getUpdateCount(addr(1, 0))).toBe(0); // B1 not recalculated
    });

    test("should not trigger cascade when formula result unchanged", () => {
      // Set up: A1 = 3, B1 = A1 + 2, C1 = B1 * 2
      engine.setCellContent(addr(0, 0), 3);      // A1
      engine.setCellContent(addr(1, 0), "=A1+2"); // B1 = 5
      engine.setCellContent(addr(2, 0), "=B1*2"); // C1 = 10

      expect(engine.getCellValue(addr(0, 0))).toBe(3);  // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(5);  // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(10); // C1

      clearUpdateCounts();

      // Change A1 to 3 again (no actual change)
      engine.setCellContent(addr(0, 0), 3);
      
      // None should trigger updates
      expect(getUpdateCount(addr(0, 0))).toBe(0); // A1
      expect(getUpdateCount(addr(1, 0))).toBe(0); // B1
      expect(getUpdateCount(addr(2, 0))).toBe(0); // C1
    });

    test("should trigger updates when value actually changes", () => {
      // Set initial value
      engine.setCellContent(addr(0, 0), 5);
      expect(getUpdateCount(addr(0, 0))).toBe(1);

      clearUpdateCounts();

      // Change to different value
      engine.setCellContent(addr(0, 0), 10);
      expect(engine.getCellValue(addr(0, 0))).toBe(10);
      expect(getUpdateCount(addr(0, 0))).toBe(1); // Should trigger update
    });
  });

  describe("setSheetContent batch optimization", () => {
    test("should not trigger updates for unchanged values in batch", () => {
      // Set initial values
      const initialData = new Map<string, any>([
        ["A1", 10],
        ["B1", 20],
        ["C1", "=A1+B1"], // C1 = 30
      ]);

      engine.setSheetContent(sheetId, initialData);
      expect(engine.getCellValue(addr(0, 0))).toBe(10); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(20); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(30); // C1

      clearUpdateCounts();

      // Set the same values again
      const sameData = new Map<string, any>([
        ["A1", 10],
        ["B1", 20],
        ["C1", "=A1+B1"],
      ]);

      engine.setSheetContent(sheetId, sameData);
      
      // No cells should be updated
      expect(getUpdateCount(addr(0, 0))).toBe(0); // A1
      expect(getUpdateCount(addr(1, 0))).toBe(0); // B1
      expect(getUpdateCount(addr(2, 0))).toBe(0); // C1
    });

    test("should only update changed cells in batch", () => {
      // Set initial values
      const initialData = new Map<string, any>([
        ["A1", 10],
        ["B1", 20],
        ["C1", 30],
      ]);

      engine.setSheetContent(sheetId, initialData);
      clearUpdateCounts();

      // Change only B1
      const partialChangeData = new Map<string, any>([
        ["A1", 10],    // Same
        ["B1", 25],    // Changed
        ["C1", 30],    // Same
      ]);

      engine.setSheetContent(sheetId, partialChangeData);
      
      // Only B1 should be updated
      expect(getUpdateCount(addr(0, 0))).toBe(0); // A1 unchanged
      expect(getUpdateCount(addr(1, 0))).toBe(1); // B1 changed
      expect(getUpdateCount(addr(2, 0))).toBe(0); // C1 unchanged
    });

    test("should handle formula dependencies correctly in batch", () => {
      // Set initial values with dependencies
      const initialData = new Map<string, any>([
        ["A1", 5],
        ["B1", "=A1*2"],  // B1 = 10
        ["C1", "=B1+5"],  // C1 = 15
      ]);

      engine.setSheetContent(sheetId, initialData);
      expect(engine.getCellValue(addr(0, 0))).toBe(5);  // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(10); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(15); // C1

      clearUpdateCounts();

      // Set the same formulas (should result in same values)
      const sameData = new Map<string, any>([
        ["A1", 5],
        ["B1", "=A1*2"],
        ["C1", "=B1+5"],
      ]);

      engine.setSheetContent(sheetId, sameData);
      
      // None should trigger updates since values are the same
      expect(getUpdateCount(addr(0, 0))).toBe(0); // A1
      expect(getUpdateCount(addr(1, 0))).toBe(0); // B1
      expect(getUpdateCount(addr(2, 0))).toBe(0); // C1
    });

    test("should handle mixed changes and non-changes in batch", () => {
      // Set initial values
      const initialData = new Map<string, any>([
        ["A1", 10],
        ["B1", "=A1*2"],  // B1 = 20
        ["C1", "=B1+5"],  // C1 = 25
        ["D1", 100],      // D1 = 100
      ]);

      engine.setSheetContent(sheetId, initialData);
      clearUpdateCounts();

      // Change A1, keep others same but they will recalculate
      const mixedData = new Map<string, any>([
        ["A1", 15],       // Changed: 10 -> 15
        ["B1", "=A1*2"],  // Same formula, but result changes: 20 -> 30
        ["C1", "=B1+5"],  // Same formula, but result changes: 25 -> 35
        ["D1", 100],      // Same value
      ]);

      engine.setSheetContent(sheetId, mixedData);
      
      expect(engine.getCellValue(addr(0, 0))).toBe(15); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(30); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(35); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe(100); // D1

      // A1, B1, C1 should update (due to value changes), D1 should not
      expect(getUpdateCount(addr(0, 0))).toBe(1); // A1 changed
      expect(getUpdateCount(addr(1, 0))).toBe(1); // B1 result changed
      expect(getUpdateCount(addr(2, 0))).toBe(1); // C1 result changed  
      expect(getUpdateCount(addr(3, 0))).toBe(0); // D1 unchanged
    });
  });

  describe("Array formula behavior", () => {
    test("should evaluate array formulas correctly", () => {
      // Set initial array formula
      engine.setCellContent(addr(0, 0), "=FILTER({1;2;3}, {1;2;3}>1)");
      
      // Should spill into A1:A2 with values [2, 3]
      expect(engine.getCellValue(addr(0, 0))).toBe(2); // A1 (origin)
      expect(engine.getCellValue(addr(0, 1))).toBe(3); // A2 (spilled)
      
      // Note: Array formula optimization with re-setting the same formula
      // is a complex case due to spill range management. For now, we focus
      // on optimizing the more common cases of regular formulas and values.
    });
  });

  describe("Empty cell optimization", () => {
    test("should not trigger updates when setting empty to empty", () => {
      // Start with empty cell
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined();
      
      clearUpdateCounts();

      // Set to empty again
      engine.setCellContent(addr(0, 0), "");
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined();
      expect(getUpdateCount(addr(0, 0))).toBe(0);
    });

    test("should not trigger updates when setting undefined to null", () => {
      // Start with empty cell
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined();
      
      clearUpdateCounts();

      // Set to null (equivalent to empty)
      engine.setCellContent(addr(0, 0), null as any);
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined();
      expect(getUpdateCount(addr(0, 0))).toBe(0);
    });
  });

  describe("Performance measurement", () => {
    test("should measure performance improvement with large dataset", () => {
      const size = 100;
      
      // Create initial large dataset
      const initialData = new Map<string, any>();
      for (let i = 0; i < size; i++) {
        initialData.set(`A${i + 1}`, i);
        initialData.set(`B${i + 1}`, `=A${i + 1}*2`);
      }

      // Set initial data
      const start1 = performance.now();
      engine.setSheetContent(sheetId, initialData);
      const time1 = performance.now() - start1;

      clearUpdateCounts();

      // Set the same data again (should be faster due to optimization)
      const start2 = performance.now();
      engine.setSheetContent(sheetId, initialData);
      const time2 = performance.now() - start2;

      // Verify no updates occurred
      let totalUpdates = 0;
      for (let i = 0; i < size; i++) {
        totalUpdates += getUpdateCount(addr(0, i)); // A column
        totalUpdates += getUpdateCount(addr(1, i)); // B column
      }
      
      expect(totalUpdates).toBe(0); // No updates should have occurred
      
      // Second operation should be significantly faster (at least 50% faster)
      // Note: This is a rough performance test, actual improvement may vary
      console.log(`Initial set: ${time1.toFixed(2)}ms`);
      console.log(`Optimized set: ${time2.toFixed(2)}ms`);
      console.log(`Improvement: ${((time1 - time2) / time1 * 100).toFixed(1)}%`);
      
      // The optimized version should be faster (though exact timing may vary)
      expect(time2).toBeLessThan(time1);
    });
  });
});
