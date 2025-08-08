import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../src/core/engine";
import type { SimpleCellAddress } from "../../src/core/types";

describe("Topological Ordering in Deferred Evaluation", () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet("TestSheet");
    sheetId = engine.getSheetId(sheetName);
  });

  const addr = (col: number, row: number): SimpleCellAddress => ({
    sheet: sheetId,
    col,
    row,
  });

  describe("Dependency Chain Evaluation", () => {
    test("should evaluate simple dependency chain in correct order", () => {
      // A1 = 10, B1 = A1 * 2, C1 = B1 + 5
      // Expected: A1=10, B1=20, C1=25
      const formulaMap = new Map([
        ["C1", "=B1+5"],  // Insert in reverse order to test topological sort
        ["B1", "=A1*2"],
        ["A1", "10"],
      ]);

      engine.setSheetContent(sheetId, formulaMap);

      expect(engine.getCellValue(addr(0, 0))).toBe(10); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(20); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(25); // C1
    });

    test("should handle complex dependency chain with multiple dependencies", () => {
      // Z1 = 10 (base value)
      // A1 = Z1 * 2 (depends on Z1) 
      // B1 = A1 + 5 (depends on A1)
      // C1 = B1 * 3 (depends on B1)
      // D1 = C1 + A1 (depends on both C1 and A1)
      //
      // Correct topological order: Z1 -> A1 -> B1 -> C1 -> D1
      
      const formulaMap = new Map([
        // Deliberately insert in "wrong" order to test topological sorting
        ["D1", "=C1+A1"],  // Depends on C1 and A1
        ["A1", "=Z1*2"],   // Depends on Z1
        ["C1", "=B1*3"],   // Depends on B1
        ["B1", "=A1+5"],   // Depends on A1
        ["Z1", "10"],      // Base value
      ]);

      engine.setSheetContent(sheetId, formulaMap);

      // Expected values:
      // Z1 = 10
      // A1 = Z1 * 2 = 10 * 2 = 20
      // B1 = A1 + 5 = 20 + 5 = 25
      // C1 = B1 * 3 = 25 * 3 = 75
      // D1 = C1 + A1 = 75 + 20 = 95

      expect(engine.getCellValue(addr(25, 0))).toBe(10); // Z1
      expect(engine.getCellValue(addr(0, 0))).toBe(20);  // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(25);  // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(75);  // C1
      expect(engine.getCellValue(addr(3, 0))).toBe(95);  // D1
    });

    test("should handle diamond dependency pattern", () => {
      // A1 = 10
      // B1 = A1 * 2, C1 = A1 + 5 (both depend on A1)
      // D1 = B1 + C1 (depends on both B1 and C1)
      
      const formulaMap = new Map([
        ["D1", "=B1+C1"],  // Depends on B1 and C1
        ["C1", "=A1+5"],   // Depends on A1
        ["B1", "=A1*2"],   // Depends on A1
        ["A1", "10"],      // Base value
      ]);

      engine.setSheetContent(sheetId, formulaMap);

      // Expected values:
      // A1 = 10
      // B1 = A1 * 2 = 10 * 2 = 20
      // C1 = A1 + 5 = 10 + 5 = 15
      // D1 = B1 + C1 = 20 + 15 = 35

      expect(engine.getCellValue(addr(0, 0))).toBe(10); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(20); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(15); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe(35); // D1
    });

    test("should handle dependencies with OFFSET function", () => {
      // Similar to the original OFFSET issue but with more complex dependencies
      // A1=1, A2=2, A3=3
      // B1=A1*10, B2=A2*10, B3=A3*10
      // C1=SUM(OFFSET(B1,0,0,3,1)) (depends on B1:B3 range)
      
      const formulaMap = new Map([
        ["C1", "=SUM(OFFSET(B1,0,0,3,1))"], // Insert first to test ordering
        ["B3", "=A3*10"],
        ["B1", "=A1*10"],
        ["A2", "2"],
        ["B2", "=A2*10"],
        ["A1", "1"],
        ["A3", "3"],
      ]);

      engine.setSheetContent(sheetId, formulaMap);

      // Expected values:
      // A1=1, A2=2, A3=3
      // B1=10, B2=20, B3=30
      // C1=SUM(B1:B3)=10+20+30=60

      expect(engine.getCellValue(addr(0, 0))).toBe(1);  // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2);  // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3);  // A3
      expect(engine.getCellValue(addr(1, 0))).toBe(10); // B1
      expect(engine.getCellValue(addr(1, 1))).toBe(20); // B2
      expect(engine.getCellValue(addr(1, 2))).toBe(30); // B3
      expect(engine.getCellValue(addr(2, 0))).toBe(60); // C1
    });
  });

  describe("Circular Dependency Detection", () => {
    test("should detect simple circular dependency", () => {
      // A1 = B1, B1 = A1 (simple cycle)
      const circularMap = new Map([
        ["A1", "=B1"],
        ["B1", "=A1"],
      ]);

      engine.setSheetContent(sheetId, circularMap);

      // Should return #CYCLE! for both cells
      expect(engine.getCellValue(addr(0, 0))).toBe("#CYCLE!"); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe("#CYCLE!"); // B1
    });

    test("should detect three-node circular dependency", () => {
      // A1 -> B1 -> C1 -> A1 (three-node cycle)
      const circularMap = new Map([
        ["A1", "=B1+1"],
        ["B1", "=C1+1"], 
        ["C1", "=A1+1"],
      ]);

      engine.setSheetContent(sheetId, circularMap);

      // All should return #CYCLE!
      expect(engine.getCellValue(addr(0, 0))).toBe("#CYCLE!"); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe("#CYCLE!"); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe("#CYCLE!"); // C1
    });

    test("should detect complex circular dependency with side branches", () => {
      // A1 = 10 (independent)
      // B1 = A1 * 2 (depends on A1, not part of cycle)
      // C1 = D1 + 1, D1 = E1 + 1, E1 = C1 + 1 (cycle: C1 -> D1 -> E1 -> C1)
      // F1 = B1 + 5 (depends on B1, not part of cycle)
      
      const mixedMap = new Map([
        ["F1", "=B1+5"],   // Independent of cycle
        ["C1", "=D1+1"],   // Part of cycle
        ["B1", "=A1*2"],   // Independent of cycle
        ["E1", "=C1+1"],   // Part of cycle
        ["D1", "=E1+1"],   // Part of cycle
        ["A1", "10"],      // Independent
      ]);

      engine.setSheetContent(sheetId, mixedMap);

      // Non-cycle cells should evaluate correctly
      expect(engine.getCellValue(addr(0, 0))).toBe(10); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(20); // B1
      expect(engine.getCellValue(addr(5, 0))).toBe(25); // F1

      // Cycle cells should return #CYCLE!
      expect(engine.getCellValue(addr(2, 0))).toBe("#CYCLE!"); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe("#CYCLE!"); // D1
      expect(engine.getCellValue(addr(4, 0))).toBe("#CYCLE!"); // E1
    });

    test("should handle cycle with OFFSET function", () => {
      // A1 = SUM(OFFSET(B1,0,0,1,1)), B1 = A1 + 1 (cycle through OFFSET)
      const offsetCycleMap = new Map([
        ["A1", "=SUM(OFFSET(B1,0,0,1,1))"],
        ["B1", "=A1+1"],
      ]);

      engine.setSheetContent(sheetId, offsetCycleMap);

      // Both should detect the cycle
      expect(engine.getCellValue(addr(0, 0))).toBe("#CYCLE!"); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe("#CYCLE!"); // B1
    });
  });

  describe("Mixed Scenarios", () => {
    test("should handle independent formulas correctly", () => {
      // Mix of independent values and dependent formulas
      const mixedMap = new Map([
        ["F1", "=D1+E1"],  // Depends on D1 and E1
        ["D1", "=A1*2"],   // Depends on A1
        ["A1", "100"],     // Independent 
        ["E1", "=B1+C1"],  // Depends on B1 and C1
        ["B1", "200"],     // Independent
        ["C1", "300"],     // Independent
      ]);

      engine.setSheetContent(sheetId, mixedMap);

      // Expected values:
      // A1 = 100, B1 = 200, C1 = 300 (independent)
      // D1 = A1 * 2 = 100 * 2 = 200
      // E1 = B1 + C1 = 200 + 300 = 500
      // F1 = D1 + E1 = 200 + 500 = 700

      expect(engine.getCellValue(addr(0, 0))).toBe(100); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(200); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(300); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe(200); // D1
      expect(engine.getCellValue(addr(4, 0))).toBe(500); // E1
      expect(engine.getCellValue(addr(5, 0))).toBe(700); // F1
    });

    test("should handle partial cycles with independent branches", () => {
      // Some cells form cycles, others are independent
      const partialCycleMap = new Map([
        ["A1", "10"],        // Independent
        ["B1", "=A1*2"],     // Depends on A1 (independent branch)
        ["C1", "=D1+1"],     // Part of cycle
        ["D1", "=C1+1"],     // Part of cycle  
        ["E1", "=B1+5"],     // Depends on B1 (independent branch)
      ]);

      engine.setSheetContent(sheetId, partialCycleMap);

      // Independent branch should work
      expect(engine.getCellValue(addr(0, 0))).toBe(10); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(20); // B1
      expect(engine.getCellValue(addr(4, 0))).toBe(25); // E1

      // Cycle should be detected
      expect(engine.getCellValue(addr(2, 0))).toBe("#CYCLE!"); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe("#CYCLE!"); // D1
    });

    test("should maintain evaluation order consistency across multiple calls", () => {
      // Test that the same formula set produces consistent results
      const formulaMap = new Map([
        ["D1", "=A1+B1+C1"],
        ["C1", "=A1*3"],
        ["B1", "=A1*2"],
        ["A1", "5"],
      ]);

      // Set content multiple times and verify consistency
      for (let i = 0; i < 3; i++) {
        engine.setSheetContent(sheetId, formulaMap);
        
        expect(engine.getCellValue(addr(0, 0))).toBe(5);  // A1
        expect(engine.getCellValue(addr(1, 0))).toBe(10); // B1
        expect(engine.getCellValue(addr(2, 0))).toBe(15); // C1
        expect(engine.getCellValue(addr(3, 0))).toBe(30); // D1
      }
    });
  });

  describe("Edge Cases", () => {
    test("should handle empty dependency chain", () => {
      // Just independent values, no formulas
      const independentMap = new Map([
        ["A1", "10"],
        ["B1", "20"],
        ["C1", "30"],
      ]);

      engine.setSheetContent(sheetId, independentMap);

      expect(engine.getCellValue(addr(0, 0))).toBe(10); // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(20); // B1
      expect(engine.getCellValue(addr(2, 0))).toBe(30); // C1
    });

    test("should handle single formula dependency", () => {
      // Only one formula depending on one value
      const singleMap = new Map([
        ["B1", "=A1*2"],
        ["A1", "7"],
      ]);

      engine.setSheetContent(sheetId, singleMap);

      expect(engine.getCellValue(addr(0, 0))).toBe(7);  // A1
      expect(engine.getCellValue(addr(1, 0))).toBe(14); // B1
    });

    test("should handle self-referencing formula", () => {
      // A1 = A1 + 1 (self-cycle)
      const selfRefMap = new Map([
        ["A1", "=A1+1"],
      ]);

      engine.setSheetContent(sheetId, selfRefMap);

      expect(engine.getCellValue(addr(0, 0))).toBe("#CYCLE!"); // A1
    });
  });
});
