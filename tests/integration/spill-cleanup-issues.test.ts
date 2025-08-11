import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../src/core/engine";
import type { SimpleCellAddress } from "../../src/core/types";

describe("Array Formula Spill Cleanup Issues", () => {
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

  describe("Issue 1: Clearing array formula should remove spilled cells", () => {
    test("should remove all spilled cells when array formula is cleared via setSheetContent", () => {
      // Step 1: Set array formula that creates spilled cells
      const withArrayFormula = new Map<string, any>([
        ["A1", "=SEQUENCE(5)"], // Should spill into A1:A5 with values [1,2,3,4,5]
      ]);

      engine.setSheetContent(sheetId, withArrayFormula);

      // Verify initial spill
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1 (origin)
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2 (spilled)
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3 (spilled)
      expect(engine.getCellValue(addr(0, 3))).toBe(4); // A4 (spilled)
      expect(engine.getCellValue(addr(0, 4))).toBe(5); // A5 (spilled)

      console.log("Initial spill values:");
      for (let i = 0; i < 5; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Step 2: Clear the array formula by setting empty content
      const clearedContent = new Map<string, any>([
        ["A1", ""], // Clear the array formula
      ]);

      engine.setSheetContent(sheetId, clearedContent);

      console.log("After clearing:");
      for (let i = 0; i < 5; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Step 3: Verify all spilled cells are cleared
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined(); // A1 should be empty
      expect(engine.getCellValue(addr(0, 1))).toBeUndefined(); // A2 should be empty
      expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3 should be empty
      expect(engine.getCellValue(addr(0, 3))).toBeUndefined(); // A4 should be empty
      expect(engine.getCellValue(addr(0, 4))).toBeUndefined(); // A5 should be empty
    });

    test("should remove all spilled cells when array formula is cleared by omitting from setSheetContent", () => {
      // Step 1: Set array formula
      const withArrayFormula = new Map<string, any>([
        ["A1", "=SEQUENCE(3)"], // Should spill into A1:A3
        ["B1", "Some other cell"], // Another cell to keep
      ]);

      engine.setSheetContent(sheetId, withArrayFormula);

      // Verify initial spill
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3
      expect(engine.getCellValue(addr(1, 0))).toBe("Some other cell"); // B1

      // Step 2: Call setSheetContent without the array formula (implicit clear)
      const withoutArrayFormula = new Map<string, any>([
        ["B1", "Some other cell"], // Keep B1, but omit A1
      ]);

      engine.setSheetContent(sheetId, withoutArrayFormula);

      // Step 3: Verify all spilled cells are cleared
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined(); // A1 should be cleared
      expect(engine.getCellValue(addr(0, 1))).toBeUndefined(); // A2 should be cleared
      expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3 should be cleared
      expect(engine.getCellValue(addr(1, 0))).toBe("Some other cell"); // B1 should remain
    });
  });

  describe("Issue 2: Modifying array formula should handle spill range changes", () => {
    test("should handle expanding array formula (smaller to larger spill)", () => {
      // Step 1: Set small array formula
      const smallArray = new Map<string, any>([
        ["A1", "=SEQUENCE(3)"], // Spills into A1:A3
      ]);

      engine.setSheetContent(sheetId, smallArray);

      // Verify initial small spill
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3
      expect(engine.getCellValue(addr(0, 3))).toBeUndefined(); // A4 should be empty

      console.log("Small array values:");
      for (let i = 0; i < 5; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Step 2: Expand to larger array
      const largerArray = new Map<string, any>([
        ["A1", "=SEQUENCE(5)"], // Now spills into A1:A5
      ]);

      engine.setSheetContent(sheetId, largerArray);

      console.log("Large array values:");
      for (let i = 0; i < 5; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Step 3: Verify expanded spill
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3
      expect(engine.getCellValue(addr(0, 3))).toBe(4); // A4 (new)
      expect(engine.getCellValue(addr(0, 4))).toBe(5); // A5 (new)
    });

    test("should handle shrinking array formula (larger to smaller spill)", () => {
      // Step 1: Set large array formula
      const largeArray = new Map<string, any>([
        ["A1", "=SEQUENCE(5)"], // Spills into A1:A5
      ]);

      engine.setSheetContent(sheetId, largeArray);

      // Verify initial large spill
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3
      expect(engine.getCellValue(addr(0, 3))).toBe(4); // A4
      expect(engine.getCellValue(addr(0, 4))).toBe(5); // A5

      console.log("Large array values:");
      for (let i = 0; i < 6; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Step 2: Shrink to smaller array
      const smallerArray = new Map<string, any>([
        ["A1", "=SEQUENCE(2)"], // Now only spills into A1:A2
      ]);

      engine.setSheetContent(sheetId, smallerArray);

      console.log("Small array values:");
      for (let i = 0; i < 6; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Step 3: Verify shrunk spill and cleanup
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3 should be cleared
      expect(engine.getCellValue(addr(0, 3))).toBeUndefined(); // A4 should be cleared
      expect(engine.getCellValue(addr(0, 4))).toBeUndefined(); // A5 should be cleared
    });

    test("should handle spill blocking and cleanup properly", () => {
      // Step 1: Create blocking scenario
      const initialContent = new Map<string, any>([
        ["A1", "=SEQUENCE(5)"], // Should spill into A1:A5
        ["A3", "Blocking value"], // This will block the spill
      ]);

      engine.setSheetContent(sheetId, initialContent);

      console.log("With blocking value:");
      for (let i = 0; i < 6; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Should get #SPILL! error due to blocking
      expect(engine.getCellValue(addr(0, 0))).toBe("#SPILL!"); // A1 should show error
      expect(engine.getCellValue(addr(0, 2))).toBe("Blocking value"); // A3 has blocking value

      // Step 2: Remove the blocking value
      const unblocked = new Map<string, any>([
        ["A1", "=SEQUENCE(5)"], // Same formula
        // A3 is omitted, so it should be cleared
      ]);

      engine.setSheetContent(sheetId, unblocked);

      console.log("After removing blocking value:");
      for (let i = 0; i < 6; i++) {
        console.log(`A${i+1}:`, engine.getCellValue(addr(0, i)));
      }

      // Now should spill properly
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3 (was blocking, now spilled)
      expect(engine.getCellValue(addr(0, 3))).toBe(4); // A4
      expect(engine.getCellValue(addr(0, 4))).toBe(5); // A5
    });
  });

  describe("Dependency tracking with spilled cells", () => {
    test("should update dependents when spilled cells are cleared", () => {
      // Step 1: Set array formula that creates spilled cells
      const withArrayFormula = new Map<string, any>([
        ["A1", "=SEQUENCE(3)"], // Should spill A1:A3 with values [1,2,3]
        ["B1", "=A3*3"],        // Depends on spilled cell A3, should be 3*3=9
      ]);

      engine.setSheetContent(sheetId, withArrayFormula);



      // Verify initial state
    //   expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
    //   expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2 (spilled)
    //   expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3 (spilled)
    //   expect(engine.getCellValue(addr(1, 0))).toBe(9); // B1 = A3*3 = 3*3 = 9

      // Step 2: Clear the array formula, which should clear spilled cells A2, A3
      const clearedArray = new Map<string, any>([
        // A1 is omitted, so it should be cleared along with its spills
        ["B1", "=A3*3"], // Keep B1 formula, but A3 should now be undefined
      ]);



      engine.setSheetContent(sheetId, clearedArray);




      // Step 3: Verify spilled cells are cleared
    //   expect(engine.getCellValue(addr(0, 0))).toBeUndefined(); // A1
    //   expect(engine.getCellValue(addr(0, 1))).toBeUndefined(); // A2
    //   expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3

      // Step 4: Verify dependent is updated
      // B1 should be re-evaluated because A3 (its dependency) was cleared
      // B1 should evaluate to 0 when referencing a cleared cell
      const b1Value = engine.getCellValue(addr(1, 0));
      expect(b1Value).toBe(0);
    });

    test("should update multiple dependents when spilled cells are cleared", () => {
      // More complex scenario with multiple dependents
      const complexScenario = new Map<string, any>([
        ["A1", "=SEQUENCE(4)"],    // A1:A4 = [1,2,3,4]
        ["B1", "=A2+A4"],         // B1 = 2+4 = 6
        ["C1", "=A3*2"],          // C1 = 3*2 = 6
        ["D1", "=SUM(A1:A4)"],    // D1 = 1+2+3+4 = 10
      ]);

      engine.setSheetContent(sheetId, complexScenario);

      // Verify initial state
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3  
      expect(engine.getCellValue(addr(0, 3))).toBe(4); // A4
      expect(engine.getCellValue(addr(1, 0))).toBe(6); // B1 = A2+A4 = 2+4
      expect(engine.getCellValue(addr(2, 0))).toBe(6); // C1 = A3*2 = 3*2
      expect(engine.getCellValue(addr(3, 0))).toBe(10); // D1 = SUM(A1:A4)

      // Clear the array formula
      const cleared = new Map<string, any>([
        ["B1", "=A2+A4"],         // Keep formulas but A2,A4 will be undefined
        ["C1", "=A3*2"],          // A3 will be undefined
        ["D1", "=SUM(A1:A4)"],    // Range A1:A4 will be empty
      ]);

      engine.setSheetContent(sheetId, cleared);

      // Verify spilled cells are cleared
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined(); // A1
      expect(engine.getCellValue(addr(0, 1))).toBeUndefined(); // A2
      expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3
      expect(engine.getCellValue(addr(0, 3))).toBeUndefined(); // A4

      // Verify all dependents are updated
      // All references to cleared cells should evaluate to 0
      const b1Value = engine.getCellValue(addr(1, 0));
      const c1Value = engine.getCellValue(addr(2, 0));
      const d1Value = engine.getCellValue(addr(3, 0));
      
      // B1 and C1 reference individual cleared cells, should evaluate to 0
      expect(b1Value).toBe(0);
      expect(c1Value).toBe(0);
      // D1 uses SUM which should handle empty ranges gracefully
      expect(d1Value).toBe(0); // SUM of empty range is always 0
    });
  });

  describe("Edge cases and complex scenarios", () => {
    test("should handle multiple array formulas in the same setSheetContent", () => {
      const multipleArrays = new Map<string, any>([
        ["A1", "=SEQUENCE(3)"],     // A1:A3
        ["C1", "=SEQUENCE(2,2)"],   // C1:D2 (2x2 grid)
      ]);

      engine.setSheetContent(sheetId, multipleArrays);

      // Verify both arrays spill correctly
      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3

      expect(engine.getCellValue(addr(2, 0))).toBe(1); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe(2); // D1
      expect(engine.getCellValue(addr(2, 1))).toBe(3); // C2
      expect(engine.getCellValue(addr(3, 1))).toBe(4); // D2

      // Clear one array
      const oneArrayCleared = new Map<string, any>([
        ["C1", "=SEQUENCE(2,2)"],   // Keep C1:D2
        // A1 is omitted
      ]);

      engine.setSheetContent(sheetId, oneArrayCleared);

      // A array should be cleared, C array should remain
      expect(engine.getCellValue(addr(0, 0))).toBeUndefined(); // A1
      expect(engine.getCellValue(addr(0, 1))).toBeUndefined(); // A2
      expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3

      expect(engine.getCellValue(addr(2, 0))).toBe(1); // C1
      expect(engine.getCellValue(addr(3, 0))).toBe(2); // D1
      expect(engine.getCellValue(addr(2, 1))).toBe(3); // C2
      expect(engine.getCellValue(addr(3, 1))).toBe(4); // D2
    });

    test("should handle replacing array formula with regular value", () => {
      // Step 1: Set array formula
      const arrayFormula = new Map<string, any>([
        ["A1", "=SEQUENCE(4)"], // A1:A4
      ]);

      engine.setSheetContent(sheetId, arrayFormula);

      expect(engine.getCellValue(addr(0, 0))).toBe(1); // A1
      expect(engine.getCellValue(addr(0, 1))).toBe(2); // A2
      expect(engine.getCellValue(addr(0, 2))).toBe(3); // A3
      expect(engine.getCellValue(addr(0, 3))).toBe(4); // A4

      // Step 2: Replace with regular value
      const regularValue = new Map<string, any>([
        ["A1", "Just a string"], // Replace array formula with string
      ]);

      engine.setSheetContent(sheetId, regularValue);

      // Only A1 should have the string, others should be cleared
      expect(engine.getCellValue(addr(0, 0))).toBe("Just a string"); // A1
      expect(engine.getCellValue(addr(0, 1))).toBeUndefined(); // A2
      expect(engine.getCellValue(addr(0, 2))).toBeUndefined(); // A3
      expect(engine.getCellValue(addr(0, 3))).toBeUndefined(); // A4
    });
  });
});
