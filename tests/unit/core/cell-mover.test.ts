import { describe, test, expect } from "bun:test";
import {
  updateReferencesForMovedCells,
  formulaReferencesCell,
  formulaReferencesRange,
  type MovedCellsInfo,
} from "../../../src/core/cell-mover";

describe("cell-mover utilities", () => {
  const workbookName = "TestWorkbook";
  const sheetName = "Sheet1";

  describe("updateReferencesForMovedCells - simple cell references", () => {
    test("should update simple cell reference when cell is moved", () => {
      // Moving A1 to D5 (offset: col+3, row+4)
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // A1
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("A1+B1", movedCells);
      expect(result).toBe("D5+B1"); // Only A1 updated to D5
    });

    test("should not update cell reference if cell not moved", () => {
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // A1
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("B1+C1", movedCells);
      expect(result).toBe("B1+C1"); // Neither moved
    });

    test("should update absolute cell references", () => {
      // Moving A1 to D5
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // A1
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("$A$1+B1", movedCells);
      expect(result).toBe("$D$5+B1"); // Absolute reference updated
    });

    test("should update mixed relative/absolute references", () => {
      // Moving A1 to D5
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // A1
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("$A1+A$1", movedCells);
      expect(result).toBe("$D5+D$5"); // Both updated
    });

    test("should update multiple references to same moved cell", () => {
      // Moving A1 to D5
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // A1
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("A1+A1*A1", movedCells);
      expect(result).toBe("D5+D5*D5"); // All instances updated
    });

    test("should update cell references in complex formula", () => {
      // Moving B2 to E10
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:1:1`]), // B2
        workbookName,
        sheetName,
        rowOffset: 8,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("IF(B2>10,B2*2,B2/2)", movedCells);
      expect(result).toBe("IF(E10>10,E10*2,E10/2)");
    });
  });

  describe("updateReferencesForMovedCells - range references", () => {
    test("should update range when entire range is moved", () => {
      // Moving A1:D5 to F10 (offset: col+5, row+9)
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([
          `${workbookName}:${sheetName}:0:0`, // A1
          `${workbookName}:${sheetName}:1:0`, // B1
          `${workbookName}:${sheetName}:2:0`, // C1
          `${workbookName}:${sheetName}:3:0`, // D1
          `${workbookName}:${sheetName}:0:1`, // A2
          `${workbookName}:${sheetName}:1:1`, // B2
          `${workbookName}:${sheetName}:2:1`, // C2
          `${workbookName}:${sheetName}:3:1`, // D2
          `${workbookName}:${sheetName}:0:2`, // A3
          `${workbookName}:${sheetName}:1:2`, // B3
          `${workbookName}:${sheetName}:2:2`, // C3
          `${workbookName}:${sheetName}:3:2`, // D3
          `${workbookName}:${sheetName}:0:3`, // A4
          `${workbookName}:${sheetName}:1:3`, // B4
          `${workbookName}:${sheetName}:2:3`, // C4
          `${workbookName}:${sheetName}:3:3`, // D4
          `${workbookName}:${sheetName}:0:4`, // A5
          `${workbookName}:${sheetName}:1:4`, // B5
          `${workbookName}:${sheetName}:2:4`, // C5
          `${workbookName}:${sheetName}:3:4`, // D5
        ]),
        workbookName,
        sheetName,
        rowOffset: 9,
        colOffset: 5,
      };

      const result = updateReferencesForMovedCells("SUM(A1:D5)", movedCells);
      expect(result).toBe("SUM(F10:I14)"); // Range updated
    });

    test("should NOT update range when only part of it is moved", () => {
      // Moving only A1:B5 (left half of A1:D5)
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([
          `${workbookName}:${sheetName}:0:0`, // A1
          `${workbookName}:${sheetName}:1:0`, // B1
          `${workbookName}:${sheetName}:0:1`, // A2
          `${workbookName}:${sheetName}:1:1`, // B2
          `${workbookName}:${sheetName}:0:2`, // A3
          `${workbookName}:${sheetName}:1:2`, // B3
          `${workbookName}:${sheetName}:0:3`, // A4
          `${workbookName}:${sheetName}:1:3`, // B4
          `${workbookName}:${sheetName}:0:4`, // A5
          `${workbookName}:${sheetName}:1:4`, // B5
        ]),
        workbookName,
        sheetName,
        rowOffset: 9,
        colOffset: 5,
      };

      const result = updateReferencesForMovedCells("SUM(A1:D5)", movedCells);
      expect(result).toBe("SUM(A1:D5)"); // Range NOT updated (partial move)
    });

    test("should update range with absolute references", () => {
      // Moving A1:B2 to D5
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([
          `${workbookName}:${sheetName}:0:0`, // A1
          `${workbookName}:${sheetName}:1:0`, // B1
          `${workbookName}:${sheetName}:0:1`, // A2
          `${workbookName}:${sheetName}:1:1`, // B2
        ]),
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("SUM($A$1:$B$2)", movedCells);
      expect(result).toBe("SUM($D$5:$E$6)");
    });

    test("should handle mixed absolute/relative range references", () => {
      // Moving A1:B2 to D5
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([
          `${workbookName}:${sheetName}:0:0`,
          `${workbookName}:${sheetName}:1:0`,
          `${workbookName}:${sheetName}:0:1`,
          `${workbookName}:${sheetName}:1:1`,
        ]),
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("SUM($A1:B$2)", movedCells);
      expect(result).toBe("SUM($D5:E$6)");
    });
  });

  describe("updateReferencesForMovedCells - cross-sheet references", () => {
    test("should not update references to different sheet", () => {
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // Sheet1!A1
        workbookName,
        sheetName: "Sheet1",
        rowOffset: 4,
        colOffset: 3,
      };

      // Reference to Sheet2!A1 should not be updated
      const result = updateReferencesForMovedCells("Sheet2!A1", movedCells);
      expect(result).toBe("Sheet2!A1");
    });

    test("should update cross-sheet references when they match", () => {
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:Sheet1:0:0`]), // Sheet1!A1
        workbookName,
        sheetName: "Sheet1",
        rowOffset: 4,
        colOffset: 3,
      };

      // Reference to Sheet1!A1 from another sheet should be updated
      const result = updateReferencesForMovedCells("Sheet1!A1", movedCells);
      expect(result).toBe("Sheet1!D5");
    });
  });

  describe("updateReferencesForMovedCells - edge cases", () => {
    test("should handle formula with no references", () => {
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]),
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("5+10", movedCells);
      expect(result).toBe("5+10"); // No change
    });

    test("should handle invalid formula gracefully", () => {
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]),
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("INVALID((((", movedCells);
      expect(result).toBe("INVALID(((("); // Return original on parse error
    });

    test("should handle single cell range (A1:A1)", () => {
      const movedCells: MovedCellsInfo = {
        cellsSet: new Set([`${workbookName}:${sheetName}:0:0`]), // A1
        workbookName,
        sheetName,
        rowOffset: 4,
        colOffset: 3,
      };

      const result = updateReferencesForMovedCells("SUM(A1:A1)", movedCells);
      expect(result).toBe("SUM(D5:D5)"); // Updated
    });
  });

  describe("formulaReferencesCell", () => {
    test("should detect simple cell reference", () => {
      const result = formulaReferencesCell("A1+B1", workbookName, sheetName, 0, 0);
      expect(result).toBe(true); // References A1
    });

    test("should return false when cell not referenced", () => {
      const result = formulaReferencesCell("B1+C1", workbookName, sheetName, 0, 0);
      expect(result).toBe(false); // Does not reference A1
    });

    test("should detect cell in range", () => {
      // Note: This checks if cell is explicitly referenced, not if it's in a range
      const result = formulaReferencesCell("SUM(A1:D5)", workbookName, sheetName, 0, 0);
      expect(result).toBe(false); // A1 is in the range, not referenced individually
    });

    test("should detect cross-sheet reference", () => {
      const result = formulaReferencesCell("Sheet1!A1", workbookName, "Sheet1", 0, 0);
      expect(result).toBe(true);
    });
  });

  describe("formulaReferencesRange", () => {
    test("should detect exact range reference", () => {
      const result = formulaReferencesRange(
        "SUM(A1:D5)",
        workbookName,
        sheetName,
        0, 0, // A1
        3, 4  // D5
      );
      expect(result).toBe(true);
    });

    test("should return false for different range", () => {
      const result = formulaReferencesRange(
        "SUM(A1:B2)",
        workbookName,
        sheetName,
        0, 0, // A1
        3, 4  // D5 (looking for A1:D5)
      );
      expect(result).toBe(false);
    });

    test("should detect cross-sheet range reference", () => {
      const result = formulaReferencesRange(
        "Sheet1!A1:D5",
        workbookName,
        "Sheet1",
        0, 0,
        3, 4
      );
      expect(result).toBe(true);
    });
  });
});

