import { describe, test as it, expect, beforeEach } from "bun:test";
import { OpenRangeEvaluator } from "./open-range-evaluator";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue, type SpreadsheetRange } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("OpenRangeEvaluator", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  /**
   * Get the cell value from the engine
   * @param ref - The cell reference, e.g. "A1"
   * @param debug - Whether to include debug information in the result
   * @returns The cell value
   */
  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  describe("Basic open range evaluation", () => {
    it("should evaluate all values in an open row range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:B)"],
          ["B10", 10],
          ["B11", 20],
          ["B12", 30],
          ["B15", 40],
        ])
      );

      // Test SUM(B10:B) - open ended row (infinite rows)
      const result = cell("A1");
      expect(result).toBe(100);
    });

    it("should evaluate all values in an open column range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:10)"],
          ["B10", 10],
          ["C10", 20],
          ["D10", 30],
          ["F10", 40],
        ])
      );

      // Test SUM(B10:10) - open ended column (infinite columns)
      const result = cell("A1");
      expect(result).toBe(100);
    });

    it("should handle empty open ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([["A1", "=SUM(B10:INFINITY)"]])
      );

      // Test SUM(B10:INFINITY) with no values
      const result = cell("A1");
      expect(result).toBe(0);
    });

    it("should handle mixed content in open ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:B)"],
          ["B10", 10],
          ["B11", "text"], // Should cause #VALUE! error
          ["B12", 20],
          ["B13", true], // Should cause #VALUE! error
        ])
      );

      const result = cell("A1");
      expect(result).toBe("#VALUE!");
    });
  });

  describe("Spilled values in open ranges", () => {
    it("should include finite spilled values in range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:INFINITY)"],
          ["A10", "=SEQUENCE(3,3)"],
        ])
      );

      // SUM(B10:INFINITY) should include the spilled values B10:C12
      const result = cell("A1", true);
      // SEQUENCE(3,3) produces 1-9, B10:C12 contains: 2,3,5,6,8,9 = 33
      expect(result).toBe(33);
    });

    it("should return error for infinite spills in range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:D)"],
          ["B100", "=SEQUENCE(INFINITY)"],
        ])
      );

      // SUM(B10:D) should detect the infinite spill and return INFINITY
      const result = cell("A1", true);
      expect(result).toBe("#REF!: Can not evaluate all cells over an infinite range");
    });

    it("should handle partial spill intersections", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:D)"],
          ["A100", "=SEQUENCE(1,INFINITY)"],
        ])
      );

      // SUM(B10:D) should only sum the intersection B100:D100
      const result = cell("A1", true);
      // Values would be 2, 3, 4 = 9
      expect(result).toBe(9);
    });

    it("should handle spills that don't intersect the range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SEQUENCE(3,2)"],
          ["A5", "=SUM(B10:INFINITY)"],
          ["B10", 100],
        ])
      );

      const result = cell("A5");
      expect(result).toBe(100);
    });
  });

  describe("Frontier candidate detection", () => {
    it("should detect and evaluate top frontier candidates", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:D)"],
          ["C8", "=SEQUENCE(5,1)"],
        ])
      );

      // SUM(B10:D) should detect C8 as a frontier candidate
      const result = cell("A1", true);
      // C8 spills to C8:C12, so C10:C12 (values 3,4,5) = 12
      expect(result).toBe(12);
    });

    it("should detect and evaluate left frontier candidates", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:INFINITY)"],
          ["A10", "=SEQUENCE(1,5)"],
        ])
      );

      // SUM(B10:INFINITY) should detect A10 as a frontier candidate
      const result = cell("A1", true);
      // A10 spills to A10:E10, so B10:E10 (values 2,3,4,5) = 14
      expect(result).toBe(14);
    });

    it("should ignore blocked frontier candidates", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(C10:INFINITY)"],
          ["A8", "=SEQUENCE(5,5)"],
          ["B9", 999],
        ])
      );

      // SUM(C10:INFINITY) should not evaluate A8 because it's blocked by B9
      const result = cell("A1");
      // Only the direct intersection matters, not the blocked spill
      expect(result).toBe(0);
    });

    it("should handle multiple frontier candidates", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:D)"],
          ["C8", "=SEQUENCE(3,1)"],
          ["A10", "=SEQUENCE(1,3)"],
        ])
      );

      // SUM(B10:D) should evaluate both candidates
      const result = cell("A1", true);
      // C10 from top spill (value 3) + B10:C10 from left spill (values 2,3)
      // Note: C10 appears in both, so we need to handle deduplication
      expect(result).toBe(8);
    });
  });

  describe("Error handling", () => {
    it("should propagate errors from cells in range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:B)"],
          ["B10", 10],
          ["B11", "=WEFWEF"], // NAME error
          ["B12", 20],
        ])
      );
      const result = cell("A1");
      expect(result).toBe("#NAME?");
    });

    it("should handle circular references in frontier candidates", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(A10:B)"],
          ["A10", "=B10"],
          ["B10", "=A10"],
        ])
      );

      const result = cell("A1");
      expect(result).toBe("#CYCLE!");
    });

    it("should handle missing sheet references", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(NonExistentSheet!B10:INFINITY)"],
        ])
      );

      const result = cell("A1");
      expect(result).toBe("#REF!");
    });
  });

  describe("Performance scenarios", () => {
    it("should efficiently handle large sheets with sparse data", () => {
      const content = new Map<string, SerializedCellValue>([
        ["A1", "=SUM(B10:B)"],
      ]);

      // Add values at various locations
      for (let i = 0; i < 100; i++) {
        content.set(`B${10 + i * 10}`, i);
      }

      const start = performance.now();

      engine.setSheetContent(sheetName, content);

      const duration = performance.now() - start;
      const result = cell("A1");

      expect(result).toBe(4950);
      expect(duration).toBeLessThan(100); // Should complete quickly
    });

    it("should handle multiple open ranges efficiently", () => {
      const content = new Map<string, SerializedCellValue>([
        ["A1", "=SUM(B10:B) + SUM(C10:C) + SUM(D10:D)"],
      ]);

      // Set up data
      for (let row = 10; row <= 20; row++) {
        for (let col = 1; col <= 5; col++) {
          const colLetter = String.fromCharCode(65 + col);
          content.set(`${colLetter}${row}`, row * col);
        }
      }

      engine.setSheetContent(sheetName, content);

      // Multiple SUMs with open ranges
      const result = cell("A1");

      expect(result).toBeDefined();
      expect(typeof result).toBe("number");
    });
  });

  describe("Integration with other functions", () => {
    it("should work with AVERAGE on open ranges", () => {
      const content = new Map<string, SerializedCellValue>([
        ["A1", "=AVERAGE(B10:B)"],
        ["B10", 10],
        ["B11", 20],
        ["B12", 30],
      ]);

      engine.setSheetContent(sheetName, content);

      const result = cell("A1", true);
      expect(result).toBe(20);
    });

    it("should work with MIN on open ranges", () => {
      const content = new Map<string, SerializedCellValue>([
        ["A1", "=MIN(B10:B)"],
        ["B10", 30],
        ["B11", 10],
        ["B12", 20],
      ]);

      engine.setSheetContent(sheetName, content);

      const result = cell("A1", true);
      expect(result).toBe(10);
    });

    it("should work with MAX on open ranges", () => {
      const content = new Map<string, SerializedCellValue>([
        ["A1", "=MAX(B10:B)"],
        ["B10", 30],
        ["B11", 10],
        ["B12", 20],
      ]);

      engine.setSheetContent(sheetName, content);

      const result = cell("A1");
      expect(result).toBe(30);
    });
  });

  describe("Cross-sheet open ranges", () => {
    const sheet2Name = "Sheet2";
    
    beforeEach(() => {
      engine.addSheet(sheet2Name);
    });

    it("should sum values from another sheet with open row range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:B)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B10", 10],
          ["B11", 20],
          ["B12", 30],
          ["B15", 40],
        ])
      );

      const result = cell("A1");
      expect(result).toBe(100);
    });

    it("should sum values from another sheet with open column range", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:10)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B10", 10],
          ["C10", 20],
          ["D10", 30],
          ["F10", 40],
        ])
      );

      const result = cell("A1");
      expect(result).toBe(100);
    });

    it("should handle cross-sheet spills intersecting with open ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:INFINITY)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["A10", "=SEQUENCE(3,3)"],
          ["B15", 100], // Direct value outside spill
        ])
      );

      const result = cell("A1");
      // SEQUENCE(3,3) spills A10:C12 with values 1-9
      // B10:C12 contains: 2,3,5,6,8,9 = 33
      // Plus B15 = 100
      // Total = 133
      expect(result).toBe(133);
    });

    it("should detect infinite spills from another sheet", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:D)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B100", "=SEQUENCE(INFINITY)"],
        ])
      );

      const result = cell("A1", true);
      expect(result).toBe("#REF!: Can not evaluate all cells over an infinite range");
    });

    it("should handle cross-sheet frontier candidates - top frontier", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:D)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["C8", "=SEQUENCE(5,1)"],
        ])
      );

      const result = cell("A1");
      // C8 spills to C8:C12, intersection with B10:D is C10:C12 (values 3,4,5) = 12
      expect(result).toBe(12);
    });

    it("should handle cross-sheet frontier candidates - left frontier", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:INFINITY)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["A10", "=SEQUENCE(1,5)"],
        ])
      );

      const result = cell("A1");
      // A10 spills to A10:E10, intersection with B10:INFINITY is B10:E10 (values 2,3,4,5) = 14
      expect(result).toBe(14);
    });

    it("should handle blocked cross-sheet frontier candidates", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!C10:INFINITY)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["A8", "=SEQUENCE(5,5)"],
          ["B9", 999], // Blocks the spill
        ])
      );

      const result = cell("A1");
      // A8 spill is blocked by B9, so no spilled values in range
      expect(result).toBe(0);
    });

    it("should handle mixed cross-sheet and same-sheet data", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B10:B) + SUM(Sheet2!B10:B)"],
          ["B10", 10],
          ["B11", 20],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B10", 100],
          ["B11", 200],
        ])
      );

      const result = cell("A1");
      expect(result).toBe(330); // (10+20) + (100+200)
    });

    it("should propagate cross-sheet errors in open ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:B)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B10", 10],
          ["B11", "=INVALID_FUNCTION()"],
          ["B12", 20],
        ])
      );

      const result = cell("A1");
      expect(result).toBe("#NAME?");
    });

    it("should handle cross-sheet circular references", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!A10:B)"],
          ["A10", "=Sheet2!B10"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B10", "=TestSheet!A10"],
        ])
      );

      const result = cell("A1");
      expect(result).toBe("#CYCLE!");
    });

    it("should handle references to non-existent sheets", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(NonExistentSheet!B10:INFINITY)"],
        ])
      );

      const result = cell("A1");
      expect(result).toBe("#REF!");
    });

    it("should handle complex cross-sheet spill chains", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:INFINITY)"],
          ["X1", "=SEQUENCE(2,2)"], // Creates spill on current sheet
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["A10", "=TestSheet!X1:Y2*10"], // References spill from TestSheet
          ["B15", 1000], // Direct value
        ])
      );

      // const result = cell("A1", true);
      // TestSheet!X1:Y2 is [1,2;3,4]
      // Sheet2!A10 becomes [10,20;30,40] (multiplied by 10)
      // B10:B11 from the spill = 20+40 = 60
      // Plus B15 = 1000
      // Total = 1060
      // expect(result).toBe(1060);
    });

    it("should work with multiple sheets and open ranges", () => {
      const sheet3Name = "Sheet3";
      engine.addSheet(sheet3Name);

      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(Sheet2!B10:B, Sheet3!C10:C)"],
        ])
      );
      
      engine.setSheetContent(
        sheet2Name,
        new Map<string, SerializedCellValue>([
          ["B10", 10],
          ["B11", 20],
        ])
      );
      
      engine.setSheetContent(
        sheet3Name,
        new Map<string, SerializedCellValue>([
          ["C10", 100],
          ["C11", 200],
        ])
      );

      const result = cell("A1");
      expect(result).toBe(330); // (10+20) + (100+200)
    });
  });

  describe("Edge cases", () => {
    it("should handle ranges starting at row/column 0", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["B1", "=SUM(A1:A)"],
        ])
      );

      const result = cell("B1");
      expect(result).toBe(30);
    });

    it("should be circular?", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B2", 20],
          ["Z100", 30],
          ["AA1", "=SUM(A1:INFINITY)"],
        ])
      );
      const result = cell("AA1");
      expect(result).toBe(FormulaError.CYCLE);
    });

    it("should handle entire sheet ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "=SUM(B1:INFINITY)"],
          ["B1", 10],
          ["B2", 20],
          ["Z100", 30],
        ])
      );
      const result = cell("A1", true);
      expect(result).toBe(60);
    });

    it("should handle nested spills with open ranges", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["X1", "=SEQUENCE(2,2)"],
          ["A10", "=X1:Y2*10"],
          ["A1", "=SUM(A10:B)"],
          ["A2", "shouldn't be evaluated"],
        ])
      );
      const result = cell("A1", true);
      // X1:Y2 is [1,2;3,4], multiplied by 10 = [10,20;30,40]
      // So A10:B11 = 10+20+30+40 = 100
      expect(result).toBe(100);
    });
    it("should handle nested spills with open ranges /2", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["X1", "=SEQUENCE(8,8)"],
          ["A10", "=X1:Y2"],
          ["A1", "=SUM(A10:B)"], // A1 can not be a frontier dependency of X1:Y2
          ["A2", "shouldn't be evaluated"],
        ])
      );
      const result = cell("A1", true);
      // X1:Y2 is [1,2;9,10]
      // So A10:B11 = 1+2+9+10 = 22
      expect(result).toBe(22);
    });
    it("should handle frontier dependencies", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A10", "=5:5"],
          ["B5", 2],
          ["D1", "=SUM(B:B)"],
          ["H5", 10]
        ])
      );
      const result = cell("D1", true);
      expect(result).toBe(4);
    });
  });
});
