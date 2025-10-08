import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from "../../../../src/core/engine";
import { visualizeSpreadsheet } from "../../../../src/core/utils/spreadsheet-visualizer";
import { type SerializedCellValue } from "../../../../src/core/types";

describe("Spreadsheet Visualizer", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("should create simple visualization without headers or row numbers", () => {
      // Set up some test data
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Name"],
          ["B1", "Age"],
          ["C1", "City"],
          ["A2", "Alice"],
          ["B2", 25],
          ["C2", "NYC"],
          ["A3", "Bob"],
          ["B3", 30],
          ["C3", "LA"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 3,
        numCols: 3,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toMatchInlineSnapshot(`
        "Name  | Age | City
        Alice | 25  | NYC 
        Bob   | 30  | LA  
        "
      `);
    });

    test("should create formatted table with headers and row numbers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Name"],
          ["B1", "Age"],
          ["A2", "Alice"],
          ["B2", 25],
          ["A3", "Bob"],
          ["B3", 30],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 3,
        numCols: 2,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(result).toMatchInlineSnapshot(`
        "    | A     | B  
        ----+-------+----
          1 | Name  | Age
          2 | Alice | 25 
          3 | Bob   | 30 
        "
      `);
    });

    test("should handle empty cells with default character", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Data"],
          ["C1", "More"],
          ["A2", "Test"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 2,
        numCols: 3,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toContain("Data");
      expect(result).toContain("More");
      expect(result).toContain("Test");
      expect(result).toContain("."); // Empty cells
    });

    test("should handle custom empty cell character", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Data"],
          ["C1", "More"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 3,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        emptyCellChar: "-",
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toContain("Data");
      expect(result).toContain("More");
      expect(result).toContain("-"); // Custom empty cell character
    });

    test("should handle custom start position", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Skip"],
          ["B1", "Skip"],
          ["C1", "Start"],
          ["D1", "Here"],
          ["C2", "Row2"],
          ["D2", "Data"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 2,
        numCols: 2,
        startRow: 1,
        startCol: 2, // Start from column C
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toContain("Start");
      expect(result).toContain("Here");
      expect(result).toContain("Row2");
      expect(result).toContain("Data");
      expect(result).not.toContain("Skip");
    });
  });

  describe("formatting options", () => {
    test("should handle long cell values with truncation", () => {
      engine.setSheetContent(
          sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "This is a very long cell value that should be truncated"],
          ["B1", "Short"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 2,
        maxColWidth: 10,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(result).toContain("...");
      expect(result).toContain("Short");
    });

    test("should respect minimum column width", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "A"],
          ["B1", "B"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 2,
        minColWidth: 5,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      // Each column should be at least 5 characters wide
      const lines = result.split("\n");
      const dataLine = lines.find(line => line.includes("A") && !line.includes("-"));
      expect(dataLine).toBeDefined();
      
      // Should have proper spacing due to minimum width
      expect(dataLine!.length).toBeGreaterThan(10); // At least 5 + 5 + separators
    });

    test("should work without column headers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Data1"],
          ["B1", "Data2"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 2,
        showColumnHeaders: false,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(result).not.toContain("A");
      expect(result).not.toContain("B");
      expect(result).toContain("Data1");
      expect(result).toContain("Data2");
      expect(result).not.toContain("---"); // No header separator
    });

    test("should work without row numbers", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Data1"],
          ["B1", "Data2"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 2,
        showRowNumbers: false,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(result).toContain("Data1");
      expect(result).toContain("Data2");
      // Should not start with numbers
      const lines = result.split("\n").filter(line => line.trim());
      const dataLines = lines.filter(line => !line.includes("-"));
      expect(dataLines[1]).not.toMatch(/^\s*\d+/);
    });

    test("should handle different start positions", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Skip"],
          ["B1", "Skip"],
          ["C2", "Start"],
          ["D2", "Here"],
          ["C3", "Row2"],
          ["D3", "Data"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 2,
        numCols: 2,
        startRow: 2,
        startCol: 2, // Start from column C
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(result).toContain("Start");
      expect(result).toContain("Here");
      expect(result).toContain("Row2");
      expect(result).toContain("Data");
      expect(result).not.toContain("Skip");
    });
  });

  describe("advanced features", () => {
    test("should work with formulas and calculated values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 20],
          ["C1", "=A1+B1"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 3,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      expect(result).toContain("10");
      expect(result).toContain("20");
      expect(result).toContain("30"); // Calculated result
    });

    test("should handle multiple sheets", () => {
      const sheet2 = "Sheet2";
      engine.addSheet({ workbookName, sheetName: sheet2 });
      
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", "Sheet1Data"]])
      );
      
      engine.setSheetContent(
        { workbookName, sheetName: sheet2 },
        new Map<string, SerializedCellValue>([["A1", "Sheet2Data"]])
      );

      const result1 = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 1,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      const result2 = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 1,
        sheetName: sheet2,
        workbookName: workbookName,
      });

      expect(result1).toContain("Sheet1Data");
      expect(result2).toContain("Sheet2Data");
    });

    test("should handle sparse data with varying column widths", () => {
      // Set up sparse data with different content lengths
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          // Row 1: Headers with different lengths
          ["A1", "ID"],
          ["B1", "Very Long Product Name"],
          ["C1", "Price"],
          ["D1", "Category"],
          ["E1", "Description"],
          
          // Row 2: Short content in some columns
          ["A2", "1"],
          ["C2", 99.99],
          ["E2", "Basic item"],
          
          // Row 3: Long content in different columns  
          ["A3", "2"],
          ["B3", "Super Ultra Mega Premium Product"],
          ["D3", "Electronics"],
          
          // Row 4: Mixed content lengths
          ["A4", "3"],
          ["B4", "Widget"],
          ["C4", 1234.56],
          ["D4", "Tools"],
          ["E4", "Professional grade tool for advanced users"],
          
          // Row 6: Sparse row with only some columns filled
          ["B6", "Gadget"],
          ["D6", "Home"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 6,
        numCols: 5,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        maxColWidth: 25, // Allow longer columns
      });

      expect(result).toMatchInlineSnapshot(`
        "    | A   | B                         | C       | D           | E                        
        ----+-----+---------------------------+---------+-------------+--------------------------
          1 | ID  | Very Long Product Name    | Price   | Category    | Description              
          2 | 1   | .                         | 99.99   | .           | Basic item               
          3 | 2   | Super Ultra Mega Premi... | .       | Electronics | .                        
          4 | 3   | Widget                    | 1234.56 | Tools       | Professional grade too...
          5 | .   | .                         | .       | .           | .                        
          6 | .   | Gadget                    | .       | Home        | .                        
        "
      `);
    });

    test("should handle very sparse data with mostly empty cells", () => {
      // Set up very sparse data
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Corner"],
          ["E1", "Far"],
          ["A5", "Bottom"],
          ["C3", "Center"],
          ["E5", "End"],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 5,
        numCols: 5,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      // With consistent row widths, each row will be padded to the same length
      // The result will have proper column alignment
      expect(result).toContain("Corner");
      expect(result).toContain("Far");
      expect(result).toContain("Bottom");
      expect(result).toContain("Center");
      expect(result).toContain("End");
      
      // Verify consistent row widths
      const lines = result.split('\n').filter(line => line.length > 0);
      if (lines.length > 1) {
        const firstLineLength = lines[0]?.length;
        for (let i = 1; i < lines.length; i++) {
          expect(lines[i]?.length).toBe(firstLineLength!);
        }
      }
    });

    test("should ensure all rows have exactly the same character width", () => {
      // Set up data with varying content lengths to test consistent row width
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Short"],
          ["B1", "Very Long Header Name"],
          ["C1", "X"],
          ["A2", "A"],
          ["B2", "Medium"],
          ["C2", "Very Long Content Here"],
          ["A3", ""],
          ["B3", ""],
          ["C3", ""],
        ])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 3,
        numCols: 3,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
      });

      // Split into lines and check that all lines have the same length
      const lines = result.split('\n').filter(line => line.length > 0); // Remove empty lines
      const lineLengths = lines.map(line => line.length);
      
      // All lines should have the same length
      const firstLineLength = lineLengths[0];
      for (let i = 1; i < lineLengths.length; i++) {
        expect(lineLengths[i]).toBe(firstLineLength!);
      }
      
      // Should have at least 3 lines (header + separator + data rows)
      expect(lines.length).toBeGreaterThanOrEqual(3);
    });
  });

  describe("edge cases", () => {
    test("should handle empty engine", () => {
      const emptyEngine = FormulaEngine.buildEmpty();
      emptyEngine.addWorkbook(workbookName);
      
      const result = visualizeSpreadsheet(emptyEngine, {
        numRows: 2,
        numCols: 2,
        workbookName: workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toContain(".");
    });

    test("should handle zero dimensions", () => {
      const result = visualizeSpreadsheet(engine, {
        numRows: 0,
        numCols: 2,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toBe("");
    });

    test("should handle single cell", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([["A1", "Single"]])
      );

      const result = visualizeSpreadsheet(engine, {
        numRows: 1,
        numCols: 1,
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        showColumnHeaders: false,
        showRowNumbers: false,
      });

      expect(result).toContain("Single");
    });
  });
});