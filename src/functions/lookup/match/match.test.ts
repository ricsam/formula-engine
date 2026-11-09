import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("MATCH function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue(
      { sheetName, workbookName, ...parseCellReference(ref) },
      debug
    );

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const address = (ref: string) => ({ sheetName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName: workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  test("structured reference position bug", () => {
    // Create a table starting at B19 with headers
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Table headers at row 19
        ["B19", "Name"],
        ["C19", "Type"],
        ["D19", "Identifier"], // Column we'll search in
        // Table data starting at row 20
        ["B20", "Car 1-A"],
        ["C20", "Vehicle"],
        ["D20", "Car 1-A"],
        ["B21", "Car 1-B"],
        ["C21", "Vehicle"],
        ["D21", "Car 1-B"],
        // ... more rows ...
        // Fill in more data rows to reach row 28
        ["B22", "Car 1-C"],
        ["C22", "Vehicle"],
        ["D22", "Car 1-C"],
        ["B23", "Car 1-D"],
        ["C23", "Vehicle"],
        ["D23", "Car 1-D"],
        ["B24", "Car 1-E"],
        ["C24", "Vehicle"],
        ["D24", "Car 1-E"],
        ["B25", "Car 1-F"],
        ["C25", "Vehicle"],
        ["D25", "Car 1-F"],
        ["B26", "Car 1-G"],
        ["C26", "Vehicle"],
        ["D26", "Car 1-G"],
        ["B27", "Car 1-H"],
        ["C27", "Vehicle"],
        ["D27", "Car 1-H"],
        ["B28", "Car 2-A"], // This is the row we're looking for
        ["C28", "Vehicle"],
        ["D28", "Car 2-A"], // This cell should match
        // Test formula
        ["F1", "=MATCH(\"Car 2-A\", INPUT_LAYOUT[Identifier], 0)"],
      ])
    );

    // Define the table
    engine.addTable({
      tableName: "INPUT_LAYOUT",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "B19",
      numRows: { type: "number", value: 10 }, // 10 data rows (plus 1 header)
      numCols: 3, // 3 columns: Name, Type, Identifier
    });

    // The match should return 9 (relative to table start: D28 is 9th row in Identifier column)
    // NOT 29 (absolute sheet position)
    // D19 = header (position 0), D20 = data row 1, ..., D28 = data row 9
    expect(cell("F1", true)).toBe(9);
  });

  test("structured reference vs absolute position bug reproduction", () => {
    // This test specifically reproduces the bug where MATCH returns 29 instead of 9
    // Create a scenario with sparse data to see if the issue occurs
    
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Table headers at row 19
        ["B19", "Name"],
        ["C19", "Type"],
        ["D19", "Identifier"], // Column we'll search in
        // Only put data in a few specific rows, including the target row
        ["D20", "Car 1-A"], // Row 1 of data
        ["D21", "Car 1-B"], // Row 2 of data
        ["D28", "Car 2-A"], // Row 9 of data (but only 3rd actual data cell)
        // Test formula
        ["F1", "=MATCH(\"Car 2-A\", INPUT_LAYOUT[Identifier], 0)"],
      ])
    );

    // Define the table with full range but sparse data
    engine.addTable({
      tableName: "INPUT_LAYOUT",
      sheetName: sheetAddress.sheetName,
      workbookName: sheetAddress.workbookName,
      start: "B19",
      numRows: { type: "number", value: 10 }, // 10 data rows (plus 1 header)
      numCols: 3, // 3 columns: Name, Type, Identifier
    });

    // If the bug exists, this might return 3 (position in sparse data) 
    // instead of 9 (actual row position relative to table start)
    const result = cell("F1", true);
    
    // The correct behavior should be to return the position relative to the table data start
    // D28 is at absolute row 28, table data starts at row 20 (row 19 is header)
    // So D28 should be position 9 in the table data (28 - 19 = 9)
    // With the fix, it should now return 9 (actual position) instead of 3 (sparse position)
    expect(result).toBe(9); // Fixed: should return actual position within range
  });

  describe("basic functionality", () => {
    test("should find exact match with match_type 0", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Banana", A1:A3, 0)');

      expect(cell("B1")).toBe(2);
    });

    test("should find exact match for numbers", () => {
      // Use SerializedCellValue to set numeric values properly
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
        ])
      );
      setCellContent("B1", "=MATCH(20, A1:A3, 0)");

      expect(cell("B1")).toBe(2);
    });

    test("should default to match_type 1", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Banana", A1:A3)');

      expect(cell("B1")).toBe(2);
    });

    test("should find the largest value less than lookup_value for match_type 1", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["A4", 40],
          ["B1", "=MATCH(25, A1:A4, 1)"],
        ])
      );

      expect(cell("B1")).toBe(2);
    });

    test("should find the smallest value greater than lookup_value for match_type -1", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 40],
          ["A2", 30],
          ["A3", 20],
          ["A4", 10],
          ["B1", "=MATCH(25, A1:A4, -1)"],
        ])
      );

      expect(cell("B1")).toBe(2);
    });

    test("should return the last duplicate for approximate matches", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 20],
          ["A4", 30],
          ["B1", "=MATCH(20, A1:A4, 1)"],
          ["C1", 30],
          ["C2", 20],
          ["C3", 20],
          ["C4", 10],
          ["D1", "=MATCH(20, C1:C4, -1)"],
        ])
      );

      expect(cell("B1")).toBe(3);
      expect(cell("D1")).toBe(3);
    });

    test("should return #N/A when no approximate candidate exists", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["B1", "=MATCH(5, A1:A2, 1)"],
          ["C1", 20],
          ["C2", 10],
          ["D1", "=MATCH(25, C1:C2, -1)"],
        ])
      );

      expect(cell("B1")).toBe(FormulaError.NA);
      expect(cell("D1")).toBe(FormulaError.NA);
    });

    test("should search unsorted data the way Excel does", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 30],
          ["A3", 20],
          ["B1", "=MATCH(25, A1:A3, 1)"],
          ["C1", 30],
          ["C2", 10],
          ["C3", 20],
          ["D1", "=MATCH(25, C1:C3, -1)"],
        ])
      );

      // The middle value sends the search to the first cell. Excel does not
      // check the order and returns the same positions.
      expect(cell("B1")).toBe(1);
      expect(cell("D1")).toBe(1);
    });

    test("should perform case-insensitive approximate text matching", () => {
      setCellContent("A1", "apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "cherry");
      setCellContent("B1", '=MATCH("Blueberry", A1:A3, 1)');

      expect(cell("B1")).toBe(2);
    });

    test("should return a column position for a horizontal approximate match", () => {
      setCellContent("A1", 10);
      setCellContent("B1", 20);
      setCellContent("C1", 30);
      setCellContent("D1", "=MATCH(25, A1:C1, 1)");

      expect(cell("D1")).toBe(2);
    });

    test("should return #N/A when not found", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Grape", A1:A3, 0)');

      expect(cell("B1")).toBe(FormulaError.NA);
    });

    test("should be case-insensitive (matching Excel behavior)", () => {
      setCellContent("A1", "apple");
      setCellContent("A2", "Banana");
      setCellContent("A3", "Cherry");
      setCellContent("B1", '=MATCH("Apple", A1:A3, 0)');

      expect(cell("B1")).toBe(1);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return #VALUE! for boolean lookup_value", () => {
      setCellContent("A1", "TRUE");
      setCellContent("A2", "FALSE");
      setCellContent("B1", "=MATCH(TRUE, A1:A2, 0)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for infinity lookup_value", () => {
      setCellContent("A1", "1");
      setCellContent("A2", "2");
      setCellContent("B1", "=MATCH(INFINITY, A1:A2, 0)");

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for invalid match_type", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, 2)');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for string match_type", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, "exact")');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });
  });

  describe("error handling", () => {
    test("should return #VALUE! for wrong number of arguments", () => {
      setCellContent("B1", "=MATCH()");
      expect(cell("B1")).toBe(FormulaError.VALUE);

      setCellContent("B2", '=MATCH("Apple")');
      expect(cell("B2")).toBe(FormulaError.VALUE);
    });

    test("should return #VALUE! for too many arguments", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=MATCH("Apple", A1:A1, 0, "extra")');

      expect(cell("B1")).toBe(FormulaError.VALUE);
    });

    test("should handle decimal match_type (should floor)", () => {
      setCellContent("A1", "Apple");
      setCellContent("A2", "Banana");
      setCellContent("B1", '=MATCH("Apple", A1:A2, 0.9)');

      expect(cell("B1")).toBe(1);
    });
  });

  describe("edge cases", () => {
    test("should handle single cell array", () => {
      setCellContent("A1", "Apple");
      setCellContent("B1", '=MATCH("Apple", A1, 0)');

      expect(cell("B1")).toBe(1);
    });

    test("should handle mixed string and number types (strict checking)", () => {
      // Set up mixed array with proper types
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", 10],
        ])
      );
      setCellContent("B1", "=MATCH(10, A1:A2, 0)"); // Number lookup in mixed array

      expect(cell("B1")).toBe(2);
    });

    test("should reject 2D arrays (must be 1D)", () => {
      // Set up a 2D array (multiple rows AND multiple columns)
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "Cherry"],
          ["A2", "Dog"],
          ["B2", "Elephant"],
          ["C2", "Fox"],
        ])
      );
      
      // Try to use MATCH with a 2D range (2 rows × 3 columns)
      setCellContent("D1", '=MATCH("Elephant", A1:C2, 0)');
      
      const result = cell("D1");
      // Should return #VALUE! error
      expect(result).toBe("#VALUE!");
    });

    test("should accept horizontal 1D arrays (single row)", () => {
      // Set up a horizontal array (1 row × multiple columns)
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["B1", "Banana"],
          ["C1", "Cherry"],
        ])
      );
      
      setCellContent("D1", '=MATCH("Cherry", A1:C1, 0)');
      expect(cell("D1")).toBe(3); // Third column
    });

    test("should accept vertical 1D arrays (single column)", () => {
      // Set up a vertical array (multiple rows × 1 column)
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Apple"],
          ["A2", "Banana"],
          ["A3", "Cherry"],
        ])
      );
      
      setCellContent("B1", '=MATCH("Cherry", A1:A3, 0)');
      expect(cell("B1")).toBe(3); // Third row
    });
  });

  describe("can use table column as lookup_array", () => {
    test("should find exact match with match_type 0", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Fruit"],
          ["A2", "Stock"],
          ["A3", "Is local"],
          ["A2", "Apple"],
          ["B2", 2],
          ["C2", "Yes"],
          ["A2", "Orange"],
          ["B2", 3],
          ["C2", "No"],
          ["A3", "Banana"],
          ["B3", 1],
          ["C3", "Maybe"],
          ["A4", "Kiwi"],
          ["B4", 4],
          ["C4", "Yes"],
          ["A5", "Pineapple"],
          ["B5", 5],
          ["C5", "No"],
          ["A6", "Pear"],
          ["B6", 6],
          ["C6", "Yes"],
          ["A7", "Strawberry"],
          ["B7", 7],
          ["C7", "No"],
          ["A8", "Watermelon"],
          ["B8", 8],
          ["C8", "Yes"],
          ["A9", "Mango"],
          ["B9", 9],
          ["C9", "No"],
          ["A10", "Pomegranate"],

          ["K1", `=MATCH("Kiwi", A2:A10, 0)`],
          ["L1", `=MATCH("Mango", A:A, 0)`],
        ])
      );

      // expect(cell("K1")).toBe(3);
      expect(cell("L1", true)).toBe(9);
    });
  });

  describe("structured table references", () => {
    test("should work with table column references across sheets", () => {
      // Create a separate sheet for the ORDERinput table
      const inputSheetName = "InputSheet";
      engine.addSheet({ workbookName, sheetName: inputSheetName });

      // Set up data on the input sheet FIRST
      engine.setSheetContent(
        { workbookName, sheetName: inputSheetName },
        new Map<string, SerializedCellValue>([
          ["A1", "OrderID"],
          ["B1", "Amount"],
          ["A2", "ORD-001"],
          ["B2", 100],
          ["A3", "ORD-002"],
          ["B3", 200],
          ["A4", "ORD-003"],
          ["B4", 300],
          ["A5", "ORD-004"],
          ["B5", 400],
          ["A6", "ORD-005"],
          ["B6", "=INVALID_FUNCTION()"],
        ])
      );

      // Set up data on the main sheet FIRST (only CurrentTable data)
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          // CurrentTable data - expand to 3 columns to include the formula column
          ["D7", "OrderID"],
          ["E7", "Status"],
          ["F7", "MatchResult"],
          ["D8", "ORD-002"],
          ["E8", "PENDING"],
          ["F8", `=MATCH([@[OrderID]], ORDERinput[OrderID], 0)`], // Should find ORD-002 at position 2
          ["D9", "ORD-004"],
          ["E9", "COMPLETE"],
          ["F9", `=MATCH([@[OrderID]], ORDERinput[OrderID], 0)`], // Should find ORD-004 at position 4

          // Test the cross-sheet table reference separately
          ["G1", `=ORDERinput[OrderID]`], // Test if cross-sheet structured reference works
        ])
      );

      // NOW create the tables after the sheet content is set
      // Create the ORDERinput table on the input sheet with infinite rows
      engine.addTable({
        tableName: "ORDERinput",
        sheetName: inputSheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "infinity", sign: "positive" }, // Infinite rows
        numCols: 2, // OrderID and Amount columns
      });

      // Create another table on the main sheet for the current row reference with infinite rows
      engine.addTable({
        tableName: "CurrentTable",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "D7",
        numRows: { type: "infinity", sign: "positive" }, // Infinite rows
        numCols: 3, // OrderID, Status, and MatchResult columns
      });

      expect(cell("F8", true)).toBe(2); // ORD-002 is at position 2 in ORDERinput
      expect(cell("F9", true)).toBe(4); // ORD-004 is at position 4 in ORDERinput
    });

    test("should handle empty table columns gracefully", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "OrderID"],
          ["B1", `=MATCH("TEST", EmptyTable[OrderID], 0)`],
        ])
      );

      // Create an empty table
      engine.addTable({
        tableName: "EmptyTable",
        sheetName: sheetAddress.sheetName,
        workbookName: sheetAddress.workbookName,
        start: "A1",
        numRows: { type: "number", value: 0 }, // No data rows, just header
        numCols: 1,
      });

      expect(cell("B1")).toBe(FormulaError.NA); // Should return error for empty table
    });
  });

  describe("approximate match follows Excel's binary search", () => {
    // Expected values are Excel's results, from the sources named in each test.

    test("finds Excel's position on unsorted text when match_type is left out", () => {
      // The workbook from the issue left out the `,0` of an exact match.
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "MAP-994"],
          ["A2", "MAP-10977"],
          ["A3", "MAP-320"],
          ["B1", '=MATCH("MAP-994", A1:A3)'],
          ["B2", '=MATCH("MAP-994", A1:A3, 1)'],
          ["B3", '=MATCH("MAP-10977", A1:A3)'],
          ["B4", '=MATCH("MAP-320", A1:A3)'],
          ["B5", '=MATCH("b", {"a";"b";"c"})'],
        ])
      );

      expect(cell("B1")).toBe(3);
      expect(cell("B2")).toBe(3);
      expect(cell("B3")).toBe(2);
      expect(cell("B4")).toBe(3);
      expect(cell("B5")).toBe(2);
    });

    test("returns Excel's rows for INDEX/MATCH over an unsorted table column", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "MAP ID"],
          ["B1", "MAP Payload"],
          ["A2", "MAP-994"],
          ["B2", "payload 994"],
          ["A3", "MAP-10977"],
          ["B3", "payload 10977"],
          ["A4", "MAP-320"],
          ["B4", "payload 320"],
          ["D1", "MAP ID"],
          ["E1", "Payload"],
          ["D2", "MAP-994"],
          ["D3", "MAP-10977"],
          ["D4", "MAP-320"],
          ...["E2", "E3", "E4"].map(
            (ref): [string, SerializedCellValue] => [
              ref,
              "=INDEX(MAPcondition[MAP Payload],MATCH([@[MAP ID]],MAPcondition[MAP ID]),0)",
            ]
          ),
        ])
      );
      engine.addTable({
        tableName: "MAPcondition",
        sheetName,
        workbookName,
        start: "A1",
        numRows: { type: "number", value: 3 },
        numCols: 2,
      });
      engine.addTable({
        tableName: "Arranged",
        sheetName,
        workbookName,
        start: "D1",
        numRows: { type: "number", value: 3 },
        numCols: 2,
      });

      expect(cell("E2")).toBe("payload 320");
      expect(cell("E3")).toBe("payload 10977");
      expect(cell("E4")).toBe("payload 320");
    });

    test("reproduces Excel's VLOOKUP(..., TRUE) results on an unsorted table", () => {
      // Excel 2010 results posted on the MrExcel forum ("Understanding Binary
      // Search", 2015): VLOOKUP(A2, $E$2:$F$18, 2, TRUE) over unsorted names.
      const table = [
        ["Yvonne", "Orlando"],
        ["Freddi", "Gdansk"],
        ["Danielle", "Buffalo"],
        ["Samantha", "Dallas"],
        ["Karen", "Seattle"],
        ["Bette", "Mexico City"],
        ["Kate", "Milwaukee"],
        ["Jo", "Austin"],
        ["Lou", "Montevideo"],
        ["Lee", "Pedacuh"],
        ["Rudy", "LaCrosse"],
        ["Zed", "Baltimore"],
        ["Thomas", "Los Angeles"],
        ["Todd", "Jacksonville"],
        ["Harrison", "Leipzig"],
        ["Craig", "Miami"],
        ["Maurice", "Montreal"],
      ];
      const excelResults = [
        ["Samantha", "LaCrosse"],
        ["Craig", FormulaError.NA],
        ["Kate", "Buffalo"],
        ["Jo", "Buffalo"],
        ["Rudy", "LaCrosse"],
        ["Lou", "Montevideo"],
        ["Todd", "Montreal"],
        ["Harrison", "Buffalo"],
        ["Yvonne", "Montreal"],
        ["Bette", FormulaError.NA],
        ["Lee", "Buffalo"],
        ["Thomas", "Los Angeles"],
        ["Zed", "Montreal"],
        ["Maurice", "Pedacuh"],
        ["Danielle", FormulaError.NA],
        ["Karen", "Buffalo"],
        ["Freddi", "Gdansk"],
        ["Xavier", "Montreal"],
      ];
      const content = new Map<string, SerializedCellValue>();
      table.forEach(([name, city], index) => {
        content.set(`E${index + 2}`, name);
        content.set(`F${index + 2}`, city);
      });
      excelResults.forEach(([name], index) => {
        content.set(`A${index + 2}`, name);
        content.set(
          `B${index + 2}`,
          `=INDEX($F$2:$F$18, MATCH(A${index + 2}, $E$2:$E$18))`
        );
      });
      engine.setSheetContent(sheetAddress, content);

      expect(
        excelResults.map((_, index) => cell(`B${index + 2}`))
      ).toEqual(excelResults.map(([, city]) => city));
    });

    test("moves on from an exact hit while the next values are equal", () => {
      // Microsoft Q&A, "Binary search explain in Lookup with duplicates?"
      // (2023), and the unsorted tables of Apache POI's lookup test workbook,
      // whose expected values are cached by Excel.
      const cases: [string, number][] = [
        ["=MATCH(1,{1,1,1,1,1,0,1})", 5],
        ["=MATCH(1,{1,1,1,1,1,0,1,-1,1,1,1,1,1,0,1})", 13],
        ["=MATCH(1,{1,1,1,1,1,0,1,2,1,1,1,1,1,0,1})", 5],
        ["=MATCH(1,{1;1;1;1;1})", 5],
        ["=MATCH(33,{33;33;33;100;33;33;33;33;33;33})", 10],
        ["=MATCH(33,{33;33;33;33;100;33;33;33;33;33})", 4],
        ["=MATCH(33,{33;33;33;33;33;33;5;33;100;33})", 6],
        ["=MATCH(33,{32;32;32;31;100;32;32;32;32;32})", 4],
        [
          "=MATCH(33,{33;33;33;33;33;33;33;33;33;33;8;33;33;6;33;33;4;33;33;33;33})",
          16,
        ],
        [
          "=MATCH(10,{33;33;33;33;33;33;33;33;33;33;8;33;33;6;33;33;4;33;33;33;33})",
          11,
        ],
        [
          "=MATCH(33,{33;33;33;33;10;33;33;33;5;33;100;33;33;33;33;33;33;33;33;33;22})",
          8,
        ],
      ];
      engine.setSheetContent(
        sheetAddress,
        new Map(cases.map(([formula], index) => [`A${index + 1}`, formula]))
      );

      expect(cases.map((_, index) => cell(`A${index + 1}`))).toEqual(
        cases.map(([, expected]) => expected)
      );
    });

    test("steps over blank cells", () => {
      // MrExcel forum, "MATCH Omega Text Issue" (2016): Excel returned 4 and 6
      // with A9:A10 filled, 15 and 17 with them blank.
      const column: (string | undefined)[] = [
        "w", "x", "y", "z", "zz", "Ω", "ΩΩ", undefined, "ΩΩ",
        "ΩΩ", undefined, "w", "x", "y", "z", "zz", "Ω", "ΩΩ",
      ];
      const fill = (values: (string | undefined)[]) => {
        const content = new Map<string, SerializedCellValue>([
          ["B1", '=MATCH("z", A1:A18)'],
          ["B2", '=MATCH("Ω", A1:A18)'],
        ]);
        values.forEach((value, index) => {
          if (value !== undefined) content.set(`A${index + 1}`, value);
        });
        engine.setSheetContent(sheetAddress, content);
      };

      fill(column);
      expect(cell("B1")).toBe(4);
      expect(cell("B2")).toBe(6);

      fill(column.map((value, index) => (index === 8 || index === 9 ? undefined : value)));
      expect(cell("B1")).toBe(15);
      expect(cell("B2")).toBe(17);
    });

    test("finds the last value of a type in a whole column", () => {
      // Excel's results: A1 "z" and A10 "a" give 10 (Microsoft Tech Community,
      // "MATCH vs XMATCH: Binary Search to Find the Last Entry within a
      // Range", 2023); the last number or text of a column, whatever its
      // order (Excel Evolution, "Lookups - understanding the binary search
      // algorithm", 2016).
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "z"],
          ["A10", "a"],
          ["B1", 54],
          ["B2", 3],
          ["B3", 8],
          ["C1", "Yellow"],
          ["C2", "Green"],
          ["C3", "Blue"],
          ["C4", "Black"],
          ["C5", "Purple"],
          ["C6", "White"],
          ["E1", '=MATCH("Ω", A:A)'],
          ["E2", "=MATCH(9.99E+307, B:B)"],
          ["E3", '=MATCH("zzzz", C:C)'],
        ])
      );

      expect(cell("E1")).toBe(10);
      expect(cell("E2")).toBe(3);
      expect(cell("E3")).toBe(6);
    });

    test("steps over errors and values of another type", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", "one"],
          ["A3", 2],
          ["A4", "=TRUE"],
          ["A5", "=INVALID_FUNCTION()"],
          ["A6", 3],
          ["A7", "three"],
          ["B1", "=MATCH(9.99E+307, A1:A7)"],
          ["B2", '=MATCH("zzz", A1:A7)'],
          ["B3", "=MATCH(2, A1:A7)"],
          ["B4", '=MATCH("one", A1:A7)'],
          ["B5", '=MATCH("a", A1:A7)'],
        ])
      );

      expect(cell("B1")).toBe(6);
      expect(cell("B2")).toBe(7);
      expect(cell("B3")).toBe(3);
      expect(cell("B4")).toBe(2);
      expect(cell("B5")).toBe(FormulaError.NA);
    });

    test("searches descending data with match_type -1", () => {
      // Microsoft Tech Community, "XMATCH and MATCH return different result for
      // approximate search lookup_array with repetitive values" (2023).
      setCellContent("A1", "=MATCH(0, {3,2,2}, -1)");
      setCellContent("A2", "=MATCH(4, {3,2,2}, -1)");
      setCellContent("A3", "=MATCH(2.5, {3,2,2}, -1)");

      expect(cell("A1")).toBe(3);
      expect(cell("A2")).toBe(FormulaError.NA);
      expect(cell("A3")).toBe(1);
    });

    test("searches a row by column position", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "MAP-994"],
          ["B1", "MAP-10977"],
          ["D1", "MAP-320"],
          ["A2", '=MATCH("MAP-994", A1:D1)'],
          ["A3", '=MATCH("MAP-994", 1:1)'],
        ])
      );

      expect(cell("A2")).toBe(4);
      expect(cell("A3")).toBe(4);
    });

    test("recalculates when a cell of the lookup array changes", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "MAP-994"],
          ["A2", "MAP-10977"],
          ["A3", '="MAP-" & "320"'],
          ["B1", '=MATCH("MAP-994", A1:A3)'],
        ])
      );
      expect(cell("B1")).toBe(3);

      setCellContent("A2", "MAP-999");
      expect(cell("B1")).toBe(1);
    });
  });

  describe("omitted arguments", () => {
    beforeEach(() => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["A2", 20],
          ["A3", 30],
          ["B1", 5],
          ["B2", 0],
          ["B3", 7],
        ])
      );
    });

    test("an omitted match_type is 0, an exact match", () => {
      setCellContent("D1", "=MATCH(25, A1:A3,)");
      setCellContent("D2", "=MATCH(20, A1:A3,)");
      // A missing match_type is 1, an approximate match.
      setCellContent("D3", "=MATCH(25, A1:A3)");

      expect(cell("D1")).toBe(FormulaError.NA);
      expect(cell("D2")).toBe(2);
      expect(cell("D3")).toBe(2);
    });

    test("an omitted lookup_value is 0", () => {
      setCellContent("D1", "=MATCH(, B1:B3, 0)");
      setCellContent("D2", "=MATCH(, B1:B3,)");
      setCellContent("D3", "=MATCH(, A1:A3, 0)");

      expect(cell("D1")).toBe(2);
      expect(cell("D2")).toBe(2);
      expect(cell("D3")).toBe(FormulaError.NA);
    });

    test("recalculates with the lookup array", () => {
      setCellContent("D1", "=MATCH(, B1:B3,)");
      expect(cell("D1")).toBe(2);

      setCellContent("B2", 1);
      setCellContent("B3", 0);
      expect(cell("D1")).toBe(3);
    });
  });
});
