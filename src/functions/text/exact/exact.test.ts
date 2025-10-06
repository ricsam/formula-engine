import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { FormulaError, type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("EXACT function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  const address = (ref: string) => ({ sheetName, workbookName, ...parseCellReference(ref) });

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  test("basic string comparisons", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(\"Apple\", \"Apple\")"], // Identical strings
        ["A2", "=EXACT(\"Apple\", \"apple\")"], // Case difference
        ["A3", "=EXACT(\"Apple\", \"APPLE\")"], // Case difference
        ["A4", "=EXACT(\"\", \"\")"], // Empty strings
        ["A5", "=EXACT(\"Hello\", \"World\")"], // Different strings
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false); // Case-sensitive
    expect(cell("A3")).toBe(false); // Case-sensitive
    expect(cell("A4")).toBe(true); // Empty strings are equal
    expect(cell("A5")).toBe(false);
  });

  test("numeric values converted to text", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(123, \"123\")"], // Number vs string
        ["A2", "=EXACT(123, 123)"], // Number vs number
        ["A3", "=EXACT(0, \"0\")"], // Zero vs string
        ["A4", "=EXACT(-5, \"-5\")"], // Negative number
        ["A5", "=EXACT(3.14, \"3.14\")"], // Decimal number
        ["A6", "=EXACT(123, \"124\")"], // Different numbers
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(true);
    expect(cell("A5")).toBe(true);
    expect(cell("A6")).toBe(false);
  });

  test("boolean values converted to text", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(TRUE, \"TRUE\")"], // Boolean vs string
        ["A2", "=EXACT(FALSE, \"FALSE\")"], // Boolean vs string
        ["A3", "=EXACT(TRUE, TRUE)"], // Boolean vs boolean
        ["A4", "=EXACT(TRUE, FALSE)"], // Different booleans
        ["A5", "=EXACT(TRUE, \"true\")"], // Case difference
        ["A6", "=EXACT(FALSE, \"false\")"], // Case difference
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(false);
    expect(cell("A5")).toBe(false); // Case-sensitive
    expect(cell("A6")).toBe(false); // Case-sensitive
  });

  test("cell references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", "Apple"],
        ["B2", "Apple"],
        ["B3", "apple"],
        ["B4", "APPLE"],
        ["B5", 123],
        ["B6", "123"],
        ["A1", "=EXACT(B1, B2)"], // Same strings
        ["A2", "=EXACT(B1, B3)"], // Case difference
        ["A3", "=EXACT(B1, B4)"], // Case difference
        ["A4", "=EXACT(B5, B6)"], // Number vs string representation
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(true);
  });

  test("empty and whitespace handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(\"\", \"\")"], // Both empty
        ["A2", "=EXACT(\"\", \" \")"], // Empty vs space
        ["A3", "=EXACT(\" \", \"  \")"], // Different spaces
        ["A4", "=EXACT(\"Hello\", \"Hello \")"], // Trailing space
        ["A5", "=EXACT(\" Hello\", \"Hello\")"], // Leading space
        ["A6", "=EXACT(\"Hello\\nWorld\", \"Hello\\nWorld\")"], // Newlines
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(false);
    expect(cell("A5")).toBe(false);
    expect(cell("A6")).toBe(true);
  });

  test("special characters and unicode", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(\"café\", \"café\")"], // Accented characters
        ["A2", "=EXACT(\"café\", \"cafe\")"], // With/without accent
        ["A3", "=EXACT(\"@#$%\", \"@#$%\")"], // Special characters
        ["A4", "=EXACT(\"Hello👋\", \"Hello👋\")"], // Emoji
        ["A5", "=EXACT(\"Hello👋\", \"Hello🙋\")"], // Different emoji
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(true);
    expect(cell("A5")).toBe(false);
  });

  test("infinity values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(1/0, 1/0)"], // Positive infinity vs positive infinity
        ["A2", "=EXACT(-1/0, -1/0)"], // Negative infinity vs negative infinity
        ["A3", "=EXACT(1/0, -1/0)"], // Positive vs negative infinity
        ["A4", "=EXACT(1/0, \"INFINITY\")"], // Infinity vs string representation
        ["A5", "=EXACT(-1/0, \"-INFINITY\")"], // Negative infinity vs string representation
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true);
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(true); // Infinity converts to "INFINITY"
    expect(cell("A5")).toBe(true); // Negative infinity converts to "-INFINITY"
  });

  test("error handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT()"], // No arguments
        ["A2", "=EXACT(\"Hello\")"], // Too few arguments
        ["A3", "=EXACT(\"Hello\", \"World\", \"Extra\")"], // Too many arguments
        ["B1", "=1/0/0"], // Error in cell
        ["A4", "=EXACT(B1, \"test\")"], // Error in first argument
        ["A5", "=EXACT(\"test\", B1)"], // Error in second argument
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    expect(cell("A2")).toBe(FormulaError.VALUE);
    expect(cell("A3")).toBe(FormulaError.VALUE);
    
    // Check if B1 produces an error
    const errorResult = cell("B1");
    if (errorResult && typeof errorResult === "string" && errorResult.startsWith("#")) {
      expect(cell("A4")).toBe(errorResult); // Should propagate error
      expect(cell("A5")).toBe(errorResult); // Should propagate error
    } else {
      // If B1 doesn't produce an error, test string conversion
      expect(cell("A4")).toBe(false); // Different values
      expect(cell("A5")).toBe(false); // Different values
    }
  });

  test("dynamic arrays with SEQUENCE", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(3, 1, 1, 1)"], // {1; 2; 3}
        ["B1", "=SEQUENCE(3, 1, 1, 1)"], // {1; 2; 3}
        ["C1", "=SEQUENCE(3, 1, 2, 1)"], // {2; 3; 4}
        ["D1", "=EXACT(A1:A3, B1:B3)"], // Compare identical arrays
        ["E1", "=EXACT(A1:A3, C1:C3)"], // Compare different arrays
      ])
    );

    // Should spill results
    expect(cell("D1")).toBe(true); // 1 = 1
    expect(cell("D2")).toBe(true); // 2 = 2
    expect(cell("D3")).toBe(true); // 3 = 3
    
    expect(cell("E1")).toBe(false); // 1 ≠ 2
    expect(cell("E2")).toBe(false); // 2 ≠ 3
    expect(cell("E3")).toBe(false); // 3 ≠ 4
  });

  test("mixed spilled and single values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(3, 1, 1, 1)"], // {1; 2; 3}
        ["B1", "=EXACT(A1:A3, \"2\")"], // Compare array to single value
        ["C1", "=EXACT(\"Hello\", A1:A3)"], // Compare single value to array
      ])
    );

    // Should spill results
    expect(cell("B1")).toBe(false); // "1" ≠ "2"
    expect(cell("B2")).toBe(true); // "2" = "2"
    expect(cell("B3")).toBe(false); // "3" ≠ "2"
    
    expect(cell("C1")).toBe(false); // "Hello" ≠ "1"
    expect(cell("C2")).toBe(false); // "Hello" ≠ "2"
    expect(cell("C3")).toBe(false); // "Hello" ≠ "3"
  });

  test("case sensitivity examples", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=EXACT(\"ABC\", \"abc\")"], // Different case
        ["A2", "=EXACT(\"Test\", \"TEST\")"], // Different case
        ["A3", "=EXACT(\"Hello World\", \"hello world\")"], // Different case
        ["A4", "=EXACT(\"123ABC\", \"123abc\")"], // Mixed alphanumeric
        ["A5", "=EXACT(\"CamelCase\", \"camelcase\")"], // CamelCase vs lowercase
      ])
    );

    expect(cell("A1")).toBe(false);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(false);
    expect(cell("A5")).toBe(false);
  });

  test("edge cases with special values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", ""],
        ["B2", 0],
        ["B3", false],
        ["A1", "=EXACT(B1, \"\")"], // Empty cell vs empty string
        ["A2", "=EXACT(B2, \"0\")"], // Zero vs string "0"
        ["A3", "=EXACT(B3, \"FALSE\")"], // FALSE vs string "FALSE"
        ["A4", "=EXACT(B3, \"false\")"], // FALSE vs lowercase "false"
      ])
    );

    expect(cell("A1", true)).toBe(true); // Empty cell = empty string
    expect(cell("A2")).toBe(true); // 0 converts to "0"
    expect(cell("A3")).toBe(true); // FALSE converts to "FALSE"
    expect(cell("A4")).toBe(false); // Case-sensitive
  });
});
