import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("OR function", () => {
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

  test("basic boolean values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(TRUE, TRUE)"], // Both true
        ["A2", "=OR(TRUE, FALSE)"], // One true
        ["A3", "=OR(FALSE, FALSE)"], // Both false
        ["A4", "=OR(TRUE)"], // Single true
        ["A5", "=OR(FALSE)"], // Single false
        ["A6", "=OR(FALSE, TRUE)"], // True as second argument
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND - returns TRUE if any is TRUE
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(true);
    expect(cell("A5")).toBe(false);
    expect(cell("A6")).toBe(true);
  });

  test("numeric values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(1, 2, 3)"], // All non-zero (truthy) - returns TRUE
        ["A2", "=OR(1, 0, 3)"], // Contains non-zero - returns TRUE
        ["A3", "=OR(0, 0)"], // All zeros (falsy) - returns FALSE
        ["A4", "=OR(-1, -5)"], // Negative numbers (truthy)
        ["A5", "=OR(0.1, 0.5)"], // Decimals (truthy)
        ["A6", "=OR(0, 0, 1)"], // One truthy value at end
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND - any truthy returns TRUE
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(true);
    expect(cell("A5")).toBe(true);
    expect(cell("A6")).toBe(true);
  });

  test("string values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(\"hello\", \"world\")"], // Non-empty strings (truthy)
        ["A2", "=OR(\"hello\", \"\")"], // One non-empty string - returns TRUE
        ["A3", "=OR(\"\", \"\")"], // All empty strings (falsy)
        ["A4", "=OR(\"0\", \"false\")"], // String representations (truthy)
        ["A5", "=OR(\"\", \"test\")"], // Empty string with non-empty
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND - any truthy returns TRUE
    expect(cell("A3")).toBe(false);
    expect(cell("A4")).toBe(true); // "0" and "false" as strings are truthy
    expect(cell("A5")).toBe(true);
  });

  test("mixed data types", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(TRUE, 1, \"hello\")"], // All truthy
        ["A2", "=OR(TRUE, 0, \"hello\")"], // Contains truthy values
        ["A3", "=OR(TRUE, 1, \"\")"], // Contains truthy values
        ["A4", "=OR(FALSE, 1, \"hello\")"], // Contains truthy values
        ["A5", "=OR(FALSE, 0, \"\")"], // All falsy - returns FALSE
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND - any truthy returns TRUE
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(true);
    expect(cell("A5")).toBe(false); // All falsy returns FALSE
  });

  test("cell references", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", true],
        ["B2", false],
        ["B3", 1],
        ["B4", 0],
        ["B5", "text"],
        ["B6", ""],
        ["A1", "=OR(B1, B3, B5)"], // TRUE, 1, "text" - all truthy
        ["A2", "=OR(B1, B2, B3)"], // TRUE, FALSE, 1 - contains truthy
        ["A3", "=OR(B1, B4, B5)"], // TRUE, 0, "text" - contains truthy
        ["A4", "=OR(B1, B3, B6)"], // TRUE, 1, "" - contains truthy
        ["A5", "=OR(B2, B4, B6)"], // FALSE, 0, "" - all falsy
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND
    expect(cell("A3")).toBe(true); // Different from AND
    expect(cell("A4")).toBe(true); // Different from AND
    expect(cell("A5")).toBe(false); // All falsy
  });

  test("range arguments", () => {
    // Test simple range with all truthy
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", 1],
        ["B2", 1],
        ["B3", 1],
        ["A1", "=OR(B1:B3)"], // All truthy numbers
      ])
    );

    expect(cell("A1")).toBe(true);
    
    // Test range with one truthy
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["C1", 0],
        ["C2", 1],
        ["C3", 0],
        ["A2", "=OR(C1:C3)"], // Contains one truthy - returns TRUE
      ])
    );

    expect(cell("A2")).toBe(true); // Different from AND
    
    // Test range with all falsy
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["D1", 0],
        ["D2", 0],
        ["D3", 0],
        ["A3", "=OR(D1:D3)"], // All zeros
      ])
    );

    expect(cell("A3")).toBe(false);
    
    // Test range with string
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["E1", ""],
        ["E2", "text"],
        ["E3", ""],
        ["A4", "=OR(E1:E3)"], // Contains non-empty string
      ])
    );

    expect(cell("A4")).toBe(true);
  });

  test("logical expressions", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", 10],
        ["B2", 5],
        ["B3", 15],
        ["A1", "=OR(B1>5, B2>0)"], // Both conditions true
        ["A2", "=OR(B1>5, B2>10)"], // One condition true
        ["A3", "=OR(B1<5, B2>0)"], // One condition true
        ["A4", "=OR(B1>B2, B3>B1)"], // Comparing cells
        ["A5", "=OR(B1<5, B2>10)"], // Both conditions false
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND
    expect(cell("A3")).toBe(true); // Different from AND
    expect(cell("A4")).toBe(true); // 10>5 OR 15>10
    expect(cell("A5")).toBe(false); // Both false
  });

  test("error handling", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR()"], // No arguments
        ["B1", "=1/0"], // Division by zero produces positive infinity
        ["B2", "=-1/0"], // Division by zero produces negative infinity
        ["A2", "=OR(B1, TRUE)"], // Positive infinity is truthy
        ["A3", "=OR(B2, TRUE)"], // Negative infinity is truthy
        ["A4", "=OR(B1, FALSE)"], // Infinity with FALSE - infinity is truthy
      ])
    );

    expect(cell("A1")).toBe(FormulaError.VALUE);
    // Division by zero produces infinity, which is truthy
    expect(cell("A2")).toBe(true); // Positive infinity OR TRUE = TRUE
    expect(cell("A3")).toBe(true); // Negative infinity OR TRUE = TRUE
    expect(cell("A4")).toBe(true); // Different from AND - infinity is truthy, so TRUE
  });

  test("short-circuit evaluation", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", "=1/0"], // Division by zero produces positive infinity
        ["A1", "=OR(TRUE, B1)"], // TRUE OR infinity - should short-circuit
        ["A2", "=OR(1, B1)"], // 1 OR infinity - should short-circuit
        ["A3", "=OR(FALSE, TRUE, B1)"], // FALSE OR TRUE OR infinity
      ])
    );

    // With short-circuit evaluation, these should all return TRUE without evaluating B1
    expect(cell("A1")).toBe(true); // TRUE OR anything = TRUE
    expect(cell("A2")).toBe(true); // 1 OR anything = TRUE
    expect(cell("A3")).toBe(true); // FALSE OR TRUE = TRUE
  });

  test("infinity values", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(1/0, TRUE)"], // Positive infinity is truthy
        ["A2", "=OR(-1/0, TRUE)"], // Negative infinity is truthy
        ["A3", "=OR(1/0, FALSE)"], // Infinity with false - infinity is truthy
        ["A4", "=OR(1/0, -1/0)"], // Both infinities are truthy
        ["A5", "=OR(FALSE, 1/0)"], // FALSE OR infinity
      ])
    );

    expect(cell("A1")).toBe(true); // Positive infinity OR TRUE = TRUE
    expect(cell("A2")).toBe(true); // Negative infinity OR TRUE = TRUE
    expect(cell("A3")).toBe(true); // Different from AND - infinity is truthy
    expect(cell("A4")).toBe(true); // Both infinities are truthy
    expect(cell("A5")).toBe(true);
  });

  test("nested OR functions", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(OR(TRUE, TRUE), OR(TRUE, TRUE))"], // Nested ORs, all true
        ["A2", "=OR(OR(FALSE, FALSE), OR(TRUE, TRUE))"], // Nested ORs, one true
        ["A3", "=OR(TRUE, OR(TRUE, TRUE))"], // Mixed nested
        ["A4", "=OR(OR(FALSE, FALSE), OR(FALSE, FALSE))"], // All false
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(false); // All false
  });

  test("OR with AND combinations", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(AND(TRUE, TRUE), FALSE)"], // (TRUE AND TRUE) OR FALSE = TRUE
        ["A2", "=OR(AND(TRUE, FALSE), FALSE)"], // (TRUE AND FALSE) OR FALSE = FALSE
        ["A3", "=OR(AND(TRUE, FALSE), TRUE)"], // (TRUE AND FALSE) OR TRUE = TRUE
        ["A4", "=AND(OR(TRUE, FALSE), TRUE)"], // (TRUE OR FALSE) AND TRUE = TRUE
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(true);
  });

  test("dynamic arrays with SEQUENCE", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=SEQUENCE(3, 1, 1, 1)"], // {1; 2; 3}
        ["B1", "=SEQUENCE(3, 1, 0, 0)"], // {0; 0; 0}
        ["C1", "=OR(A1:A3)"], // OR of {1, 2, 3} - all truthy
        ["C2", "=OR(B1:B3)"], // OR of {0, 0, 0} - all falsy
      ])
    );

    expect(cell("C1")).toBe(true);
    expect(cell("C2")).toBe(false); // All zeros
  });

  test("large number of arguments", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(1,1,1,1,1,1,1,1,1,1)"], // 10 truthy arguments
        ["A2", "=OR(0,0,0,0,1,0,0,0,0,0)"], // 10 arguments with one truthy
        ["A3", "=OR(0,0,0,0,0,0,0,0,0,0)"], // 10 falsy arguments
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(true); // Different from AND - one truthy returns TRUE
    expect(cell("A3")).toBe(false); // All falsy
  });

  test("empty cells in ranges", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["B1", 0],
        ["B2", ""], // Empty cell
        ["B3", 0],
        ["A1", "=OR(B1:B3)"], // Range with all falsy (empty string is skipped)
      ])
    );

    // Note: evaluateAllCells typically skips empty cells in ranges
    // This is consistent with Excel's behavior where empty cells are ignored
    expect(cell("A1")).toBe(false); // All zeros, empty is skipped
    
    // Test with one truthy value
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["C1", 0],
        ["C3", 1],
        // C2 is undefined (not set)
        ["A2", "=OR(C1:C3)"], // Range with one truthy
      ])
    );
    
    expect(cell("A2")).toBe(true);
  });

  test("boolean edge cases", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        ["A1", "=OR(TRUE)"], // Single TRUE
        ["A2", "=OR(FALSE)"], // Single FALSE
        ["A3", "=OR(1=1)"], // Expression that evaluates to TRUE
        ["A4", "=OR(1=2)"], // Expression that evaluates to FALSE
        ["A5", "=OR(1=1, 2=2, 3=3)"], // Multiple true expressions
        ["A6", "=OR(1=2, 2=3, 3=4)"], // Multiple false expressions
        ["A7", "=OR(1=2, 2=2, 3=4)"], // One true expression
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(false);
    expect(cell("A5")).toBe(true);
    expect(cell("A6")).toBe(false); // All false
    expect(cell("A7")).toBe(true); // One true
  });

  test("comparison with AND behavior", () => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>([
        // Same inputs, different results
        ["B1", true],
        ["B2", false],
        ["A1", "=OR(B1, B2)"], // TRUE OR FALSE = TRUE
        ["A2", "=AND(B1, B2)"], // TRUE AND FALSE = FALSE
        
        ["C1", 1],
        ["C2", 0],
        ["A3", "=OR(C1, C2)"], // 1 OR 0 = TRUE
        ["A4", "=AND(C1, C2)"], // 1 AND 0 = FALSE
      ])
    );

    expect(cell("A1")).toBe(true);
    expect(cell("A2")).toBe(false);
    expect(cell("A3")).toBe(true);
    expect(cell("A4")).toBe(false);
  });
});
