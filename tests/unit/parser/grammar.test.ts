import { describe, expect, test } from "bun:test";
import {
  CELL_REFERENCE_PATTERNS,
  OPERATOR_ASSOCIATIVITY,
  OPERATOR_PRECEDENCE,
  REQUIRED_ARG_FUNCTIONS,
  SPECIAL_CONSTANTS,
  SPECIAL_FUNCTIONS,
  VARIADIC_FUNCTIONS,
  compareOperatorPrecedence,
  getOperatorAssociativity,
  getOperatorPrecedence,
  isBinaryOperator,
  isReservedKeyword,
  isValidColumn,
  isValidRow,
  parse3DReference,
  parseCellReference,
  parseStructuredReference,
} from "../../../src/parser/grammar";

describe("Operator Precedence", () => {
  test("should define correct precedence levels", () => {
    // Comparison operators (lowest)
    expect(OPERATOR_PRECEDENCE["="]).toBe(1);
    expect(OPERATOR_PRECEDENCE["<>"]).toBe(1);
    expect(OPERATOR_PRECEDENCE["<"]).toBe(1);
    expect(OPERATOR_PRECEDENCE[">"]).toBe(1);
    expect(OPERATOR_PRECEDENCE["<="]).toBe(1);
    expect(OPERATOR_PRECEDENCE[">="]).toBe(1);

    // Concatenation
    expect(OPERATOR_PRECEDENCE["&"]).toBe(2);

    // Addition/Subtraction
    expect(OPERATOR_PRECEDENCE["+"]).toBe(3);
    expect(OPERATOR_PRECEDENCE["-"]).toBe(3);

    // Multiplication/Division
    expect(OPERATOR_PRECEDENCE["*"]).toBe(4);
    expect(OPERATOR_PRECEDENCE["/"]).toBe(4);

    // Exponentiation (highest)
    expect(OPERATOR_PRECEDENCE["^"]).toBe(5);
  });

  test("should identify binary operators correctly", () => {
    expect(isBinaryOperator("+")).toBe(true);
    expect(isBinaryOperator("-")).toBe(true);
    expect(isBinaryOperator("*")).toBe(true);
    expect(isBinaryOperator("/")).toBe(true);
    expect(isBinaryOperator("^")).toBe(true);
    expect(isBinaryOperator("&")).toBe(true);
    expect(isBinaryOperator("=")).toBe(true);
    expect(isBinaryOperator("<>")).toBe(true);

    expect(isBinaryOperator("%")).toBe(false);
    expect(isBinaryOperator("!")).toBe(false);
    expect(isBinaryOperator("~")).toBe(false);
  });

  test("should get operator precedence correctly", () => {
    expect(getOperatorPrecedence("+")).toBe(3);
    expect(getOperatorPrecedence("*")).toBe(4);
    expect(getOperatorPrecedence("^")).toBe(5);
    expect(getOperatorPrecedence("unknown")).toBe(0);
  });

  test("should compare operator precedence correctly", () => {
    expect(compareOperatorPrecedence("*", "+")).toBeGreaterThan(0);
    expect(compareOperatorPrecedence("+", "*")).toBeLessThan(0);
    expect(compareOperatorPrecedence("+", "-")).toBe(0);
    expect(compareOperatorPrecedence("^", "*")).toBeGreaterThan(0);
  });
});

describe("Operator Associativity", () => {
  test("should define correct associativity", () => {
    // Most operators are left-associative
    expect(OPERATOR_ASSOCIATIVITY["+"]).toBe("left");
    expect(OPERATOR_ASSOCIATIVITY["-"]).toBe("left");
    expect(OPERATOR_ASSOCIATIVITY["*"]).toBe("left");
    expect(OPERATOR_ASSOCIATIVITY["/"]).toBe("left");
    expect(OPERATOR_ASSOCIATIVITY["&"]).toBe("left");

    // Exponentiation is right-associative
    expect(OPERATOR_ASSOCIATIVITY["^"]).toBe("right");
  });

  test("should get operator associativity correctly", () => {
    expect(getOperatorAssociativity("+")).toBe("left");
    expect(getOperatorAssociativity("^")).toBe("right");
    expect(getOperatorAssociativity("unknown")).toBe("left");
  });
});

describe("Special Functions", () => {
  test("should identify special functions that dont require parentheses", () => {
    expect(SPECIAL_FUNCTIONS.has("PI")).toBe(true);
    expect(SPECIAL_FUNCTIONS.has("TRUE")).toBe(true);
    expect(SPECIAL_FUNCTIONS.has("FALSE")).toBe(true);
    expect(SPECIAL_FUNCTIONS.has("NA")).toBe(true);

    expect(SPECIAL_FUNCTIONS.has("SUM")).toBe(false);
    expect(SPECIAL_FUNCTIONS.has("AVERAGE")).toBe(false);
  });
});

describe("Function Categories", () => {
  test("should identify variadic functions", () => {
    expect(VARIADIC_FUNCTIONS.has("SUM")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("PRODUCT")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("COUNT")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("MAX")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("MIN")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("AND")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("OR")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("CONCATENATE")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("CHOOSE")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("IFS")).toBe(true);
    expect(VARIADIC_FUNCTIONS.has("SWITCH")).toBe(true);

    expect(VARIADIC_FUNCTIONS.has("ABS")).toBe(false);
    expect(VARIADIC_FUNCTIONS.has("SQRT")).toBe(false);
  });

  test("should identify functions requiring at least one argument", () => {
    expect(REQUIRED_ARG_FUNCTIONS.has("SUM")).toBe(true);
    expect(REQUIRED_ARG_FUNCTIONS.has("AVERAGE")).toBe(true);
    expect(REQUIRED_ARG_FUNCTIONS.has("CONCATENATE")).toBe(true);

    expect(REQUIRED_ARG_FUNCTIONS.has("PI")).toBe(false);
    expect(REQUIRED_ARG_FUNCTIONS.has("TODAY")).toBe(false);
  });
});

describe("Reserved Keywords", () => {
  test("should identify reserved keywords", () => {
    expect(isReservedKeyword("TRUE")).toBe(true);
    expect(isReservedKeyword("FALSE")).toBe(true);
    expect(isReservedKeyword("NULL")).toBe(true);
    expect(isReservedKeyword("AND")).toBe(true);
    expect(isReservedKeyword("OR")).toBe(true);
    expect(isReservedKeyword("NOT")).toBe(true);

    // Case insensitive
    expect(isReservedKeyword("true")).toBe(true);
    expect(isReservedKeyword("False")).toBe(true);

    // Not reserved
    expect(isReservedKeyword("SUM")).toBe(false);
    expect(isReservedKeyword("MyVariable")).toBe(false);
  });
});

describe("Cell Reference Patterns", () => {
  test("should validate column references", () => {
    expect(isValidColumn("A")).toBe(true);
    expect(isValidColumn("Z")).toBe(true);
    expect(isValidColumn("AA")).toBe(true);
    expect(isValidColumn("XFD")).toBe(true);

    expect(isValidColumn("1")).toBe(false);
    expect(isValidColumn("A1")).toBe(false);
    expect(isValidColumn("")).toBe(false);
    expect(isValidColumn("$A")).toBe(false);
  });

  test("should validate row references", () => {
    expect(isValidRow("1")).toBe(true);
    expect(isValidRow("100")).toBe(true);
    expect(isValidRow("1048576")).toBe(true);

    expect(isValidRow("0")).toBe(false);
    expect(isValidRow("A")).toBe(false);
    expect(isValidRow("1A")).toBe(false);
    expect(isValidRow("")).toBe(false);
    expect(isValidRow("$1")).toBe(false);
  });

  test("should parse cell references correctly", () => {
    const a1 = parseCellReference("A1");
    expect(a1).toEqual({
      sheet: undefined,
      colAbsolute: false,
      col: "A",
      rowAbsolute: false,
      row: "1",
    });

    const absolute = parseCellReference("$B$2");
    expect(absolute).toEqual({
      sheet: undefined,
      colAbsolute: true,
      col: "B",
      rowAbsolute: true,
      row: "2",
    });

    const mixed1 = parseCellReference("$C3");
    expect(mixed1).toEqual({
      sheet: undefined,
      colAbsolute: true,
      col: "C",
      rowAbsolute: false,
      row: "3",
    });

    const mixed2 = parseCellReference("D$4");
    expect(mixed2).toEqual({
      sheet: undefined,
      colAbsolute: false,
      col: "D",
      rowAbsolute: true,
      row: "4",
    });
  });

  test("should parse sheet-qualified references", () => {
    const simple = parseCellReference("Sheet1!A1");
    expect(simple).toEqual({
      sheet: "Sheet1",
      colAbsolute: false,
      col: "A",
      rowAbsolute: false,
      row: "1",
    });

    const quoted = parseCellReference("'My Sheet'!B2");
    expect(quoted).toEqual({
      sheet: "My Sheet",
      colAbsolute: false,
      col: "B",
      rowAbsolute: false,
      row: "2",
    });

    const complex = parseCellReference("'Sheet-123'!$C$3");
    expect(complex).toEqual({
      sheet: "Sheet-123",
      colAbsolute: true,
      col: "C",
      rowAbsolute: true,
      row: "3",
    });
  });

  test("should return null for invalid references", () => {
    expect(parseCellReference("")).toBeNull();
    expect(parseCellReference("123")).toBeNull();
    expect(parseCellReference("ABC")).toBeNull();
    expect(parseCellReference("A0")).toBeNull();
    expect(parseCellReference("1A")).toBeNull();
    expect(parseCellReference("$$A1")).toBeNull();
  });
});

describe("Special Constants", () => {
  test("should recognize INFINITY as a special constant", () => {
    expect(SPECIAL_CONSTANTS.has("INFINITY")).toBe(true);
    expect(SPECIAL_CONSTANTS.size).toBe(1);
  });
});

describe("3D Range Parsing", () => {
  test("should parse basic 3D range references", () => {
    const result = parse3DReference("Sheet1:Sheet3!A1");
    expect(result).toEqual({
      startSheet: "Sheet1",
      endSheet: "Sheet3",
      reference: "A1",
    });
  });

  test("should parse 3D range with quoted sheet names", () => {
    const result = parse3DReference("'Sheet 1':'Sheet 3'!B2");
    expect(result).toEqual({
      startSheet: "Sheet 1",
      endSheet: "Sheet 3",
      reference: "B2",
    });
  });

  test("should parse 3D range with range reference", () => {
    const result = parse3DReference("Sheet1:Sheet5!A1:B10");
    expect(result).toEqual({
      startSheet: "Sheet1",
      endSheet: "Sheet5",
      reference: "A1:B10",
    });
  });

  test("should return null for invalid 3D ranges", () => {
    expect(parse3DReference("Sheet1!A1")).toBeNull(); // Not a 3D range
    expect(parse3DReference("Sheet1:Sheet2")).toBeNull(); // Missing reference
    expect(parse3DReference("Sheet1:")).toBeNull(); // Incomplete
  });
});

describe("Structured Reference Parsing", () => {
  test("should parse simple table column reference", () => {
    const result = parseStructuredReference("Table1[Sales]");
    expect(result).toEqual({
      tableName: "Table1",
      columnSpec: "Sales",
      selector: undefined,
      isCurrentRow: undefined,
      cols: {
        startCol: "Sales",
        endCol: "Sales",
      },
    });
  });

  test("should parse current row reference without table", () => {
    const result = parseStructuredReference("@Sales");
    expect(result).toEqual({
      tableName: "",
      columnSpec: "Sales",
      isCurrentRow: true,
      selector: undefined,
      cols: {
        startCol: "Sales",
        endCol: "Sales",
      },
    });
  });

  test("should parse table with current row reference", () => {
    const result = parseStructuredReference("Table1[@Sales]");
    expect(result).toEqual({
      tableName: "Table1",
      columnSpec: "Sales",
      isCurrentRow: true,
      selector: undefined,
      cols: {
        startCol: "Sales",
        endCol: "Sales",
      },
    });
  });

  test("should parse table with selector and column", () => {
    const result = parseStructuredReference("Table1[[#Headers],[Sales]]");
    expect(result).toEqual({
      tableName: "Table1",
      columnSpec: "[#Headers],[Sales]",
      selector: "#Headers",
      isCurrentRow: undefined,
      cols: {
        startCol: "Sales",
        endCol: "Sales",
      },
    });
  });

  test("should parse column range references", () => {
    const result = parseStructuredReference("Table1[Sales:Quantity]");
    expect(result).toEqual({
      tableName: "Table1",
      columnSpec: "Sales:Quantity",
      selector: undefined,
      isCurrentRow: undefined,
      cols: {
        startCol: "Sales",
        endCol: "Quantity",
      },
    });
  });

  test("should parse selector with column range", () => {
    const result = parseStructuredReference("Table1[[#Data],[Sales:Quantity]]");
    expect(result).toEqual({
      tableName: "Table1",
      columnSpec: "[#Data],[Sales:Quantity]",
      selector: "#Data",
      isCurrentRow: undefined,
      cols: {
        startCol: "Sales",
        endCol: "Quantity",
      },
    });
  });

  test("should return null for invalid structured references", () => {
    expect(parseStructuredReference("NotATable")).toBeNull();
    expect(parseStructuredReference("Table1")).toBeNull(); // Missing brackets
    expect(parseStructuredReference("@")).toBeNull(); // @ without column name
  });
});

describe("Cell Reference Pattern Updates", () => {
  test("should have pattern for 3D ranges", () => {
    const pattern = CELL_REFERENCE_PATTERNS.SHEET_RANGE_QUALIFIED;

    // Test unquoted sheets
    expect("Sheet1:Sheet3!A1".match(pattern)).toBeTruthy();
    expect("Start:End!B2:C3".match(pattern)).toBeTruthy();

    // Test quoted sheets
    expect("'Sheet 1':'Sheet 3'!A1".match(pattern)).toBeTruthy();
    expect("'Start':'End'!B2".match(pattern)).toBeTruthy();

    // Test mixed
    expect("Sheet1:'End Sheet'!A1".match(pattern)).toBeTruthy();
  });

  test("should have pattern for table references", () => {
    const pattern = CELL_REFERENCE_PATTERNS.TABLE_REFERENCE;

    expect("Table1[Column1]".match(pattern)).toBeTruthy();
    expect("MyTable[Sales]".match(pattern)).toBeTruthy();
    expect("Table_123[My Column]".match(pattern)).toBeTruthy();
  });

  test("should have pattern for current row references", () => {
    const pattern = CELL_REFERENCE_PATTERNS.CURRENT_ROW_REFERENCE;

    expect("@Column1".match(pattern)).toBeTruthy();
    expect("@Sales".match(pattern)).toBeTruthy();
    expect("@My_Column".match(pattern)).toBeTruthy();
  });

  test("should have pattern for table selectors", () => {
    const pattern = CELL_REFERENCE_PATTERNS.TABLE_SELECTOR;

    expect("#All".match(pattern)).toBeTruthy();
    expect("#Data".match(pattern)).toBeTruthy();
    expect("#Headers".match(pattern)).toBeTruthy();

    // Case insensitive
    expect("#all".match(pattern)).toBeTruthy();
    expect("#DATA".match(pattern)).toBeTruthy();

    // Invalid selectors
    expect("#ThisRow".match(pattern)).toBeFalsy(); // Removed #ThisRow
    expect("#Invalid".match(pattern)).toBeFalsy();
    expect("#".match(pattern)).toBeFalsy();
  });
});
