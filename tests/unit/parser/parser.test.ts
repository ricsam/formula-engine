// @ts-nocheck
import { describe, expect, test } from "bun:test";
import { FormulaError } from "../../../src/core/types";
import type { ArrayNode } from "../../../src/parser/ast";
import { ParseError, parseFormula } from "../../../src/parser/parser";
import { parse } from "path";

describe("Parser - Basic Values", () => {
  test("should parse numbers", () => {
    const ast = parseFormula("42");
    expect(ast.type).toBe("value");
    expect(ast.value).toEqual({ type: "number", value: 42 });

    const decimal = parseFormula("3.14159");
    expect(decimal.value).toEqual({ type: "number", value: 3.14159 });

    const scientific = parseFormula("1.23E-4");
    expect(scientific.value).toEqual({ type: "number", value: 0.000123 });
  });

  test("should parse strings", () => {
    const ast = parseFormula('"hello world"');
    expect(ast.type).toBe("value");
    expect(ast.value).toEqual({ type: "string", value: "hello world" });

    const escaped = parseFormula('"Say ""Hello"""');
    expect(escaped.value).toEqual({ type: "string", value: 'Say "Hello"' });
  });

  test("should parse booleans", () => {
    const trueAst = parseFormula("TRUE");
    expect(trueAst.value).toEqual({ type: "boolean", value: true });

    const falseAst = parseFormula("FALSE");
    expect(falseAst.value).toEqual({ type: "boolean", value: false });
  });

  test("should parse error values", () => {
    const ast = parseFormula("#DIV/0!");
    expect(ast.type).toBe("error");
    expect(ast.error).toBe(FormulaError.DIV0);
  });

  test("should parse empty formula", () => {
    const ast = parseFormula("");
    expect(ast.type).toBe("empty");
  });
});

describe("Parser - Cell References", () => {
  test("should parse simple cell references", () => {
    const ast = parseFormula("A1");
    expect(ast.type).toBe("reference");
    expect(ast.address).toEqual({ colIndex: 0, rowIndex: 0 });
    expect(ast.isAbsolute).toEqual({ col: false, row: false });
  });

  test("should parse absolute cell references", () => {
    const ast = parseFormula("$B$2");
    expect(ast.type).toBe("reference");
    expect(ast.address).toEqual({ colIndex: 1, rowIndex: 1 });
    expect(ast.isAbsolute).toEqual({ col: true, row: true });
  });

  test("should parse mixed cell references", () => {
    const mixed1 = parseFormula("$C3");
    expect(mixed1.address).toEqual({ colIndex: 2, rowIndex: 2 });
    expect(mixed1.isAbsolute).toEqual({ col: true, row: false });

    const mixed2 = parseFormula("D$4");

    expect(mixed2.address).toEqual({ colIndex: 3, rowIndex: 3 });
    expect(mixed2.isAbsolute).toEqual({ col: false, row: true });
  });

  test("should parse range references", () => {
    const ast = parseFormula("A1:B2");
    expect(ast.type).toBe("range");
    expect(ast.range).toEqual({
      start: { col: 0, row: 0 },
      end: {
        col: {
          type: "number",
          value: 1,
        },
        row: {
          type: "number",
          value: 1,
        },
      },
    });
  });

  test("should parse absolute range references", () => {
    const ast = parseFormula("$A$1:$B$2");
    expect(ast.type).toBe("range");
    expect(ast.isAbsolute).toEqual({
      start: { col: true, row: true },
      end: { col: true, row: true },
    });
  });

  test("should normalize range references", () => {
    const ast = parseFormula("B2:A1");
    expect(ast.range).toEqual({
      start: { col: 0, row: 0 },
      end: {
        col: {
          type: "number",
          value: 1,
        },
        row: { type: "number", value: 1 },
      },
    });
  });
});

describe("Parser - Operators", () => {
  test("should parse binary operators", () => {
    const add = parseFormula("1+2");
    expect(add.type).toBe("binary-op");
    expect(add.operator).toBe("+");
    expect(add.left.type).toBe("value");
    expect(add.left.value).toEqual({ type: "number", value: 1 });
    expect(add.right.type).toBe("value");
    expect(add.right.value).toEqual({
      type: "number",
      value: 2,
    });

    const mul = parseFormula("3*4");
    expect(mul.operator).toBe("*");

    const concat = parseFormula('"A"&"B"');
    expect(concat.operator).toBe("&");
  });

  test("should parse comparison operators", () => {
    const eq = parseFormula("A1=10");
    expect(eq.operator).toBe("=");

    const ne = parseFormula("A1<>10");
    expect(ne.operator).toBe("<>");

    const gte = parseFormula("A1>=10");
    expect(gte.operator).toBe(">=");
  });

  test("should parse unary operators", () => {
    const neg = parseFormula("-5");

    expect(neg.type).toBe("unary-op");
    expect(neg.operator).toBe("-");
    expect(neg.operand.type).toBe("value");

    expect(neg.operand.value).toEqual({
      type: "number",
      value: 5,
    });

    const pos = parseFormula("+5");

    expect(pos.operator).toBe("+");
  });

  test("should parse percentage operator", () => {
    const pct = parseFormula("50%");

    expect(pct.type).toBe("unary-op");
    expect(pct.operator).toBe("%");
    expect(pct.operand.type).toBe("value");

    expect(pct.operand.value).toEqual({
      type: "number",
      value: 50,
    });
  });

  test("should respect operator precedence", () => {
    const ast = parseFormula("1+2*3");

    expect(ast.operator).toBe("+");
    expect(ast.left.type).toBe("value");

    expect(ast.left.value).toEqual({ type: "number", value: 1 });
    expect(ast.right.type).toBe("binary-op");

    expect(ast.right.operator).toBe("*");

    const ast2 = parseFormula("2^3^4");

    expect(ast2.operator).toBe("^");
    expect(ast2.left.type).toBe("value");

    expect(ast2.left.value).toEqual({
      type: "number",
      value: 2,
    });
    // Right associative - should parse as 2^(3^4)
    expect(ast2.right.type).toBe("binary-op");

    expect(ast2.right.operator).toBe("^");
  });

  test("should handle parentheses", () => {
    const ast = parseFormula("(1+2)*3");

    expect(ast.operator).toBe("*");
    expect(ast.left.type).toBe("binary-op");

    expect(ast.left.operator).toBe("+");
    expect(ast.right.type).toBe("value");
    expect(ast.right.value).toEqual({
      type: "number",
      value: 3,
    });
  });
});

describe("Parser - Functions", () => {
  test("should parse function calls with no arguments", () => {
    const ast = parseFormula("PI()");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("PI");
    expect(ast.args).toEqual([]);
  });

  test("should parse function calls with arguments", () => {
    const ast = parseFormula("SUM(1,2,3)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
    const [arg1, arg2, arg3] = ast.args;
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    expect(arg2.type).toBe("value");
    expect(arg2.value).toEqual({
      type: "number",
      value: 2,
    });
    expect(arg3.type).toBe("value");
    expect(arg3.value).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("should parse nested function calls", () => {
    const ast = parseFormula("SUM(1,ABS(-2),3)");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
    const absNode = ast.args[1];
    expect(absNode.type).toBe("function");
    expect(absNode.name).toBe("ABS");
  });

  test("should parse functions with range arguments", () => {
    const ast = parseFormula("SUM(A1:A10)");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(1);
    const rangeNode = ast.args[0];
    expect(rangeNode.type).toBe("range");
  });

  test("should normalize function names to uppercase", () => {
    const ast = parseFormula("sum(1,2)");
    expect(ast.name).toBe("SUM");
  });
});

describe("Parser - Arrays", () => {
  test("should parse array literals", () => {
    const ast = parseFormula("{1,2,3}") as ArrayNode;
    expect(ast.type).toBe("array");
    expect(ast.elements).toHaveLength(1);
    expect(ast.elements[0]).toHaveLength(3);
    const firstElement = ast.elements[0];
    const [arg1, arg2, arg3] = firstElement;
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    expect(arg2.type).toBe("value");
    expect(arg2.value).toEqual({
      type: "number",
      value: 2,
    });
    expect(arg3.type).toBe("value");
    expect(arg3.value).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("should parse 2D array literals", () => {
    const ast = parseFormula("{1,2;3,4}");
    expect(ast.type).toBe("array");
    expect(ast.elements).toHaveLength(2);
    expect(ast.elements[0]).toHaveLength(2);
    expect(ast.elements[1]).toHaveLength(2);
    const firstElement = ast.elements[0];
    const [arg1, arg2] = firstElement;
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    const secondElement = ast.elements[1];
    const last = secondElement[1];
    expect(last.type).toBe("value");
    expect(last.value).toEqual({
      type: "number",
      value: 4,
    });
  });

  test("should parse arrays with mixed types", () => {
    const ast = parseFormula('{1,"text",TRUE}') as ArrayNode;
    expect(ast.elements[0]).toHaveLength(3);
    const firstElement = ast.elements[0];
    const [arg1, arg2, arg3] = firstElement;
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    expect(arg2.type).toBe("value");
    expect(arg2.value).toEqual({
      type: "string",
      value: "text",
    });
    expect(arg3.type).toBe("value");
    expect(arg3.value).toEqual({
      type: "boolean",
      value: true,
    });
  });

  test("should enforce consistent row lengths", () => {
    expect(() => parseFormula("{1,2;3}")).toThrow(ParseError);
    expect(() => parseFormula("{1;2,3}")).toThrow(ParseError);
  });

  test("should handle empty arrays", () => {
    const ast = parseFormula("{}");
    expect(ast.type).toBe("array");
    expect(ast.elements).toHaveLength(1);
    expect(ast.elements[0]).toHaveLength(1);
    const emptyNode = ast.elements[0]?.[0];
    expect(emptyNode.type).toBe("empty");
  });
});

describe("Parser - Named Expressions", () => {
  test("should parse named expressions", () => {
    const ast = parseFormula("TaxRate");
    expect(ast.type).toBe("named-expression");
    expect(ast.name).toBe("TaxRate");
    expect(ast.sheetName).toBeUndefined();
  });

  test("should distinguish between functions and named expressions", () => {
    const func = parseFormula("SUM(1,2)");
    expect(func.type).toBe("function");

    const name = parseFormula("SUM");
    expect(name.type).toBe("named-expression");
  });

  test("should parse sheet-qualified named expressions", () => {
    const ast = parseFormula("Sheet1!TaxRate");
    expect(ast.type).toBe("named-expression");
    expect(ast.name).toBe("TaxRate");
    expect(ast.sheetName).toBe("Sheet1");
  });

  test("should parse sheet-qualified named expressions with quoted sheet names", () => {
    const ast = parseFormula("'My Sheet'!TaxRate");
    expect(ast.type).toBe("named-expression");
    expect(ast.name).toBe("TaxRate");
    expect(ast.sheetName).toBe("My Sheet");
  });
});

describe("Parser - Complex Expressions", () => {
  test("should parse complex arithmetic expressions", () => {
    const formula = "(A1+B1)*2-C1/3+D1^2";
    const ast = parseFormula(formula);
    expect(ast.type).toBe("binary-op");
  });

  test("should parse nested expressions with functions", () => {
    const formula = "IF(SUM(A1:A10)>100,MAX(B1:B10),MIN(C1:C10))";
    const ast = parseFormula(formula);
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("IF");
    expect(ast.args).toHaveLength(3);

    const condition = ast.args[0];
    expect(condition.type).toBe("binary-op");
    expect(condition.operator).toBe(">");
  });

  test("should parse array operations", () => {
    const formula = "SUM(A1:A10*B1:B10)";
    const ast = parseFormula(formula);
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(1);

    const mul = ast.args[0];
    expect(mul.type).toBe("binary-op");
    expect(mul.operator).toBe("*");
  });
});

describe("Parser - INFINITY Literal", () => {
  test("should parse INFINITY as a literal", () => {
    const ast = parseFormula("INFINITY");
    expect(ast.type).toBe("infinity");
  });

  test("should parse INFINITY in expressions", () => {
    const ast = parseFormula("SEQUENCE(INFINITY)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SEQUENCE");
    expect(ast.args[0]?.type).toBe("infinity");
  });

  test("should parse arithmetic with INFINITY", () => {
    const ast = parseFormula("1/INFINITY");
    expect(ast.type).toBe("binary-op");
    expect(ast.operator).toBe("/");
    expect(ast.right.type).toBe("infinity");
  });

  test("should parse arithmetic with -INFINITY", () => {
    const ast = parseFormula("-INFINITY");
    expect(ast.type).toBe("unary-op");
    expect(ast.operator).toBe("-");
    expect(ast.operand.type).toBe("infinity");
  });
});

describe("Parser - 3D Ranges", () => {
  test("should parse basic 3D range", () => {
    const ast = parseFormula("Sheet1:Sheet3!A1");
    expect(ast.type).toBe("3d-range");
    expect(ast.startSheet).toBe("Sheet1");
    expect(ast.endSheet).toBe("Sheet3");
    expect(ast.reference.type).toBe("reference");
  });

  test("should parse 3D range with quoted sheet names", () => {
    const ast = parseFormula("'Sheet 1':'Sheet 3'!A1");
    expect(ast.type).toBe("3d-range");
    expect(ast.startSheet).toBe("Sheet 1");
    expect(ast.endSheet).toBe("Sheet 3");
    expect(ast.reference.type).toBe("reference");
  });

  test("should parse 3D range with cell range", () => {
    const ast = parseFormula("Sheet1:Sheet5!A1:B10");
    expect(ast.type).toBe("3d-range");
    expect(ast.startSheet).toBe("Sheet1");
    expect(ast.endSheet).toBe("Sheet5");
    expect(ast.reference.type).toBe("range");
  });

  test("should parse 3D range in function", () => {
    const ast = parseFormula("AVERAGE(Sheet1:Sheet3!A1)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("AVERAGE");
    expect(ast.args[0]?.type).toBe("3d-range");
  });
});

describe("Parser - Structured References", () => {
  test("should parse simple table reference", () => {
    const ast = parseFormula("Table1[Sales]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({ startCol: "Sales", endCol: "Sales" });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse current row reference", () => {
    const ast = parseFormula("[@Sales]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({ startCol: "Sales", endCol: "Sales" });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse table with current row reference", () => {
    const ast = parseFormula("Table1[@Sales]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({ startCol: "Sales", endCol: "Sales" });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse table with selector", () => {
    const ast = parseFormula("Table1[[#Headers],[Sales]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({ startCol: "Sales", endCol: "Sales" });
    expect(ast.selector).toBe("#Headers");
  });

  test("should parse table reference in formula", () => {
    const ast = parseFormula("SUM(Table1[Sales])");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args[0]?.type).toBe("structured-reference");
  });

  test("should parse current row reference in arithmetic", () => {
    const ast = parseFormula("[@Price]*[@Quantity]");
    expect(ast.type).toBe("binary-op");

    expect(ast.operator).toBe("*");
    expect(ast.left.type).toBe("structured-reference");
    expect(ast.right.type).toBe("structured-reference");

    expect(ast.left.isCurrentRow).toBe(true);
    expect(ast.left.cols).toEqual({ startCol: "Price", endCol: "Price" });

    expect(ast.right.isCurrentRow).toBe(true);
    expect(ast.right.cols).toEqual({
      startCol: "Quantity",
      endCol: "Quantity",
    });
  });

  test("should parse table selector only", () => {
    const ast = parseFormula("Table1[#Data]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.selector).toBe("#Data");
    expect(ast.cols).toBeUndefined();
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse column range references", () => {
    const ast = parseFormula("Table1[Sales:Quantity]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({ startCol: "Sales", endCol: "Quantity" });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse selector with column range", () => {
    const ast = parseFormula("Table1[[#Data],[Sales:Quantity]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({ startCol: "Sales", endCol: "Quantity" });
    expect(ast.selector).toBe("#Data");
  });

  test("should parse bracketed column range syntax", () => {
    const ast = parseFormula("Table1[[num]:[result]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "num",
      endCol: "result",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse bracketed current row column range syntax", () => {
    const ast = parseFormula("Table1[@[num]:[result]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "num",
      endCol: "result",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse simple current row column range syntax", () => {
    const ast = parseFormula("Table1[@num:result]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "num",
      endCol: "result",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse column names with dashes", () => {
    const ast = parseFormula("Table1[ORDER-ID]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "ORDER-ID",
      endCol: "ORDER-ID",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse column names with equals signs", () => {
    const ast = parseFormula("Table1[Number of cars to prepare = number of ERTs required]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "Number of cars to prepare = number of ERTs required",
      endCol: "Number of cars to prepare = number of ERTs required",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse current row reference with equals signs in column names", () => {
    const ast = parseFormula("[@[Status = Active]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "Status = Active",
      endCol: "Status = Active",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse column names with decimal numbers", () => {
    const ast = parseFormula("Table1[13.3uM concentration]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "13.3uM concentration",
      endCol: "13.3uM concentration",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse column names with percentage signs", () => {
    const ast = parseFormula("[@[2.5% solution]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "2.5% solution",
      endCol: "2.5% solution",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse column names with Unicode symbols", () => {
    const testCases = [
      { formula: "[@[Final Stock Concentration (µM)]]", expected: "Final Stock Concentration (µM)" },
      { formula: "Table1[β-coefficient (°C)]", expected: "β-coefficient (°C)" },
      { formula: "[@[€ price]]", expected: "€ price" },
      { formula: "Table1[Δ temperature]", expected: "Δ temperature" },
      { formula: "[@[λ wavelength (nm)]]", expected: "λ wavelength (nm)" },
      { formula: "Table1[π value]", expected: "π value" },
      { formula: "[@[Ω resistance]]", expected: "Ω resistance" },
      { formula: "Table1[日本語 column]", expected: "日本語 column" }
    ];

    testCases.forEach(({ formula, expected }) => {
      const ast = parseFormula(formula);
      expect(ast.type).toBe("structured-reference");
      expect(ast.cols).toEqual({
        startCol: expected,
        endCol: expected,
      });
    });
  });

  test("should parse complex scientific column names", () => {
    const ast = parseFormula("[@[Total volume of 13.3uM detergent to prepare (uL)]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "Total volume of 13.3uM detergent to prepare (uL)",
      endCol: "Total volume of 13.3uM detergent to prepare (uL)",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse current row reference with column names containing dashes", () => {
    const ast = parseFormula("[@CUSTOMER-ID]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "CUSTOMER-ID",
      endCol: "CUSTOMER-ID",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse bracketed current row reference with dashed column names", () => {
    const ast = parseFormula("Table1[@[ORDER-ID]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "ORDER-ID",
      endCol: "ORDER-ID",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse column range with dashed column names", () => {
    const ast = parseFormula("Table1[ORDER-ID:CUSTOMER-ID]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "ORDER-ID",
      endCol: "CUSTOMER-ID",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  // Tests for bare column references (Excel table formulas)
  test("should parse bare column reference", () => {
    const ast = parseFormula("[result]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "result",
      endCol: "result",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse bare column range reference", () => {
    const ast = parseFormula("[num:result]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "num",
      endCol: "result",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse bare selector reference", () => {
    const ast = parseFormula("[#Data]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.selector).toBe("#Data");
    expect(ast.isCurrentRow).toBe(false);
  });

  // Tests for the user's Excel table formulas
  test("should parse Excel table formula: [@num] * 10", () => {
    const ast = parseFormula("[@num] * 10");
    expect(ast.type).toBe("binary-op");
    expect(ast.operator).toBe("*");
    expect(ast.left.type).toBe("structured-reference");
    expect(ast.left.tableName).toBeUndefined();
    expect(ast.left.isCurrentRow).toBe(true);
    expect(ast.left.cols?.startCol).toBe("num");
  });

  test("should parse Excel table formula: SUM([result])", () => {
    const ast = parseFormula("SUM([result])");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(1);
    expect(ast.args[0].type).toBe("structured-reference");
    expect(ast.args[0].tableName).toBeUndefined();
    expect(ast.args[0].cols?.startCol).toBe("result");
  });

  test("should parse INDEX+MATCH formula with dashed column names", () => {
    const ast = parseFormula(
      "INDEX(Table1[ORDER-ID], MATCH([@[CUSTOMER-ID]], Table1[CUSTOMER-ID],0))"
    );
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("INDEX");
    expect(ast.args).toHaveLength(2);

    // First argument: Table1[ORDER-ID]
    expect(ast.args[0].type).toBe("structured-reference");
    expect(ast.args[0].tableName).toBe("Table1");
    expect(ast.args[0].cols?.startCol).toBe("ORDER-ID");

    // Second argument: MATCH function
    expect(ast.args[1].type).toBe("function");
    expect(ast.args[1].name).toBe("MATCH");
    expect(ast.args[1].args).toHaveLength(3);

    // MATCH first argument: [@[CUSTOMER-ID]]
    expect(ast.args[1].args[0].type).toBe("structured-reference");
    expect(ast.args[1].args[0].tableName).toBeUndefined();
    expect(ast.args[1].args[0].isCurrentRow).toBe(true);
    expect(ast.args[1].args[0].cols?.startCol).toBe("CUSTOMER-ID");

    // MATCH second argument: Table1[CUSTOMER-ID]
    expect(ast.args[1].args[1].type).toBe("structured-reference");
    expect(ast.args[1].args[1].tableName).toBe("Table1");
    expect(ast.args[1].args[1].cols?.startCol).toBe("CUSTOMER-ID");

    // MATCH third argument: 0
    expect(ast.args[1].args[2].type).toBe("value");
    expect(ast.args[1].args[2].value.value).toBe(0);
  });

  test("should parse column names with spaces", () => {
    const ast = parseFormula("Table1[CAR ID]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "CAR ID",
      endCol: "CAR ID",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse current row reference with spaces in column names", () => {
    const ast = parseFormula("[@CUSTOMER ID]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBeUndefined();
    expect(ast.cols).toEqual({
      startCol: "CUSTOMER ID",
      endCol: "CUSTOMER ID",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse bracketed current row reference with spaces in column names", () => {
    const ast = parseFormula("Table1[@[CAR ID]]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "CAR ID",
      endCol: "CAR ID",
    });
    expect(ast.isCurrentRow).toBe(true);
  });

  test("should parse column range with mixed spaces and dashes", () => {
    const ast = parseFormula("Table1[CAR ID:ORDER-ID]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({
      startCol: "CAR ID",
      endCol: "ORDER-ID",
    });
    expect(ast.isCurrentRow).toBe(false);
  });

  test("should parse INDEX+MATCH formula with spaces and mixed column name types", () => {
    const ast = parseFormula(
      "INDEX(Table1[CAR ID], MATCH([@[CUSTOMER_ID]], Table1[ORDER-ID],0))"
    );
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("INDEX");
    expect(ast.args).toHaveLength(2);

    // First argument: Table1[CAR ID]
    expect(ast.args[0].type).toBe("structured-reference");
    expect(ast.args[0].tableName).toBe("Table1");
    expect(ast.args[0].cols?.startCol).toBe("CAR ID");

    // Second argument: MATCH function
    expect(ast.args[1].type).toBe("function");
    expect(ast.args[1].name).toBe("MATCH");
    expect(ast.args[1].args).toHaveLength(3);

    // MATCH first argument: [@[CUSTOMER_ID]]
    expect(ast.args[1].args[0].type).toBe("structured-reference");
    expect(ast.args[1].args[0].tableName).toBeUndefined();
    expect(ast.args[1].args[0].isCurrentRow).toBe(true);
    expect(ast.args[1].args[0].cols?.startCol).toBe("CUSTOMER_ID");

    // MATCH second argument: Table1[ORDER-ID]
    expect(ast.args[1].args[1].type).toBe("structured-reference");
    expect(ast.args[1].args[1].tableName).toBe("Table1");
    expect(ast.args[1].args[1].cols?.startCol).toBe("ORDER-ID");

    // MATCH third argument: 0
    expect(ast.args[1].args[2].type).toBe("value");
    expect(ast.args[1].args[2].value.value).toBe(0);
  });
});

describe("Parser - Complex Formulas with New Syntax", () => {
  test("should parse formula with all new features", () => {
    const ast = parseFormula(
      "SUM(Sheet1:Sheet3!A1, Table1[Sales], SEQUENCE(INFINITY))"
    );
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
    expect(ast.args[0]?.type).toBe("3d-range");
    expect(ast.args[1]?.type).toBe("structured-reference");
    expect(ast.args[2]?.type).toBe("function");
    expect(ast.args[2].args[0]?.type).toBe("infinity");
  });

  test("should parse nested structured references", () => {
    const ast = parseFormula("VLOOKUP([@ID], ProductTable[#All], 2, FALSE)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("VLOOKUP");
    expect(ast.args[0]?.type).toBe("structured-reference");
    expect(ast.args[1]?.type).toBe("structured-reference");
    expect(ast.args[0].isCurrentRow).toBe(true);
    expect(ast.args[0].cols).toEqual({ startCol: "ID", endCol: "ID" });
    expect(ast.args[1].selector).toBe("#All");
    expect(ast.args[1].cols).toBeUndefined();
  });

  test("should parse current row references with spaces in column names", () => {
    // Test [@[column name]] syntax for columns with spaces
    const ast1 = parseFormula("[@[bla bla]]");
    expect(ast1.type).toBe("structured-reference");
    expect(ast1.tableName).toBeUndefined();
    expect(ast1.isCurrentRow).toBe(true);
    expect(ast1.cols).toEqual({ startCol: "bla bla", endCol: "bla bla" });

    // Test Table1[@[column name]] syntax
    const ast2 = parseFormula("Table1[@[Net Sales]]");
    expect(ast2.type).toBe("structured-reference");
    expect(ast2.tableName).toBe("Table1");
    expect(ast2.isCurrentRow).toBe(true);
    expect(ast2.cols).toEqual({ startCol: "Net Sales", endCol: "Net Sales" });

    // Test column range with spaces [@[col1]:[col2]]
    const ast3 = parseFormula("Table1[@[Net Sales]:[Gross Profit]]");
    expect(ast3.type).toBe("structured-reference");
    expect(ast3.tableName).toBe("Table1");
    expect(ast3.isCurrentRow).toBe(true);
    expect(ast3.cols).toEqual({
      startCol: "Net Sales",
      endCol: "Gross Profit",
    });

    // Test simple current row range with spaces
    const ast4 = parseFormula("[@Net Sales:Gross Profit]");
    expect(ast4.type).toBe("structured-reference");
    expect(ast4.tableName).toBeUndefined();
    expect(ast4.isCurrentRow).toBe(true);
    expect(ast4.cols).toEqual({
      startCol: "Net Sales",
      endCol: "Gross Profit",
    });
  });
});

describe("Parser - Open-Ended Ranges", () => {
  test("should parse A5:INFINITY (both row and column unbounded)", () => {
    const ast = parseFormula("A5:INFINITY");
    expect(ast.type).toBe("range");
    expect(ast.range).toEqual({
      start: { col: 0, row: 4 }, // A5 -> (0, 4) in 0-based indexing
      end: {
        col: { type: "infinity", sign: "positive" },
        row: { type: "infinity", sign: "positive" },
      },
    });
    expect(ast.isAbsolute).toEqual({
      start: { col: false, row: false },
      end: { col: false, row: false },
    });
  });

  test("should parse A5:D (open down only - bounded columns)", () => {
    const ast = parseFormula("A5:D");
    expect(ast.type).toBe("range");
    expect(ast.range).toEqual({
      start: { col: 0, row: 4 }, // A5 -> (0, 4) in 0-based indexing
      end: {
        col: { type: "number", value: 3 }, // D -> column 3
        row: { type: "infinity", sign: "positive" },
      },
    });
    expect(ast.isAbsolute).toEqual({
      start: { col: false, row: false },
      end: { col: false, row: false },
    });
  });

  test("should parse A5:15 (open right only - bounded rows)", () => {
    const ast = parseFormula("A5:15");
    expect(ast.type).toBe("range");
    expect(ast.range).toEqual({
      start: { col: 0, row: 4 }, // A5 -> (0, 4) in 0-based indexing
      end: {
        col: { type: "infinity", sign: "positive" },
        row: { type: "number", value: 14 }, // Row 15 -> 14 in 0-based indexing
      },
    });
    expect(ast.isAbsolute).toEqual({
      start: { col: false, row: false },
      end: { col: false, row: false },
    });
  });

  test("should parse absolute open-ended ranges", () => {
    // $A$5:INFINITY
    const ast1 = parseFormula("$A$5:INFINITY");
    expect(ast1.type).toBe("range");
    expect(ast1.isAbsolute).toEqual({
      start: { col: true, row: true },
      end: { col: false, row: false }, // INFINITY is never absolute
    });

    // $A$5:$D
    const ast2 = parseFormula("$A$5:$D");
    expect(ast2.type).toBe("range");
    expect(ast2.isAbsolute).toEqual({
      start: { col: true, row: true },
      end: { col: true, row: false }, // Row is infinite, so not absolute
    });

    // $A$5:$15
    const ast3 = parseFormula("$A$5:$15");
    expect(ast3.type).toBe("range");
    expect(ast3.isAbsolute).toEqual({
      start: { col: true, row: true },
      end: { col: false, row: true }, // Column is infinite, so not absolute
    });
  });

  test("should parse mixed absolute open-ended ranges", () => {
    // A$5:D
    const ast1 = parseFormula("A$5:D");
    expect(ast1.type).toBe("range");
    expect(ast1.isAbsolute).toEqual({
      start: { col: false, row: true },
      end: { col: false, row: false },
    });

    // $A5:15
    const ast2 = parseFormula("$A5:15");
    expect(ast2.type).toBe("range");
    expect(ast2.isAbsolute).toEqual({
      start: { col: true, row: false },
      end: { col: false, row: false },
    });
  });

  test("should parse open-ended ranges with sheet names", () => {
    // Sheet1!A5:INFINITY
    const ast1 = parseFormula("Sheet1!A5:INFINITY");
    expect(ast1.type).toBe("range");
    expect(ast1.sheetName).toBe("Sheet1");
    expect(ast1.range.end.col.type).toBe("infinity");
    expect(ast1.range.end.row.type).toBe("infinity");

    // 'My Sheet'!A5:D
    const ast2 = parseFormula("'My Sheet'!A5:D");
    expect(ast2.type).toBe("range");
    expect(ast2.sheetName).toBe("My Sheet");
    expect(ast2.range.end.col.type).toBe("number");
    expect(ast2.range.end.row.type).toBe("infinity");

    // Sheet1!A5:15
    const ast3 = parseFormula("Sheet1!A5:15");
    expect(ast3.type).toBe("range");
    expect(ast3.sheetName).toBe("Sheet1");
    expect(ast3.range.end.col.type).toBe("infinity");
    expect(ast3.range.end.row.type).toBe("number");
  });

  test("should parse open-ended ranges in function calls", () => {
    const ast = parseFormula("SUM(A5:INFINITY, B1:D, C1:10)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);

    // First argument: A5:INFINITY
    expect(ast.args[0].type).toBe("range");
    expect(ast.args[0].range.end.col.type).toBe("infinity");
    expect(ast.args[0].range.end.row.type).toBe("infinity");

    // Second argument: B1:D
    expect(ast.args[1].type).toBe("range");
    expect(ast.args[1].range.end.col.type).toBe("number");
    expect(ast.args[1].range.end.row.type).toBe("infinity");

    // Third argument: C1:10
    expect(ast.args[2].type).toBe("range");
    expect(ast.args[2].range.end.col.type).toBe("infinity");
    expect(ast.args[2].range.end.row.type).toBe("number");
  });

  test("should parse complex open-ended range expressions", () => {
    const ast = parseFormula("(A5:D) + (B1:INFINITY)");
    expect(ast.type).toBe("binary-op");
    expect(ast.operator).toBe("+");
    expect(ast.left.type).toBe("range");
    expect(ast.right.type).toBe("range");
  });

  test("should handle edge cases and larger ranges", () => {
    // Large column range
    const ast1 = parseFormula("Z100:AA");
    expect(ast1.type).toBe("range");
    expect(ast1.range.start.col).toBe(25); // Z -> 25
    expect(ast1.range.start.row).toBe(99); // Row 100 -> 99 in 0-based
    expect(ast1.range.end.col.value).toBe(26); // AA -> 26
    expect(ast1.range.end.row.type).toBe("infinity");

    // Large row range
    const ast2 = parseFormula("AA100:1000");
    expect(ast2.type).toBe("range");
    expect(ast2.range.start.col).toBe(26); // AA -> 26
    expect(ast2.range.start.row).toBe(99); // Row 100 -> 99 in 0-based
    expect(ast2.range.end.col.type).toBe("infinity");
    expect(ast2.range.end.row.value).toBe(999); // Row 1000 -> 999 in 0-based

    // Very large range
    const ast3 = parseFormula("ZZ999:INFINITY");
    expect(ast3.type).toBe("range");
    expect(ast3.range.end.col.type).toBe("infinity");
    expect(ast3.range.end.row.type).toBe("infinity");
  });

  test("should distinguish open-ended ranges from normal ranges", () => {
    // Normal range A5:D5 should NOT be parsed as open-ended
    const ast1 = parseFormula("A5:D5");
    expect(ast1.type).toBe("range");
    expect(ast1.range.end.col.type).toBe("number");
    expect(ast1.range.end.row.type).toBe("number");
    expect(ast1.range.end.col.value).toBe(3); // D -> 3
    expect(ast1.range.end.row.value).toBe(4); // Row 5 -> 4 in 0-based

    // Normal range A5:A15 should NOT be parsed as open-ended
    const ast2 = parseFormula("A5:A15");
    expect(ast2.type).toBe("range");
    expect(ast2.range.end.col.type).toBe("number");
    expect(ast2.range.end.row.type).toBe("number");
  });
});

describe("Parser - Workbook References", () => {
  test("should parse workbook sheet alias", () => {
    const ast = parseFormula("[MyWorkbook]Sheet1");
    expect(ast.type).toBe("range");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.sheetName).toBe("Sheet1");
    // Should be equivalent to [MyWorkbook]Sheet1!A1:INFINITY
    expect(ast.range.start).toEqual({ col: 0, row: 0 });
    expect(ast.range.end.col.type).toBe("infinity");
    expect(ast.range.end.row.type).toBe("infinity");
  });

  test("should parse workbook sheet alias with spaces", () => {
    const ast = parseFormula("[My Workbook]'Sheet With Spaces'");
    expect(ast.type).toBe("range");
    expect(ast.workbookName).toBe("My Workbook");
    expect(ast.sheetName).toBe("Sheet With Spaces");
  });

  test("should parse workbook cell reference", () => {
    const ast = parseFormula("[MyWorkbook]Sheet1!A1");
    expect(ast.type).toBe("reference");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.sheetName).toBe("Sheet1");
    expect(ast.address).toEqual({ colIndex: 0, rowIndex: 0 });
  });

  test("should parse workbook absolute cell reference", () => {
    const ast = parseFormula("[MyWorkbook]'Sheet With Spaces'!$A$1");
    expect(ast.type).toBe("reference");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.sheetName).toBe("Sheet With Spaces");
    expect(ast.address).toEqual({ colIndex: 0, rowIndex: 0 });
    expect(ast.isAbsolute).toEqual({ col: true, row: true });
  });

  test("should parse workbook range reference", () => {
    const ast = parseFormula("[MyWorkbook]Sheet1!A1:C5");
    expect(ast.type).toBe("range");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.sheetName).toBe("Sheet1");
    expect(ast.range.start).toEqual({ col: 0, row: 0 });
    expect(ast.range.end.col.value).toBe(2); // Column C
    expect(ast.range.end.row.value).toBe(4); // Row 5 (0-based)
  });

  test("should parse workbook 3D range reference", () => {
    const ast = parseFormula("[MyWorkbook]Sheet1:Sheet3!A1:C5");
    expect(ast.type).toBe("3d-range");
    expect(ast.startSheet).toBe("Sheet1");
    expect(ast.endSheet).toBe("Sheet3");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.reference.type).toBe("range");
    // The inner reference should not have the workbook name (it's on the 3D range node)
    expect(ast.reference.workbookName).toBeUndefined();
  });

  test("should parse workbook named expression", () => {
    const ast = parseFormula("[MyWorkbook]Sheet1!TaxRate");
    expect(ast.type).toBe("named-expression");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.sheetName).toBe("Sheet1");
    expect(ast.name).toBe("TaxRate");
  });

  test("should parse workbook table reference", () => {
    const ast = parseFormula("[MyWorkbook]Sheet1!Table1[Column1]");
    expect(ast.type).toBe("structured-reference");
    expect(ast.workbookName).toBe("MyWorkbook");
    expect(ast.tableName).toBe("Table1");
    expect(ast.cols).toEqual({ startCol: "Column1", endCol: "Column1" });
  });

  test("should parse workbook infinite ranges", () => {
    // Column range
    const ast1 = parseFormula("[MyWorkbook]Sheet1!A:A");
    expect(ast1.type).toBe("range");
    expect(ast1.workbookName).toBe("MyWorkbook");
    expect(ast1.sheetName).toBe("Sheet1");
    expect(ast1.range.end.row.type).toBe("infinity");

    // Row range
    const ast2 = parseFormula("[MyWorkbook]Sheet1!1:1");
    expect(ast2.type).toBe("range");
    expect(ast2.workbookName).toBe("MyWorkbook");
    expect(ast2.sheetName).toBe("Sheet1");
    expect(ast2.range.end.col.type).toBe("infinity");

    // Open-ended range
    const ast3 = parseFormula("[MyWorkbook]Sheet1!A1:INFINITY");
    expect(ast3.type).toBe("range");
    expect(ast3.workbookName).toBe("MyWorkbook");
    expect(ast3.sheetName).toBe("Sheet1");
    expect(ast3.range.end.col.type).toBe("infinity");
    expect(ast3.range.end.row.type).toBe("infinity");
  });

  test("should parse workbook references in functions", () => {
    const ast = parseFormula("SUM([MyWorkbook]Sheet1!A1:A10)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args[0].type).toBe("range");
    expect(ast.args[0].workbookName).toBe("MyWorkbook");
    expect(ast.args[0].sheetName).toBe("Sheet1");
  });

  test("should parse complex workbook names", () => {
    // Workbook name with spaces, numbers, and special characters
    const testCases = [
      "[Budget 2024]Sheet1",
      "[My-Workbook]Sheet1",
      "[Workbook_v2]Sheet1",
      "[Report (Final)]Sheet1",
      "[Data=Analysis]Sheet1",
      "[50% Complete]Sheet1",
    ];

    testCases.forEach(formula => {
      expect(() => parseFormula(formula)).not.toThrow();
      const ast = parseFormula(formula);
      expect(ast.type).toBe("range");
      expect(ast.workbookName).toBeDefined();
      expect(ast.sheetName).toBe("Sheet1");
    });
  });

  test("should distinguish workbook references from bare column references", () => {
    // This should be parsed as a bare column reference, not a workbook reference
    const ast1 = parseFormula("[Column1]");
    expect(ast1.type).toBe("structured-reference");
    expect(ast1.workbookName).toBeUndefined();
    expect(ast1.cols?.startCol).toBe("Column1");

    // This should be parsed as a workbook reference
    const ast2 = parseFormula("[MyWorkbook]Sheet1");
    expect(ast2.type).toBe("range");
    expect(ast2.workbookName).toBe("MyWorkbook");
    expect(ast2.sheetName).toBe("Sheet1");
  });

  test("should handle mixed workbook and same-workbook references", () => {
    const ast = parseFormula("SUM([External]Sheet1!A1:A10, 'Local Sheet'!B1:B10)");
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(2);
    
    // First argument: external workbook reference
    expect(ast.args[0].workbookName).toBe("External");
    expect(ast.args[0].sheetName).toBe("Sheet1");
    
    // Second argument: same-workbook reference (no workbook name)
    expect(ast.args[1].workbookName).toBeUndefined();
    expect(ast.args[1].sheetName).toBe("Local Sheet");
  });
});

describe("Parser - Error Handling", () => {
  test("should throw on invalid syntax", () => {
    expect(() => parseFormula("=")).toThrow(ParseError);
    expect(() => parseFormula("1+")).toThrow(ParseError);
    expect(() => parseFormula("(1+2")).toThrow(ParseError);
    expect(() => parseFormula("SUM(")).toThrow(ParseError);
  });

  test("should throw on invalid tokens", () => {
    expect(() => parseFormula("A1]")).toThrow(ParseError);
    expect(() => parseFormula("[A1")).toThrow(ParseError);
  });

  test("should provide helpful error messages", () => {
    try {
      parseFormula("SUM(1,2,)");
    } catch (e) {
      expect(e).toBeInstanceOf(ParseError);
      expect((e as ParseError).message).toContain("Unexpected token");
    }
  });
});
