import { describe, expect, test } from "bun:test";
import { FormulaError } from "../../../src/core/types";
import type { ArrayNode } from "../../../src/parser/ast";
import { ParseError, parseFormula } from "../../../src/parser/parser";
import { parse } from "path";

describe("Parser - Basic Values", () => {
  test("should parse numbers", () => {
    const ast = parseFormula("42");
    if (ast.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(ast.type).toBe("value");
    expect(ast.value).toEqual({ type: "number", value: 42 });

    const decimal = parseFormula("3.14159");
    if (decimal.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(decimal.value).toEqual({ type: "number", value: 3.14159 });

    const scientific = parseFormula("1.23E-4");
    if (scientific.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(scientific.value).toEqual({ type: "number", value: 0.000123 });
  });

  test("should parse strings", () => {
    const ast = parseFormula('"hello world"');
    expect(ast.type).toBe("value");
    if (ast.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(ast.value).toEqual({ type: "string", value: "hello world" });

    const escaped = parseFormula('"Say ""Hello"""');
    if (escaped.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(escaped.value).toEqual({ type: "string", value: 'Say "Hello"' });
  });

  test("should parse booleans", () => {
    const trueAst = parseFormula("TRUE");
    if (trueAst.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(trueAst.value).toEqual({ type: "boolean", value: true });

    const falseAst = parseFormula("FALSE");
    if (falseAst.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(falseAst.value).toEqual({ type: "boolean", value: false });
  });

  test("should parse error values", () => {
    const ast = parseFormula("#DIV/0!");
    if (ast.type !== "error") {
      throw new Error("Expected error node");
    }
    expect(ast.type).toBe("error");
    expect(ast.error).toBe(FormulaError.DIV0);
  });

  test("should parse empty formula", () => {
    const ast = parseFormula("");
    if (ast.type !== "empty") {
      throw new Error("Expected empty node, got " + ast.type);
    }
    expect(ast.type).toBe("empty");
  });
});

describe("Parser - Cell References", () => {
  test("should parse simple cell references", () => {
    const ast = parseFormula("A1");
    if (ast.type !== "reference") {
      throw new Error("Expected reference node");
    }
    expect(ast.type).toBe("reference");
    expect(ast.address).toEqual({ col: 0, row: 0 });
    expect(ast.isAbsolute).toEqual({ col: false, row: false });
  });

  test("should parse absolute cell references", () => {
    const ast = parseFormula("$B$2");
    if (ast.type !== "reference") {
      throw new Error("Expected reference node");
    }
    expect(ast.type).toBe("reference");
    expect(ast.address).toEqual({ col: 1, row: 1 });
    expect(ast.isAbsolute).toEqual({ col: true, row: true });
  });

  test("should parse mixed cell references", () => {
    const mixed1 = parseFormula("$C3");
    if (mixed1.type !== "reference") {
      throw new Error("Expected reference node");
    }
    expect(mixed1.address).toEqual({ col: 2, row: 2 });
    expect(mixed1.isAbsolute).toEqual({ col: true, row: false });

    const mixed2 = parseFormula("D$4");
    if (mixed2.type !== "reference") {
      throw new Error("Expected reference node");
    }
    expect(mixed2.address).toEqual({ col: 3, row: 3 });
    expect(mixed2.isAbsolute).toEqual({ col: false, row: true });
  });

  test("should parse range references", () => {
    const ast = parseFormula("A1:B2");
    if (ast.type !== "range") {
      throw new Error("Expected range node");
    }
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
    if (ast.type !== "range") {
      throw new Error("Expected range node");
    }
    expect(ast.type).toBe("range");
    expect(ast.isAbsolute).toEqual({
      start: { col: true, row: true },
      end: { col: true, row: true },
    });
  });

  test("should normalize range references", () => {
    const ast = parseFormula("B2:A1");
    if (ast.type !== "range") {
      throw new Error("Expected range node");
    }
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
    if (add.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(add.type).toBe("binary-op");
    expect(add.operator).toBe("+");
    expect(add.left.type).toBe("value");
    if (add.left.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(add.left.value).toEqual({ type: "number", value: 1 });
    expect(add.right.type).toBe("value");
    if (add.right.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(add.right.value).toEqual({
      type: "number",
      value: 2,
    });

    const mul = parseFormula("3*4");
    if (mul.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(mul.operator).toBe("*");

    const concat = parseFormula('"A"&"B"');
    if (concat.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(concat.operator).toBe("&");
  });

  test("should parse comparison operators", () => {
    const eq = parseFormula("A1=10");
    if (eq.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(eq.operator).toBe("=");

    const ne = parseFormula("A1<>10");
    if (ne.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ne.operator).toBe("<>");

    const gte = parseFormula("A1>=10");
    if (gte.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(gte.operator).toBe(">=");
  });

  test("should parse unary operators", () => {
    const neg = parseFormula("-5");
    if (neg.type !== "unary-op") {
      throw new Error("Expected unary op node");
    }
    expect(neg.type).toBe("unary-op");
    expect(neg.operator).toBe("-");
    expect(neg.operand.type).toBe("value");
    if (neg.operand.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(neg.operand.value).toEqual({
      type: "number",
      value: 5,
    });

    const pos = parseFormula("+5");
    if (pos.type !== "unary-op") {
      throw new Error("Expected unary op node");
    }
    expect(pos.operator).toBe("+");
  });

  test("should parse percentage operator", () => {
    const pct = parseFormula("50%");
    if (pct.type !== "unary-op") {
      throw new Error("Expected unary op node");
    }
    expect(pct.type).toBe("unary-op");
    expect(pct.operator).toBe("%");
    expect(pct.operand.type).toBe("value");
    if (pct.operand.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(pct.operand.value).toEqual({
      type: "number",
      value: 50,
    });
  });

  test("should respect operator precedence", () => {
    const ast = parseFormula("1+2*3");
    if (ast.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ast.operator).toBe("+");
    expect(ast.left.type).toBe("value");
    if (ast.left.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(ast.left.value).toEqual({ type: "number", value: 1 });
    expect(ast.right.type).toBe("binary-op");
    if (ast.right.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ast.right.operator).toBe("*");

    const ast2 = parseFormula("2^3^4");
    if (ast2.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ast2.operator).toBe("^");
    expect(ast2.left.type).toBe("value");
    if (ast2.left.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(ast2.left.value).toEqual({
      type: "number",
      value: 2,
    });
    // Right associative - should parse as 2^(3^4)
    expect(ast2.right.type).toBe("binary-op");
    if (ast2.right.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ast2.right.operator).toBe("^");
  });

  test("should handle parentheses", () => {
    const ast = parseFormula("(1+2)*3");
    if (ast.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ast.operator).toBe("*");
    expect(ast.left.type).toBe("binary-op");
    if (ast.left.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(ast.left.operator).toBe("+");
    expect(ast.right.type).toBe("value");
    if (ast.right.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(ast.right.value).toEqual({
      type: "number",
      value: 3,
    });
  });
});

describe("Parser - Functions", () => {
  test("should parse function calls with no arguments", () => {
    const ast = parseFormula("PI()");
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("PI");
    expect(ast.args).toEqual([]);
  });

  test("should parse function calls with arguments", () => {
    const ast = parseFormula("SUM(1,2,3)");
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
    const [arg1, arg2, arg3] = ast.args;
    if (!arg1 || arg1.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    if (!arg2 || arg2.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg2.type).toBe("value");
    expect(arg2.value).toEqual({
      type: "number",
      value: 2,
    });
    if (!arg3 || arg3.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg3.type).toBe("value");
    expect(arg3.value).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("should parse nested function calls", () => {
    const ast = parseFormula("SUM(1,ABS(-2),3)");
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
    const absNode = ast.args[1];
    if (!absNode || absNode.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(absNode.type).toBe("function");
    expect(absNode.name).toBe("ABS");
  });

  test("should parse functions with range arguments", () => {
    const ast = parseFormula("SUM(A1:A10)");
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(1);
    const rangeNode = ast.args[0];
    if (!rangeNode || rangeNode.type !== "range") {
      throw new Error("Expected range node");
    }
    expect(rangeNode.type).toBe("range");
  });

  test("should validate function argument counts", () => {
    expect(() => parseFormula("ABS()")).toThrow(ParseError);
    expect(() => parseFormula("ABS(1,2)")).toThrow(ParseError);
    expect(() => parseFormula("IF(1)")).toThrow(ParseError);
    expect(() => parseFormula("IF(1,2,3,4)")).toThrow(ParseError);
  });

  test("should normalize function names to uppercase", () => {
    const ast = parseFormula("sum(1,2)");
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
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
    if (!firstElement) {
      throw new Error("Expected first element");
    }
    const [arg1, arg2, arg3] = firstElement;
    if (!arg1 || arg1.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    if (!arg2 || arg2.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg2.type).toBe("value");
    expect(arg2.value).toEqual({
      type: "number",
      value: 2,
    });
    if (!arg3 || arg3.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg3.type).toBe("value");
    expect(arg3.value).toEqual({
      type: "number",
      value: 3,
    });
  });

  test("should parse 2D array literals", () => {
    const ast = parseFormula("{1,2;3,4}");
    expect(ast.type).toBe("array");
    if (ast.type !== "array") {
      throw new Error("Expected array node");
    }
    expect(ast.elements).toHaveLength(2);
    expect(ast.elements[0]).toHaveLength(2);
    expect(ast.elements[1]).toHaveLength(2);
    const firstElement = ast.elements[0];
    if (!firstElement) {
      throw new Error("Expected first element");
    }
    const [arg1, arg2] = firstElement;
    if (!arg1 || arg1.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    const secondElement = ast.elements[1];
    if (!secondElement) {
      throw new Error("Expected value node");
    }
    const last = secondElement[1];
    if (!last || last.type !== "value") {
      throw new Error("Expected value node");
    }
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
    if (!firstElement) {
      throw new Error("Expected first element");
    }
    const [arg1, arg2, arg3] = firstElement;
    if (!arg1 || arg1.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg1.type).toBe("value");
    expect(arg1.value).toEqual({
      type: "number",
      value: 1,
    });
    if (!arg2 || arg2.type !== "value") {
      throw new Error("Expected value node");
    }
    expect(arg2.type).toBe("value");
    expect(arg2.value).toEqual({
      type: "string",
      value: "text",
    });
    if (!arg3 || arg3.type !== "value") {
      throw new Error("Expected value node");
    }
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
    if (ast.type !== "array") {
      throw new Error("Expected array node");
    }
    expect(ast.type).toBe("array");
    expect(ast.elements).toHaveLength(1);
    expect(ast.elements[0]).toHaveLength(1);
    const emptyNode = ast.elements[0]?.[0];
    if (!emptyNode || emptyNode.type !== "empty") {
      throw new Error("Expected empty node");
    }
    expect(emptyNode.type).toBe("empty");
  });
});

describe("Parser - Named Expressions", () => {
  test("should parse named expressions", () => {
    const ast = parseFormula("TaxRate");
    if (ast.type !== "named-expression") {
      throw new Error("Expected named expression node");
    }
    expect(ast.type).toBe("named-expression");
    expect(ast.name).toBe("TaxRate");
    expect(ast.sheetName).toBeUndefined();
  });

  test("should distinguish between functions and named expressions", () => {
    const func = parseFormula("SUM(1,2)");
    if (func.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(func.type).toBe("function");

    const name = parseFormula("SUM");
    if (name.type !== "named-expression") {
      throw new Error("Expected named expression node");
    }
    expect(name.type).toBe("named-expression");
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
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(ast.type).toBe("function");
    expect(ast.name).toBe("IF");
    expect(ast.args).toHaveLength(3);

    const condition = ast.args[0];
    if (!condition || condition.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
    expect(condition.type).toBe("binary-op");
    expect(condition.operator).toBe(">");
  });

  test("should parse array operations", () => {
    const formula = "SUM(A1:A10*B1:B10)";
    const ast = parseFormula(formula);
    if (ast.type !== "function") {
      throw new Error("Expected function node");
    }
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(1);

    const mul = ast.args[0];
    if (!mul || mul.type !== "binary-op") {
      throw new Error("Expected binary op node");
    }
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
    if (ast.type === "function") {
      expect(ast.name).toBe("SEQUENCE");
      expect(ast.args[0]?.type).toBe("infinity");
    }
  });

  test("should parse arithmetic with INFINITY", () => {
    const ast = parseFormula("1/INFINITY");
    expect(ast.type).toBe("binary-op");
    if (ast.type === "binary-op") {
      expect(ast.operator).toBe("/");
      expect(ast.right.type).toBe("infinity");
    }
  });
});

describe("Parser - 3D Ranges", () => {
  test("should parse basic 3D range", () => {
    const ast = parseFormula("Sheet1:Sheet3!A1");
    expect(ast.type).toBe("3d-range");
    if (ast.type === "3d-range") {
      expect(ast.startSheet).toBe("Sheet1");
      expect(ast.endSheet).toBe("Sheet3");
      expect(ast.reference.type).toBe("reference");
    }
  });

  test("should parse 3D range with quoted sheet names", () => {
    const ast = parseFormula("'Sheet 1':'Sheet 3'!A1");
    expect(ast.type).toBe("3d-range");
    if (ast.type === "3d-range") {
      expect(ast.startSheet).toBe("Sheet 1");
      expect(ast.endSheet).toBe("Sheet 3");
      expect(ast.reference.type).toBe("reference");
    }
  });

  test("should parse 3D range with cell range", () => {
    const ast = parseFormula("Sheet1:Sheet5!A1:B10");
    expect(ast.type).toBe("3d-range");
    if (ast.type === "3d-range") {
      expect(ast.startSheet).toBe("Sheet1");
      expect(ast.endSheet).toBe("Sheet5");
      expect(ast.reference.type).toBe("range");
    }
  });

  test("should parse 3D range in function", () => {
    const ast = parseFormula("AVERAGE(Sheet1:Sheet3!A1)");
    expect(ast.type).toBe("function");
    if (ast.type === "function") {
      expect(ast.name).toBe("AVERAGE");
      expect(ast.args[0]?.type).toBe("3d-range");
    }
  });
});

describe("Parser - Structured References", () => {
  test("should parse simple table reference", () => {
    const ast = parseFormula("Table1[Sales]");
    expect(ast.type).toBe("structured-reference");
    if (ast.type === "structured-reference") {
      expect(ast.tableName).toBe("Table1");
      expect(ast.columnName).toBe("Sales");
      expect(ast.isCurrentRow).toBe(false);
    }
  });

  test("should parse current row reference", () => {
    const ast = parseFormula("[@Sales]");
    expect(ast.type).toBe("structured-reference");
    if (ast.type === "structured-reference") {
      expect(ast.tableName).toBe("");
      expect(ast.columnName).toBe("Sales");
      expect(ast.isCurrentRow).toBe(true);
    }
  });

  test("should parse table with current row reference", () => {
    const ast = parseFormula("Table1[@Sales]");
    expect(ast.type).toBe("structured-reference");
    if (ast.type === "structured-reference") {
      expect(ast.tableName).toBe("Table1");
      expect(ast.columnName).toBe("Sales");
      expect(ast.isCurrentRow).toBe(true);
    }
  });

  test("should parse table with selector", () => {
    const ast = parseFormula("Table1[[#Headers],[Sales]]");
    expect(ast.type).toBe("structured-reference");
    if (ast.type === "structured-reference") {
      expect(ast.tableName).toBe("Table1");
      expect(ast.columnName).toBe("Sales");
      expect(ast.selector).toBe("#Headers");
    }
  });

  test("should parse table reference in formula", () => {
    const ast = parseFormula("SUM(Table1[Sales])");
    expect(ast.type).toBe("function");
    if (ast.type === "function") {
      expect(ast.name).toBe("SUM");
      expect(ast.args[0]?.type).toBe("structured-reference");
    }
  });

  test("should parse current row reference in arithmetic", () => {
    const ast = parseFormula("[@Price]*[@Quantity]");
    expect(ast.type).toBe("binary-op");
    if (ast.type === "binary-op") {
      expect(ast.operator).toBe("*");
      expect(ast.left.type).toBe("structured-reference");
      expect(ast.right.type).toBe("structured-reference");
      if (ast.left.type === "structured-reference") {
        expect(ast.left.isCurrentRow).toBe(true);
        expect(ast.left.columnName).toBe("Price");
      }
      if (ast.right.type === "structured-reference") {
        expect(ast.right.isCurrentRow).toBe(true);
        expect(ast.right.columnName).toBe("Quantity");
      }
    }
  });

  test("should parse table selector only", () => {
    const ast = parseFormula("Table1[[#Data]]");
    expect(ast.type).toBe("structured-reference");
    if (ast.type === "structured-reference") {
      expect(ast.tableName).toBe("Table1");
      expect(ast.selector).toBe("#Data");
      expect(ast.columnName).toBeUndefined();
    }
  });
});

describe("Parser - Complex Formulas with New Syntax", () => {
  test("should parse formula with all new features", () => {
    const ast = parseFormula("SUM(Sheet1:Sheet3!A1, Table1[Sales], SEQUENCE(INFINITY))");
    expect(ast.type).toBe("function");
    if (ast.type === "function") {
      expect(ast.name).toBe("SUM");
      expect(ast.args).toHaveLength(3);
      expect(ast.args[0]?.type).toBe("3d-range");
      expect(ast.args[1]?.type).toBe("structured-reference");
      expect(ast.args[2]?.type).toBe("function");
      if (ast.args[2]?.type === "function") {
        expect(ast.args[2].args[0]?.type).toBe("infinity");
      }
    }
  });

  test("should parse nested structured references", () => {
    const ast = parseFormula("VLOOKUP([@ID], ProductTable[#All], 2, FALSE)");
    expect(ast.type).toBe("function");
    if (ast.type === "function") {
      expect(ast.name).toBe("VLOOKUP");
      expect(ast.args[0]?.type).toBe("structured-reference");
      expect(ast.args[1]?.type).toBe("structured-reference");
      if (ast.args[0]?.type === "structured-reference") {
        expect(ast.args[0].isCurrentRow).toBe(true);
      }
      if (ast.args[1]?.type === "structured-reference") {
        expect(ast.args[1].selector).toBe("#All");
      }
    }
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
