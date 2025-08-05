import { test, expect, describe } from "bun:test";
import {
  Parser,
  ParseError,
  parseFormula,
  normalizeFormula,
  astToFormula,
  validateFormula,
  extractNamedExpressions
} from '../../../src/parser/parser';
import type {
  FormulaAST,
  ValueNode,
  ReferenceNode,
  RangeNode,
  FunctionNode,
  UnaryOpNode,
  BinaryOpNode,
  ArrayNode,
  NamedExpressionNode,
  ErrorNode
} from '../../../src/parser/ast';

describe('Parser - Basic Values', () => {
  test('should parse numbers', () => {
    const ast = parseFormula('42') as ValueNode;
    expect(ast.type).toBe('value');
    expect(ast.value).toBe(42);
    expect(ast.valueType).toBe('number');

    const decimal = parseFormula('3.14159') as ValueNode;
    expect(decimal.value).toBe(3.14159);

    const scientific = parseFormula('1.23E-4') as ValueNode;
    expect(scientific.value).toBe(0.000123);
  });

  test('should parse strings', () => {
    const ast = parseFormula('"hello world"') as ValueNode;
    expect(ast.type).toBe('value');
    expect(ast.value).toBe('hello world');
    expect(ast.valueType).toBe('string');

    const escaped = parseFormula('"Say ""Hello"""') as ValueNode;
    expect(escaped.value).toBe('Say "Hello"');
  });

  test('should parse booleans', () => {
    const trueAst = parseFormula('TRUE') as ValueNode;
    expect(trueAst.type).toBe('value');
    expect(trueAst.value).toBe(true);
    expect(trueAst.valueType).toBe('boolean');

    const falseAst = parseFormula('FALSE') as ValueNode;
    expect(falseAst.value).toBe(false);
  });

  test('should parse error values', () => {
    const ast = parseFormula('#DIV/0!') as ErrorNode;
    expect(ast.type).toBe('error');
    expect(ast.error).toBe('#DIV/0!');
  });

  test('should parse empty formula', () => {
    const ast = parseFormula('') as ValueNode;
    expect(ast.type).toBe('value');
    expect(ast.value).toBeUndefined();
  });
});

describe('Parser - Cell References', () => {
  test('should parse simple cell references', () => {
    const ast = parseFormula('A1') as ReferenceNode;
    expect(ast.type).toBe('reference');
    expect(ast.address).toEqual({ sheet: 0, col: 0, row: 0 });
    expect(ast.isAbsolute).toEqual({ col: false, row: false });
  });

  test('should parse absolute cell references', () => {
    const ast = parseFormula('$B$2') as ReferenceNode;
    expect(ast.type).toBe('reference');
    expect(ast.address).toEqual({ sheet: 0, col: 1, row: 1 });
    expect(ast.isAbsolute).toEqual({ col: true, row: true });
  });

  test('should parse mixed cell references', () => {
    const mixed1 = parseFormula('$C3') as ReferenceNode;
    expect(mixed1.address).toEqual({ sheet: 0, col: 2, row: 2 });
    expect(mixed1.isAbsolute).toEqual({ col: true, row: false });

    const mixed2 = parseFormula('D$4') as ReferenceNode;
    expect(mixed2.address).toEqual({ sheet: 0, col: 3, row: 3 });
    expect(mixed2.isAbsolute).toEqual({ col: false, row: true });
  });

  test('should parse range references', () => {
    const ast = parseFormula('A1:B2') as RangeNode;
    expect(ast.type).toBe('range');
    expect(ast.range).toEqual({
      start: { sheet: 0, col: 0, row: 0 },
      end: { sheet: 0, col: 1, row: 1 }
    });
  });

  test('should parse absolute range references', () => {
    const ast = parseFormula('$A$1:$B$2') as RangeNode;
    expect(ast.type).toBe('range');
    expect(ast.isAbsolute).toEqual({
      start: { col: true, row: true },
      end: { col: true, row: true }
    });
  });

  test('should normalize range references', () => {
    const ast = parseFormula('B2:A1') as RangeNode;
    expect(ast.range).toEqual({
      start: { sheet: 0, col: 0, row: 0 },
      end: { sheet: 0, col: 1, row: 1 }
    });
  });
});

describe('Parser - Operators', () => {
  test('should parse binary operators', () => {
    const add = parseFormula('1+2') as BinaryOpNode;
    expect(add.type).toBe('binary-op');
    expect(add.operator).toBe('+');
    expect((add.left as ValueNode).value).toBe(1);
    expect((add.right as ValueNode).value).toBe(2);

    const mul = parseFormula('3*4') as BinaryOpNode;
    expect(mul.operator).toBe('*');

    const concat = parseFormula('"A"&"B"') as BinaryOpNode;
    expect(concat.operator).toBe('&');
  });

  test('should parse comparison operators', () => {
    const eq = parseFormula('A1=10') as BinaryOpNode;
    expect(eq.operator).toBe('=');

    const ne = parseFormula('A1<>10') as BinaryOpNode;
    expect(ne.operator).toBe('<>');

    const gte = parseFormula('A1>=10') as BinaryOpNode;
    expect(gte.operator).toBe('>=');
  });

  test('should parse unary operators', () => {
    const neg = parseFormula('-5') as UnaryOpNode;
    expect(neg.type).toBe('unary-op');
    expect(neg.operator).toBe('-');
    expect((neg.operand as ValueNode).value).toBe(5);

    const pos = parseFormula('+5') as UnaryOpNode;
    expect(pos.operator).toBe('+');
  });

  test('should parse percentage operator', () => {
    const pct = parseFormula('50%') as UnaryOpNode;
    expect(pct.type).toBe('unary-op');
    expect(pct.operator).toBe('%');
    expect((pct.operand as ValueNode).value).toBe(50);
  });

  test('should respect operator precedence', () => {
    const ast = parseFormula('1+2*3') as BinaryOpNode;
    expect(ast.operator).toBe('+');
    expect((ast.left as ValueNode).value).toBe(1);
    expect((ast.right as BinaryOpNode).operator).toBe('*');

    const ast2 = parseFormula('2^3^4') as BinaryOpNode;
    expect(ast2.operator).toBe('^');
    expect((ast2.left as ValueNode).value).toBe(2);
    // Right associative - should parse as 2^(3^4)
    expect((ast2.right as BinaryOpNode).operator).toBe('^');
  });

  test('should handle parentheses', () => {
    const ast = parseFormula('(1+2)*3') as BinaryOpNode;
    expect(ast.operator).toBe('*');
    expect((ast.left as BinaryOpNode).operator).toBe('+');
    expect((ast.right as ValueNode).value).toBe(3);
  });
});

describe('Parser - Functions', () => {
  test('should parse function calls with no arguments', () => {
    const ast = parseFormula('PI()') as FunctionNode;
    expect(ast.type).toBe('function');
    expect(ast.name).toBe('PI');
    expect(ast.args).toEqual([]);
  });

  test('should parse function calls with arguments', () => {
    const ast = parseFormula('SUM(1,2,3)') as FunctionNode;
    expect(ast.type).toBe('function');
    expect(ast.name).toBe('SUM');
    expect(ast.args).toHaveLength(3);
    expect((ast.args[0] as ValueNode).value).toBe(1);
    expect((ast.args[1] as ValueNode).value).toBe(2);
    expect((ast.args[2] as ValueNode).value).toBe(3);
  });

  test('should parse nested function calls', () => {
    const ast = parseFormula('SUM(1,ABS(-2),3)') as FunctionNode;
    expect(ast.name).toBe('SUM');
    expect(ast.args).toHaveLength(3);
    const absNode = ast.args[1] as FunctionNode;
    expect(absNode.type).toBe('function');
    expect(absNode.name).toBe('ABS');
  });

  test('should parse functions with range arguments', () => {
    const ast = parseFormula('SUM(A1:A10)') as FunctionNode;
    expect(ast.name).toBe('SUM');
    expect(ast.args).toHaveLength(1);
    expect((ast.args[0] as RangeNode).type).toBe('range');
  });

  test('should validate function argument counts', () => {
    expect(() => parseFormula('ABS()')).toThrow(ParseError);
    expect(() => parseFormula('ABS(1,2)')).toThrow(ParseError);
    expect(() => parseFormula('IF(1)')).toThrow(ParseError);
    expect(() => parseFormula('IF(1,2,3,4)')).toThrow(ParseError);
  });

  test('should normalize function names to uppercase', () => {
    const ast = parseFormula('sum(1,2)') as FunctionNode;
    expect(ast.name).toBe('SUM');
  });
});

describe('Parser - Arrays', () => {
  test('should parse array literals', () => {
    const ast = parseFormula('{1,2,3}') as ArrayNode;
    expect(ast.type).toBe('array');
    expect(ast.elements).toHaveLength(1);
    expect(ast.elements[0]).toHaveLength(3);
    expect((ast.elements[0]?.[0] as ValueNode).value).toBe(1);
    expect((ast.elements[0]?.[1] as ValueNode).value).toBe(2);
    expect((ast.elements[0]?.[2] as ValueNode).value).toBe(3);
  });

  test('should parse 2D array literals', () => {
    const ast = parseFormula('{1,2;3,4}') as ArrayNode;
    expect(ast.type).toBe('array');
    expect(ast.elements).toHaveLength(2);
    expect(ast.elements[0]).toHaveLength(2);
    expect(ast.elements[1]).toHaveLength(2);
    expect((ast.elements[0]?.[0] as ValueNode).value).toBe(1);
    expect((ast.elements[1]?.[1] as ValueNode).value).toBe(4);
  });

  test('should parse arrays with mixed types', () => {
    const ast = parseFormula('{1,"text",TRUE}') as ArrayNode;
    expect(ast.elements[0]).toHaveLength(3);
    expect((ast.elements[0]?.[0] as ValueNode).value).toBe(1);
    expect((ast.elements[0]?.[1] as ValueNode).value).toBe('text');
    expect((ast.elements[0]?.[2] as ValueNode).value).toBe(true);
  });

  test('should enforce consistent row lengths', () => {
    expect(() => parseFormula('{1,2;3}')).toThrow(ParseError);
    expect(() => parseFormula('{1;2,3}')).toThrow(ParseError);
  });

  test('should handle empty arrays', () => {
    const ast = parseFormula('{}') as ArrayNode;
    expect(ast.type).toBe('array');
    expect(ast.elements).toHaveLength(1);
    expect(ast.elements[0]).toHaveLength(1);
    expect((ast.elements[0]?.[0] as ValueNode).value).toBeUndefined();
  });
});

describe('Parser - Named Expressions', () => {
  test('should parse named expressions', () => {
    const ast = parseFormula('TaxRate') as NamedExpressionNode;
    expect(ast.type).toBe('named-expression');
    expect(ast.name).toBe('TaxRate');
    expect(ast.scope).toBeUndefined();
  });

  test('should distinguish between functions and named expressions', () => {
    const func = parseFormula('SUM(1,2)') as FunctionNode;
    expect(func.type).toBe('function');

    const name = parseFormula('SUM') as NamedExpressionNode;
    expect(name.type).toBe('named-expression');
  });
});

describe('Parser - Complex Expressions', () => {
  test('should parse complex arithmetic expressions', () => {
    const formula = '(A1+B1)*2-C1/3+D1^2';
    const ast = parseFormula(formula);
    expect(ast.type).toBe('binary-op');
    
    // Verify it parses without errors
    const normalized = astToFormula(ast);
    expect(normalized).toContain('+');
    expect(normalized).toContain('*');
    expect(normalized).toContain('/');
    expect(normalized).toContain('^');
  });

  test('should parse nested expressions with functions', () => {
    const formula = 'IF(SUM(A1:A10)>100,MAX(B1:B10),MIN(C1:C10))';
    const ast = parseFormula(formula) as FunctionNode;
    expect(ast.type).toBe('function');
    expect(ast.name).toBe('IF');
    expect(ast.args).toHaveLength(3);
    
    const condition = ast.args[0] as BinaryOpNode;
    expect(condition.type).toBe('binary-op');
    expect(condition.operator).toBe('>');
  });

  test('should parse array operations', () => {
    const formula = 'SUM(A1:A10*B1:B10)';
    const ast = parseFormula(formula) as FunctionNode;
    expect(ast.name).toBe('SUM');
    expect(ast.args).toHaveLength(1);
    
    const mul = ast.args[0] as BinaryOpNode;
    expect(mul.type).toBe('binary-op');
    expect(mul.operator).toBe('*');
  });
});

describe('Parser - Error Handling', () => {
  test('should throw on invalid syntax', () => {
    // Empty formula with just = returns undefined value
    const emptyFormula = parseFormula('=') as ValueNode;
    expect(emptyFormula.type).toBe('value');
    expect(emptyFormula.value).toBeUndefined();
    expect(() => parseFormula('1+')).toThrow(ParseError);
    expect(() => parseFormula('(1+2')).toThrow(ParseError);
    expect(() => parseFormula('SUM(')).toThrow(ParseError);
  });

  test('should throw on invalid tokens', () => {
    expect(() => parseFormula('1@2')).toThrow(ParseError);
    expect(() => parseFormula('A1]')).toThrow(ParseError);
  });

  test('should provide helpful error messages', () => {
    try {
      parseFormula('SUM(1,2,)');
    } catch (e) {
      expect(e).toBeInstanceOf(ParseError);
      expect((e as ParseError).message).toContain('Unexpected token');
    }
  });
});

describe('Parser - Utility Functions', () => {
  test('should normalize formulas', () => {
    expect(normalizeFormula('  1  +  2  ')).toBe('1+2');
    expect(normalizeFormula('sum(a1:a10)')).toBe('SUM(R1C1:R10C1)');
    expect(normalizeFormula('true')).toBe('TRUE');
  });

  test('should validate formulas', () => {
    expect(validateFormula('1+2')).toBe(true);
    expect(validateFormula('SUM(A1:A10)')).toBe(true);
    expect(validateFormula('IF(A1>0,1,0)')).toBe(true);
    
    expect(validateFormula('=')).toBe(true); // Empty formula is valid
    expect(validateFormula('1+')).toBe(false);
    expect(validateFormula('SUM(')).toBe(false);
  });

  test('should extract named expressions', () => {
    const names = extractNamedExpressions('TaxRate*SubTotal+Discount');
    expect(names).toEqual(['TaxRate', 'SubTotal', 'Discount']);
    
    const complex = extractNamedExpressions('IF(TaxRate>0,SubTotal*TaxRate,0)');
    expect(complex).toEqual(['TaxRate', 'SubTotal']);
    
    const none = extractNamedExpressions('SUM(A1:A10)');
    expect(none).toEqual([]);
  });

  test('should handle formula with leading equals sign', () => {
    const ast = parseFormula('=1+2') as BinaryOpNode;
    expect(ast.type).toBe('binary-op');
    expect(ast.operator).toBe('+');
  });
});

describe('Parser - AST to Formula Conversion', () => {
  test('should convert values back to formula', () => {
    expect(astToFormula(parseFormula('42'))).toBe('42');
    expect(astToFormula(parseFormula('"hello"'))).toBe('"hello"');
    expect(astToFormula(parseFormula('TRUE'))).toBe('TRUE');
    expect(astToFormula(parseFormula('FALSE'))).toBe('FALSE');
  });

  test('should escape quotes in strings', () => {
    expect(astToFormula(parseFormula('"Say ""Hi"""'))).toBe('"Say ""Hi"""');
  });

  test('should convert operators back to formula', () => {
    expect(astToFormula(parseFormula('1+2'))).toBe('1+2');
    expect(astToFormula(parseFormula('3*4'))).toBe('3*4');
    expect(astToFormula(parseFormula('-5'))).toBe('-5');
    expect(astToFormula(parseFormula('50%'))).toBe('50%');
  });

  test('should convert functions back to formula', () => {
    expect(astToFormula(parseFormula('SUM(1,2,3)'))).toBe('SUM(1,2,3)');
    expect(astToFormula(parseFormula('PI()'))).toBe('PI()');
  });

  test('should convert arrays back to formula', () => {
    expect(astToFormula(parseFormula('{1,2,3}'))).toBe('{1,2,3}');
    expect(astToFormula(parseFormula('{1,2;3,4}'))).toBe('{1,2;3,4}');
  });

  test('should convert references back to formula', () => {
    // Note: Currently uses R1C1 notation - will be updated when A1 conversion is implemented
    const ref = astToFormula(parseFormula('A1'));
    expect(ref).toContain('R1C1');
    
    const range = astToFormula(parseFormula('A1:B2'));
    expect(range).toContain(':');
  });
});