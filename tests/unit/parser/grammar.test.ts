import { test, expect, describe } from "bun:test";
import {
  OPERATOR_PRECEDENCE,
  OPERATOR_ASSOCIATIVITY,
  isBinaryOperator,
  getOperatorPrecedence,
  getOperatorAssociativity,
  compareOperatorPrecedence,
  SPECIAL_FUNCTIONS,
  VARIADIC_FUNCTIONS,
  REQUIRED_ARG_FUNCTIONS,
  RESERVED_KEYWORDS,
  isReservedKeyword,
  CELL_REFERENCE_PATTERNS,
  isValidColumn,
  isValidRow,
  parseCellReference,
  FUNCTION_CONSTRAINTS,
  getFunctionConstraints,
  validateFunctionArgCount
} from '../../../src/parser/grammar';

describe('Operator Precedence', () => {
  test('should define correct precedence levels', () => {
    // Comparison operators (lowest)
    expect(OPERATOR_PRECEDENCE['=']).toBe(1);
    expect(OPERATOR_PRECEDENCE['<>']).toBe(1);
    expect(OPERATOR_PRECEDENCE['<']).toBe(1);
    expect(OPERATOR_PRECEDENCE['>']).toBe(1);
    expect(OPERATOR_PRECEDENCE['<=']).toBe(1);
    expect(OPERATOR_PRECEDENCE['>=']).toBe(1);
    
    // Concatenation
    expect(OPERATOR_PRECEDENCE['&']).toBe(2);
    
    // Addition/Subtraction
    expect(OPERATOR_PRECEDENCE['+']).toBe(3);
    expect(OPERATOR_PRECEDENCE['-']).toBe(3);
    
    // Multiplication/Division
    expect(OPERATOR_PRECEDENCE['*']).toBe(4);
    expect(OPERATOR_PRECEDENCE['/']).toBe(4);
    
    // Exponentiation (highest)
    expect(OPERATOR_PRECEDENCE['^']).toBe(5);
  });

  test('should identify binary operators correctly', () => {
    expect(isBinaryOperator('+')).toBe(true);
    expect(isBinaryOperator('-')).toBe(true);
    expect(isBinaryOperator('*')).toBe(true);
    expect(isBinaryOperator('/')).toBe(true);
    expect(isBinaryOperator('^')).toBe(true);
    expect(isBinaryOperator('&')).toBe(true);
    expect(isBinaryOperator('=')).toBe(true);
    expect(isBinaryOperator('<>')).toBe(true);
    
    expect(isBinaryOperator('%')).toBe(false);
    expect(isBinaryOperator('!')).toBe(false);
    expect(isBinaryOperator('~')).toBe(false);
  });

  test('should get operator precedence correctly', () => {
    expect(getOperatorPrecedence('+')).toBe(3);
    expect(getOperatorPrecedence('*')).toBe(4);
    expect(getOperatorPrecedence('^')).toBe(5);
    expect(getOperatorPrecedence('unknown')).toBe(0);
  });

  test('should compare operator precedence correctly', () => {
    expect(compareOperatorPrecedence('*', '+')).toBeGreaterThan(0);
    expect(compareOperatorPrecedence('+', '*')).toBeLessThan(0);
    expect(compareOperatorPrecedence('+', '-')).toBe(0);
    expect(compareOperatorPrecedence('^', '*')).toBeGreaterThan(0);
  });
});

describe('Operator Associativity', () => {
  test('should define correct associativity', () => {
    // Most operators are left-associative
    expect(OPERATOR_ASSOCIATIVITY['+']).toBe('left');
    expect(OPERATOR_ASSOCIATIVITY['-']).toBe('left');
    expect(OPERATOR_ASSOCIATIVITY['*']).toBe('left');
    expect(OPERATOR_ASSOCIATIVITY['/']).toBe('left');
    expect(OPERATOR_ASSOCIATIVITY['&']).toBe('left');
    
    // Exponentiation is right-associative
    expect(OPERATOR_ASSOCIATIVITY['^']).toBe('right');
  });

  test('should get operator associativity correctly', () => {
    expect(getOperatorAssociativity('+')).toBe('left');
    expect(getOperatorAssociativity('^')).toBe('right');
    expect(getOperatorAssociativity('unknown')).toBe('left');
  });
});

describe('Special Functions', () => {
  test('should identify special functions that dont require parentheses', () => {
    expect(SPECIAL_FUNCTIONS.has('PI')).toBe(true);
    expect(SPECIAL_FUNCTIONS.has('TRUE')).toBe(true);
    expect(SPECIAL_FUNCTIONS.has('FALSE')).toBe(true);
    expect(SPECIAL_FUNCTIONS.has('NA')).toBe(true);
    
    expect(SPECIAL_FUNCTIONS.has('SUM')).toBe(false);
    expect(SPECIAL_FUNCTIONS.has('AVERAGE')).toBe(false);
  });
});

describe('Function Categories', () => {
  test('should identify variadic functions', () => {
    expect(VARIADIC_FUNCTIONS.has('SUM')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('PRODUCT')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('COUNT')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('MAX')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('MIN')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('AND')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('OR')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('CONCATENATE')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('CHOOSE')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('IFS')).toBe(true);
    expect(VARIADIC_FUNCTIONS.has('SWITCH')).toBe(true);
    
    expect(VARIADIC_FUNCTIONS.has('ABS')).toBe(false);
    expect(VARIADIC_FUNCTIONS.has('SQRT')).toBe(false);
  });

  test('should identify functions requiring at least one argument', () => {
    expect(REQUIRED_ARG_FUNCTIONS.has('SUM')).toBe(true);
    expect(REQUIRED_ARG_FUNCTIONS.has('AVERAGE')).toBe(true);
    expect(REQUIRED_ARG_FUNCTIONS.has('CONCATENATE')).toBe(true);
    
    expect(REQUIRED_ARG_FUNCTIONS.has('PI')).toBe(false);
    expect(REQUIRED_ARG_FUNCTIONS.has('TODAY')).toBe(false);
  });
});

describe('Reserved Keywords', () => {
  test('should identify reserved keywords', () => {
    expect(isReservedKeyword('TRUE')).toBe(true);
    expect(isReservedKeyword('FALSE')).toBe(true);
    expect(isReservedKeyword('NULL')).toBe(true);
    expect(isReservedKeyword('AND')).toBe(true);
    expect(isReservedKeyword('OR')).toBe(true);
    expect(isReservedKeyword('NOT')).toBe(true);
    
    // Case insensitive
    expect(isReservedKeyword('true')).toBe(true);
    expect(isReservedKeyword('False')).toBe(true);
    
    // Not reserved
    expect(isReservedKeyword('SUM')).toBe(false);
    expect(isReservedKeyword('MyVariable')).toBe(false);
  });
});

describe('Cell Reference Patterns', () => {
  test('should validate column references', () => {
    expect(isValidColumn('A')).toBe(true);
    expect(isValidColumn('Z')).toBe(true);
    expect(isValidColumn('AA')).toBe(true);
    expect(isValidColumn('XFD')).toBe(true);
    
    expect(isValidColumn('1')).toBe(false);
    expect(isValidColumn('A1')).toBe(false);
    expect(isValidColumn('')).toBe(false);
    expect(isValidColumn('$A')).toBe(false);
  });

  test('should validate row references', () => {
    expect(isValidRow('1')).toBe(true);
    expect(isValidRow('100')).toBe(true);
    expect(isValidRow('1048576')).toBe(true);
    
    expect(isValidRow('0')).toBe(false);
    expect(isValidRow('A')).toBe(false);
    expect(isValidRow('1A')).toBe(false);
    expect(isValidRow('')).toBe(false);
    expect(isValidRow('$1')).toBe(false);
  });

  test('should parse cell references correctly', () => {
    const a1 = parseCellReference('A1');
    expect(a1).toEqual({
      sheet: undefined,
      colAbsolute: false,
      col: 'A',
      rowAbsolute: false,
      row: '1'
    });

    const absolute = parseCellReference('$B$2');
    expect(absolute).toEqual({
      sheet: undefined,
      colAbsolute: true,
      col: 'B',
      rowAbsolute: true,
      row: '2'
    });

    const mixed1 = parseCellReference('$C3');
    expect(mixed1).toEqual({
      sheet: undefined,
      colAbsolute: true,
      col: 'C',
      rowAbsolute: false,
      row: '3'
    });

    const mixed2 = parseCellReference('D$4');
    expect(mixed2).toEqual({
      sheet: undefined,
      colAbsolute: false,
      col: 'D',
      rowAbsolute: true,
      row: '4'
    });
  });

  test('should parse sheet-qualified references', () => {
    const simple = parseCellReference('Sheet1!A1');
    expect(simple).toEqual({
      sheet: 'Sheet1',
      colAbsolute: false,
      col: 'A',
      rowAbsolute: false,
      row: '1'
    });

    const quoted = parseCellReference("'My Sheet'!B2");
    expect(quoted).toEqual({
      sheet: 'My Sheet',
      colAbsolute: false,
      col: 'B',
      rowAbsolute: false,
      row: '2'
    });

    const complex = parseCellReference("'Sheet-123'!$C$3");
    expect(complex).toEqual({
      sheet: 'Sheet-123',
      colAbsolute: true,
      col: 'C',
      rowAbsolute: true,
      row: '3'
    });
  });

  test('should return null for invalid references', () => {
    expect(parseCellReference('')).toBeNull();
    expect(parseCellReference('123')).toBeNull();
    expect(parseCellReference('ABC')).toBeNull();
    expect(parseCellReference('A0')).toBeNull();
    expect(parseCellReference('1A')).toBeNull();
    expect(parseCellReference('$$A1')).toBeNull();
  });
});

describe('Function Constraints', () => {
  test('should define constraints for math functions', () => {
    expect(getFunctionConstraints('ABS')).toEqual({ exactArgs: 1 });
    expect(getFunctionConstraints('POWER')).toEqual({ exactArgs: 2 });
    expect(getFunctionConstraints('LOG')).toEqual({ minArgs: 1, maxArgs: 2 });
    expect(getFunctionConstraints('ROUND')).toEqual({ exactArgs: 2 });
    expect(getFunctionConstraints('PI')).toEqual({ exactArgs: 0 });
  });

  test('should define constraints for statistical functions', () => {
    expect(getFunctionConstraints('SUM')).toEqual({ minArgs: 1 });
    expect(getFunctionConstraints('COUNT')).toEqual({ minArgs: 1 });
    expect(getFunctionConstraints('AVERAGE')).toEqual({ minArgs: 1 });
    expect(getFunctionConstraints('MEDIAN')).toEqual({ minArgs: 1 });
    expect(getFunctionConstraints('COVAR')).toEqual({ exactArgs: 2 });
  });

  test('should define constraints for logical functions', () => {
    expect(getFunctionConstraints('IF')).toEqual({ exactArgs: 3 });
    expect(getFunctionConstraints('AND')).toEqual({ minArgs: 1 });
    expect(getFunctionConstraints('OR')).toEqual({ minArgs: 1 });
    expect(getFunctionConstraints('NOT')).toEqual({ exactArgs: 1 });
    expect(getFunctionConstraints('IFS')).toEqual({ minArgs: 2 });
  });

  test('should define constraints for lookup functions', () => {
    expect(getFunctionConstraints('VLOOKUP')).toEqual({ minArgs: 3, maxArgs: 4 });
    expect(getFunctionConstraints('INDEX')).toEqual({ minArgs: 2, maxArgs: 3 });
    expect(getFunctionConstraints('MATCH')).toEqual({ minArgs: 2, maxArgs: 3 });
    expect(getFunctionConstraints('XLOOKUP')).toEqual({ minArgs: 3, maxArgs: 6 });
  });

  test('should return undefined for unknown functions', () => {
    expect(getFunctionConstraints('UNKNOWN_FUNCTION')).toBeUndefined();
  });
});

describe('Function Argument Validation', () => {
  test('should validate exact argument counts', () => {
    expect(validateFunctionArgCount('ABS', 1)).toBe(true);
    expect(validateFunctionArgCount('ABS', 0)).toBe(false);
    expect(validateFunctionArgCount('ABS', 2)).toBe(false);
    
    expect(validateFunctionArgCount('POWER', 2)).toBe(true);
    expect(validateFunctionArgCount('POWER', 1)).toBe(false);
    expect(validateFunctionArgCount('POWER', 3)).toBe(false);
  });

  test('should validate minimum argument counts', () => {
    expect(validateFunctionArgCount('SUM', 1)).toBe(true);
    expect(validateFunctionArgCount('SUM', 10)).toBe(true);
    expect(validateFunctionArgCount('SUM', 100)).toBe(true);
    expect(validateFunctionArgCount('SUM', 0)).toBe(false);
  });

  test('should validate argument ranges', () => {
    expect(validateFunctionArgCount('LOG', 1)).toBe(true);
    expect(validateFunctionArgCount('LOG', 2)).toBe(true);
    expect(validateFunctionArgCount('LOG', 0)).toBe(false);
    expect(validateFunctionArgCount('LOG', 3)).toBe(false);
    
    expect(validateFunctionArgCount('VLOOKUP', 3)).toBe(true);
    expect(validateFunctionArgCount('VLOOKUP', 4)).toBe(true);
    expect(validateFunctionArgCount('VLOOKUP', 2)).toBe(false);
    expect(validateFunctionArgCount('VLOOKUP', 5)).toBe(false);
  });

  test('should allow any argument count for unknown functions', () => {
    expect(validateFunctionArgCount('CUSTOM_FUNCTION', 0)).toBe(true);
    expect(validateFunctionArgCount('CUSTOM_FUNCTION', 5)).toBe(true);
    expect(validateFunctionArgCount('CUSTOM_FUNCTION', 100)).toBe(true);
  });

  test('should handle case-insensitive function names', () => {
    expect(validateFunctionArgCount('sum', 1)).toBe(true);
    expect(validateFunctionArgCount('SUM', 1)).toBe(true);
    expect(validateFunctionArgCount('Sum', 1)).toBe(true);
  });
});