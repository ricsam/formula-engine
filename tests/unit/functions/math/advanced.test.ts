import { test, expect, describe, beforeEach } from "bun:test";
import { Evaluator } from '../../../../src/evaluator/evaluator';
import { Parser } from '../../../../src/parser/parser';
import { DependencyGraph } from '../../../../src/evaluator/dependency-graph';
import { ErrorHandler } from '../../../../src/evaluator/error-handler';
import { functionRegistry } from '../../../../src/functions/index';
import type { EvaluationContext } from '../../../../src/evaluator/evaluator';
import type { CellValue, SimpleCellAddress, SimpleCellRange } from '../../../../src/core/types';

describe('Advanced Math Functions', () => {
  let evaluator: Evaluator;
  let errorHandler: ErrorHandler;
  let dependencyGraph: DependencyGraph;

  // Helper function to evaluate a formula
  const evaluateFormula = (formula: string): CellValue => {
    const ast = Parser.parse(formula, 0);
    
    const context: EvaluationContext = {
      currentSheet: 0,
      namedExpressions: new Map(),
      getCellValue: (address: SimpleCellAddress) => undefined,
      getRangeValues: (range: SimpleCellRange) => [[]],
      getFunction: (name: string) => functionRegistry.get(name),
      errorHandler,
      evaluationStack: new Set(),
    };
    
    const result = evaluator.evaluate(ast, context);
    return result.value as any;
  };

  beforeEach(() => {
    errorHandler = new ErrorHandler();
    dependencyGraph = new DependencyGraph();
    
    // Get all functions from the registry
    const functions = functionRegistry.getAllFunctions();
    
    evaluator = new Evaluator(dependencyGraph, functions, errorHandler);
  });

  describe('Basic Mathematical Functions', () => {
    test('ABS function', () => {
      expect(evaluateFormula('=ABS(5)')).toBe(5);
      expect(evaluateFormula('=ABS(-5)')).toBe(5);
      expect(evaluateFormula('=ABS(0)')).toBe(0);
      expect(evaluateFormula('=ABS(-3.14)')).toBe(3.14);
    });

    test('SIGN function', () => {
      expect(evaluateFormula('=SIGN(5)')).toBe(1);
      expect(evaluateFormula('=SIGN(-5)')).toBe(-1);
      expect(evaluateFormula('=SIGN(0)')).toBe(0);
      expect(evaluateFormula('=SIGN(3.14)')).toBe(1);
      expect(evaluateFormula('=SIGN(-2.5)')).toBe(-1);
    });

    test('SQRT function', () => {
      expect(evaluateFormula('=SQRT(4)')).toBe(2);
      expect(evaluateFormula('=SQRT(9)')).toBe(3);
      expect(evaluateFormula('=SQRT(0)')).toBe(0);
      expect(evaluateFormula('=SQRT(2)')).toBeCloseTo(1.414, 3);
    });

    test('SQRT with negative number returns error', () => {
      expect(evaluateFormula('=SQRT(-1)')).toBe('#NUM!');
      expect(evaluateFormula('=SQRT(-4)')).toBe('#NUM!');
    });

    test('POWER function', () => {
      expect(evaluateFormula('=POWER(2, 3)')).toBe(8);
      expect(evaluateFormula('=POWER(5, 2)')).toBe(25);
      expect(evaluateFormula('=POWER(10, 0)')).toBe(1);
      expect(evaluateFormula('=POWER(4, 0.5)')).toBe(2);
      expect(evaluateFormula('=POWER(2, -2)')).toBe(0.25);
    });
  });

  describe('Exponential and Logarithmic Functions', () => {
    test('EXP function', () => {
      expect(evaluateFormula('=EXP(0)')).toBe(1);
      expect(evaluateFormula('=EXP(1)')).toBeCloseTo(Math.E, 5);
      expect(evaluateFormula('=EXP(2)')).toBeCloseTo(Math.E * Math.E, 5);
    });

    test('LN function', () => {
      expect(evaluateFormula('=LN(1)')).toBe(0);
      expect(evaluateFormula('=LN(' + Math.E + ')')).toBeCloseTo(1, 5);
      expect(evaluateFormula('=LN(10)')).toBeCloseTo(2.303, 3);
    });

    test('LN with invalid arguments returns error', () => {
      expect(evaluateFormula('=LN(0)')).toBe('#NUM!');
      expect(evaluateFormula('=LN(-1)')).toBe('#NUM!');
    });

    test('LOG function', () => {
      expect(evaluateFormula('=LOG(100)')).toBe(2); // Default base 10
      expect(evaluateFormula('=LOG(1000, 10)')).toBeCloseTo(3, 5);
      expect(evaluateFormula('=LOG(8, 2)')).toBeCloseTo(3, 5);
      expect(evaluateFormula('=LOG(1)')).toBe(0);
    });

    test('LOG with invalid arguments returns error', () => {
      expect(evaluateFormula('=LOG(0)')).toBe('#NUM!');
      expect(evaluateFormula('=LOG(-1)')).toBe('#NUM!');
      expect(evaluateFormula('=LOG(10, 0)')).toBe('#NUM!');
      expect(evaluateFormula('=LOG(10, 1)')).toBe('#NUM!');
      expect(evaluateFormula('=LOG(10, -1)')).toBe('#NUM!');
    });

    test('LOG10 function', () => {
      expect(evaluateFormula('=LOG10(10)')).toBe(1);
      expect(evaluateFormula('=LOG10(100)')).toBe(2);
      expect(evaluateFormula('=LOG10(1000)')).toBe(3);
      expect(evaluateFormula('=LOG10(1)')).toBe(0);
    });

    test('LOG10 with invalid arguments returns error', () => {
      expect(evaluateFormula('=LOG10(0)')).toBe('#NUM!');
      expect(evaluateFormula('=LOG10(-1)')).toBe('#NUM!');
    });
  });

  describe('Trigonometric Functions', () => {
    test('SIN function', () => {
      expect(evaluateFormula('=SIN(0)')).toBe(0);
      expect(evaluateFormula('=SIN(' + Math.PI / 2 + ')')).toBeCloseTo(1, 5);
      expect(evaluateFormula('=SIN(' + Math.PI + ')')).toBeCloseTo(0, 5);
    });

    test('COS function', () => {
      expect(evaluateFormula('=COS(0)')).toBe(1);
      expect(evaluateFormula('=COS(' + Math.PI / 2 + ')')).toBeCloseTo(0, 5);
      expect(evaluateFormula('=COS(' + Math.PI + ')')).toBeCloseTo(-1, 5);
    });

    test('TAN function', () => {
      expect(evaluateFormula('=TAN(0)')).toBe(0);
      expect(evaluateFormula('=TAN(' + Math.PI / 4 + ')')).toBeCloseTo(1, 5);
      expect(evaluateFormula('=TAN(' + Math.PI + ')')).toBeCloseTo(0, 5);
    });

    test('ASIN function', () => {
      expect(evaluateFormula('=ASIN(0)')).toBe(0);
      expect(evaluateFormula('=ASIN(1)')).toBeCloseTo(Math.PI / 2, 5);
      expect(evaluateFormula('=ASIN(-1)')).toBeCloseTo(-Math.PI / 2, 5);
      expect(evaluateFormula('=ASIN(0.5)')).toBeCloseTo(Math.PI / 6, 5);
    });

    test('ASIN with invalid arguments returns error', () => {
      expect(evaluateFormula('=ASIN(2)')).toBe('#NUM!');
      expect(evaluateFormula('=ASIN(-2)')).toBe('#NUM!');
    });

    test('ACOS function', () => {
      expect(evaluateFormula('=ACOS(1)')).toBe(0);
      expect(evaluateFormula('=ACOS(0)')).toBeCloseTo(Math.PI / 2, 5);
      expect(evaluateFormula('=ACOS(-1)')).toBeCloseTo(Math.PI, 5);
      expect(evaluateFormula('=ACOS(0.5)')).toBeCloseTo(Math.PI / 3, 5);
    });

    test('ACOS with invalid arguments returns error', () => {
      expect(evaluateFormula('=ACOS(2)')).toBe('#NUM!');
      expect(evaluateFormula('=ACOS(-2)')).toBe('#NUM!');
    });

    test('ATAN function', () => {
      expect(evaluateFormula('=ATAN(0)')).toBe(0);
      expect(evaluateFormula('=ATAN(1)')).toBeCloseTo(Math.PI / 4, 5);
      expect(evaluateFormula('=ATAN(-1)')).toBeCloseTo(-Math.PI / 4, 5);
    });

    test('ATAN2 function', () => {
      expect(evaluateFormula('=ATAN2(0, 1)')).toBe(0);
      expect(evaluateFormula('=ATAN2(1, 1)')).toBeCloseTo(Math.PI / 4, 5);
      expect(evaluateFormula('=ATAN2(1, 0)')).toBeCloseTo(Math.PI / 2, 5);
      expect(evaluateFormula('=ATAN2(0, -1)')).toBeCloseTo(Math.PI, 5);
    });
  });

  describe('Angle Conversion Functions', () => {
    test('DEGREES function', () => {
      expect(evaluateFormula('=DEGREES(0)')).toBe(0);
      expect(evaluateFormula('=DEGREES(' + Math.PI + ')')).toBeCloseTo(180, 5);
      expect(evaluateFormula('=DEGREES(' + Math.PI / 2 + ')')).toBeCloseTo(90, 5);
      expect(evaluateFormula('=DEGREES(' + 2 * Math.PI + ')')).toBeCloseTo(360, 5);
    });

    test('RADIANS function', () => {
      expect(evaluateFormula('=RADIANS(0)')).toBe(0);
      expect(evaluateFormula('=RADIANS(180)')).toBeCloseTo(Math.PI, 5);
      expect(evaluateFormula('=RADIANS(90)')).toBeCloseTo(Math.PI / 2, 5);
      expect(evaluateFormula('=RADIANS(360)')).toBeCloseTo(2 * Math.PI, 5);
    });

    test('PI function', () => {
      expect(evaluateFormula('=PI()')).toBe(Math.PI);
    });
  });

  describe('Rounding Functions', () => {
    test('ROUND function', () => {
      expect(evaluateFormula('=ROUND(2.15, 1)')).toBeCloseTo(2.2, 5);
      expect(evaluateFormula('=ROUND(2.149, 1)')).toBeCloseTo(2.1, 5);
      expect(evaluateFormula('=ROUND(-1.475, 2)')).toBeCloseTo(-1.47, 5); // JavaScript rounding behavior
      expect(evaluateFormula('=ROUND(21.5, -1)')).toBe(20);
      expect(evaluateFormula('=ROUND(626.3, -3)')).toBe(1000);
      expect(evaluateFormula('=ROUND(1.98, 0)')).toBe(2);
      expect(evaluateFormula('=ROUND(-50.55, 0)')).toBe(-51);
    });

    test('ROUNDUP function', () => {
      expect(evaluateFormula('=ROUNDUP(3.2, 0)')).toBe(4);
      expect(evaluateFormula('=ROUNDUP(76.9, 0)')).toBe(77);
      expect(evaluateFormula('=ROUNDUP(3.14159, 3)')).toBe(3.142);
      expect(evaluateFormula('=ROUNDUP(-3.14159, 1)')).toBe(-3.2);
      expect(evaluateFormula('=ROUNDUP(31415.92654, -2)')).toBe(31500);
    });

    test('ROUNDDOWN function', () => {
      expect(evaluateFormula('=ROUNDDOWN(3.2, 0)')).toBe(3);
      expect(evaluateFormula('=ROUNDDOWN(76.9, 0)')).toBe(76);
      expect(evaluateFormula('=ROUNDDOWN(3.14159, 3)')).toBe(3.141);
      expect(evaluateFormula('=ROUNDDOWN(-3.14159, 1)')).toBe(-3.1);
      expect(evaluateFormula('=ROUNDDOWN(31415.92654, -2)')).toBe(31400);
    });
  });

  describe('CEILING and FLOOR Functions', () => {
    test('CEILING function', () => {
      expect(evaluateFormula('=CEILING(2.5, 1)')).toBe(3);
      expect(evaluateFormula('=CEILING(-2.5, -2)')).toBe(-4);
      expect(evaluateFormula('=CEILING(-2.5, 2)')).toBe('#NUM!');
      expect(evaluateFormula('=CEILING(1.5, 0.1)')).toBeCloseTo(1.5, 5);
      expect(evaluateFormula('=CEILING(0.234, 0.01)')).toBeCloseTo(0.24, 5);
    });

    test('FLOOR function', () => {
      expect(evaluateFormula('=FLOOR(2.5, 1)')).toBe(2);
      expect(evaluateFormula('=FLOOR(-2.5, -2)')).toBe(-2);
      expect(evaluateFormula('=FLOOR(-2.5, 2)')).toBe('#NUM!');
      expect(evaluateFormula('=FLOOR(1.5, 0.1)')).toBeCloseTo(1.5, 5);
      expect(evaluateFormula('=FLOOR(0.234, 0.01)')).toBeCloseTo(0.23, 5);
    });
  });

  describe('Integer and Truncation Functions', () => {
    test('INT function', () => {
      expect(evaluateFormula('=INT(8.9)')).toBe(8);
      expect(evaluateFormula('=INT(-8.9)')).toBe(-9);
      expect(evaluateFormula('=INT(0.5)')).toBe(0);
      expect(evaluateFormula('=INT(-0.5)')).toBe(-1);
    });

    test('TRUNC function', () => {
      expect(evaluateFormula('=TRUNC(8.9)')).toBe(8);
      expect(evaluateFormula('=TRUNC(-8.9)')).toBe(-8);
      expect(evaluateFormula('=TRUNC(3.14159, 2)')).toBe(3.14);
      expect(evaluateFormula('=TRUNC(-3.14159, 1)')).toBe(-3.1);
      expect(evaluateFormula('=TRUNC(31415.92654, -2)')).toBe(31400);
    });

    test('MOD function', () => {
      expect(evaluateFormula('=MOD(3, 2)')).toBe(1);
      expect(evaluateFormula('=MOD(-3, 2)')).toBe(1);
      expect(evaluateFormula('=MOD(3, -2)')).toBe(-1);
      expect(evaluateFormula('=MOD(-3, -2)')).toBe(-1);
      expect(evaluateFormula('=MOD(10, 3)')).toBe(1);
    });

    test('MOD with zero divisor returns error', () => {
      expect(evaluateFormula('=MOD(10, 0)')).toBe('#DIV/0!');
    });
  });

  describe('Even and Odd Functions', () => {
    test('EVEN function', () => {
      expect(evaluateFormula('=EVEN(1.5)')).toBe(2);
      expect(evaluateFormula('=EVEN(3)')).toBe(4);
      expect(evaluateFormula('=EVEN(2)')).toBe(2);
      expect(evaluateFormula('=EVEN(-1)')).toBe(-2);
      expect(evaluateFormula('=EVEN(-2)')).toBe(-2);
      expect(evaluateFormula('=EVEN(0)')).toBe(0);
    });

    test('ODD function', () => {
      expect(evaluateFormula('=ODD(1.5)')).toBe(3);
      expect(evaluateFormula('=ODD(3)')).toBe(3);
      expect(evaluateFormula('=ODD(2)')).toBe(3);
      expect(evaluateFormula('=ODD(-1)')).toBe(-1);
      expect(evaluateFormula('=ODD(-2)')).toBe(-3);
      expect(evaluateFormula('=ODD(0)')).toBe(1);
    });
  });

  describe('Special Functions', () => {
    test('FACT function', () => {
      expect(evaluateFormula('=FACT(0)')).toBe(1);
      expect(evaluateFormula('=FACT(1)')).toBe(1);
      expect(evaluateFormula('=FACT(5)')).toBe(120);
      expect(evaluateFormula('=FACT(10)')).toBe(3628800);
    });

    test('FACT with invalid arguments returns error', () => {
      expect(evaluateFormula('=FACT(-1)')).toBe('#NUM!');
      expect(evaluateFormula('=FACT(1.5)')).toBe('#NUM!');
      expect(evaluateFormula('=FACT(171)')).toBe('#NUM!'); // Too large
    });

    test('DECIMAL function', () => {
      expect(evaluateFormula('=DECIMAL("FF", 16)')).toBe(255);
      expect(evaluateFormula('=DECIMAL("111", 2)')).toBe(7);
      expect(evaluateFormula('=DECIMAL("123", 8)')).toBe(83);
      expect(evaluateFormula('=DECIMAL("ZZ", 36)')).toBe(1295);
    });

    test('DECIMAL with invalid arguments returns error', () => {
      expect(evaluateFormula('=DECIMAL("FF", 10)')).toBe('#NUM!'); // Invalid for base 10
      expect(evaluateFormula('=DECIMAL("123", 1)')).toBe('#NUM!'); // Invalid base
      expect(evaluateFormula('=DECIMAL("123", 37)')).toBe('#NUM!'); // Invalid base
    });
  });

  describe('Type Coercion', () => {
    test('String to number coercion', () => {
      expect(evaluateFormula('=ABS("-5")')).toBe(5);
      expect(evaluateFormula('=SQRT("9")')).toBe(3);
      expect(evaluateFormula('=POWER("2", "3")')).toBe(8);
    });

    test('Boolean to number coercion', () => {
      expect(evaluateFormula('=ABS(TRUE)')).toBe(1);
      expect(evaluateFormula('=ABS(FALSE)')).toBe(0);
      expect(evaluateFormula('=SQRT(TRUE)')).toBe(1);
    });

    test('Invalid string returns error', () => {
      expect(evaluateFormula('=ABS("abc")')).toBe('#VALUE!');
      expect(evaluateFormula('=SQRT("xyz")')).toBe('#VALUE!');
    });
  });

  describe('Error Propagation', () => {
    test('Error in arguments propagates', () => {
      expect(evaluateFormula('=ABS(SQRT(-1))')).toBe('#NUM!');
      expect(evaluateFormula('=POWER(#VALUE!, 2)')).toBe('#VALUE!');
    });

    test('Wrong number of arguments', () => {
      expect(() => evaluateFormula('=ABS()')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=ABS(1, 2)')).toThrow('Invalid number of arguments');
      expect(() => evaluateFormula('=PI(1)')).toThrow('Invalid number of arguments');
    });
  });
});
