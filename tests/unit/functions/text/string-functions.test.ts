import { test, expect, describe, beforeEach } from "bun:test";
import { FormulaEngine } from '../../../../src/core/engine';

describe('Text Functions Tests', () => {
  let engine: FormulaEngine;
  let sheetId: number;

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    const sheetName = engine.addSheet('TestSheet');
    sheetId = engine.getSheetId(sheetName);
  });

  describe('CONCATENATE function', () => {
    test('should concatenate multiple strings', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=CONCATENATE("Hello"," ","World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello World');
    });

    test('should concatenate numbers and strings', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=CONCATENATE("Number: ",123)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Number: 123');
    });

    test('should handle single argument', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=CONCATENATE("Hello")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello');
    });

    test('should skip undefined/null values', () => {
      const testData = new Map([['A1', 'Hello'], ['A3', 'World']]);
      engine.setSheetContents(sheetId, testData);
      
      const address = { sheet: sheetId, col: 1, row: 0 };
      engine.setCellContents(address, '=CONCATENATE(A1," ",A2," ",A3)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello  World');
    });
  });

  describe('LEN function', () => {
    test('should return length of string', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=LEN("Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(11);
    });

    test('should return length of number as string', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=LEN(12345)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(5);
    });

    test('should return 0 for empty/null values', () => {
      const address = { sheet: sheetId, col: 1, row: 0 }; // B1 instead of A1
      engine.setCellContents(address, '=LEN(A1)'); // A1 is empty
      
      const result = engine.getCellValue(address);
      expect(result).toBe(0);
    });

    test('should return 0 for empty string', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=LEN("")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(0);
    });
  });

  describe('UPPER and LOWER functions', () => {
    test('UPPER should convert to uppercase', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=UPPER("Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('HELLO WORLD');
    });

    test('LOWER should convert to lowercase', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=LOWER("Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('hello world');
    });

    test('should handle numbers', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      engine.setCellContents(address1, '=UPPER(123)');
      engine.setCellContents(address2, '=LOWER(456)');
      
      expect(engine.getCellValue(address1)).toBe('123');
      expect(engine.getCellValue(address2)).toBe('456');
    });

    test('should handle empty values', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      engine.setCellContents(address1, '=UPPER(A10)'); // Empty cell
      engine.setCellContents(address2, '=LOWER(A10)'); // Empty cell
      
      expect(engine.getCellValue(address1)).toBe('');
      expect(engine.getCellValue(address2)).toBe('');
    });
  });

  describe('TRIM function', () => {
    test('should remove leading and trailing spaces', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=TRIM("  Hello World  ")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello World');
    });

    test('should replace multiple spaces with single space', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=TRIM("Hello    World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello World');
    });

    test('should handle empty values', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=TRIM(A10)'); // Empty cell
      
      const result = engine.getCellValue(address);
      expect(result).toBe('');
    });
  });

  describe('LEFT and RIGHT functions', () => {
    test('LEFT should return leftmost characters', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=LEFT("Hello World",5)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello');
    });

    test('LEFT should default to 1 character', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=LEFT("Hello")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('H');
    });

    test('RIGHT should return rightmost characters', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=RIGHT("Hello World",5)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('World');
    });

    test('RIGHT should default to 1 character', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=RIGHT("Hello")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('o');
    });

    test('should handle requests longer than string', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      engine.setCellContents(address1, '=LEFT("Hi",10)');
      engine.setCellContents(address2, '=RIGHT("Hi",10)');
      
      expect(engine.getCellValue(address1)).toBe('Hi');
      expect(engine.getCellValue(address2)).toBe('Hi');
    });

    test('should return error for negative numbers', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      engine.setCellContents(address1, '=LEFT("Hello",-1)');
      engine.setCellContents(address2, '=RIGHT("Hello",-1)');
      
      expect(engine.getCellValue(address1)).toBe('#VALUE!');
      expect(engine.getCellValue(address2)).toBe('#VALUE!');
    });
  });

  describe('MID function', () => {
    test('should return characters from middle', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=MID("Hello World",7,5)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('World');
    });

    test('should handle start position beyond string length', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=MID("Hello",10,5)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('');
    });

    test('should handle length longer than remaining characters', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=MID("Hello",3,10)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('llo');
    });

    test('should return error for invalid parameters', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      const address3 = { sheet: sheetId, col: 0, row: 2 };
      
      engine.setCellContents(address1, '=MID("Hello",0,3)'); // Start < 1
      engine.setCellContents(address2, '=MID("Hello",2,-1)'); // Length < 0
      engine.setCellContents(address3, '=MID("Hello","a",3)'); // Non-numeric start
      
      expect(engine.getCellValue(address1)).toBe('#VALUE!');
      expect(engine.getCellValue(address2)).toBe('#VALUE!');
      expect(engine.getCellValue(address3)).toBe('#VALUE!');
    });
  });

  describe('FIND function', () => {
    test('should find text (case-sensitive)', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=FIND("World","Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(7);
    });

    test('should be case-sensitive', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=FIND("world","Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('#VALUE!');
    });

    test('should support start position', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=FIND("l","Hello World",4)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(4); // Second 'l' in "Hello"
    });

    test('should return error when not found', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=FIND("xyz","Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('#VALUE!');
    });
  });

  describe('SEARCH function', () => {
    test('should find text (case-insensitive)', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SEARCH("world","Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(7);
    });

    test('should support wildcards', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      
      engine.setCellContents(address1, '=SEARCH("W*d","Hello World")');
      engine.setCellContents(address2, '=SEARCH("W?rld","Hello World")');
      
      expect(engine.getCellValue(address1)).toBe(7);
      expect(engine.getCellValue(address2)).toBe(7);
    });

    test('should support start position', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SEARCH("l","Hello World",4)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe(4); // Second 'l' in "Hello"
    });

    test('should return error when not found', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SEARCH("xyz","Hello World")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('#VALUE!');
    });
  });

  describe('SUBSTITUTE function', () => {
    test('should substitute all instances by default', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SUBSTITUTE("Hello World Hello","Hello","Hi")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hi World Hi');
    });

    test('should substitute specific instance', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SUBSTITUTE("Hello World Hello","Hello","Hi",2)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello World Hi');
    });

    test('should return original text when old text not found', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SUBSTITUTE("Hello World","xyz","abc")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello World');
    });

    test('should handle empty old text', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=SUBSTITUTE("Hello","","X")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello');
    });
  });

  describe('REPLACE function', () => {
    test('should replace characters at specific position', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=REPLACE("Hello World",7,5,"Universe")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('Hello Universe');
    });

    test('should handle replacement beyond string length', () => {
      const address = { sheet: sheetId, col: 0, row: 0 };
      engine.setCellContents(address, '=REPLACE("Hello",10,3,"X")');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('HelloX');
    });

    test('should return error for invalid parameters', () => {
      const address1 = { sheet: sheetId, col: 0, row: 0 };
      const address2 = { sheet: sheetId, col: 0, row: 1 };
      
      engine.setCellContents(address1, '=REPLACE("Hello",0,3,"X")'); // Start < 1
      engine.setCellContents(address2, '=REPLACE("Hello",2,-1,"X")'); // Length < 0
      
      expect(engine.getCellValue(address1)).toBe('#VALUE!');
      expect(engine.getCellValue(address2)).toBe('#VALUE!');
    });
  });

  describe('Integration with cell references', () => {
    test('should work with cell references', () => {
      const testData = new Map([
        ['A1', 'Hello'],
        ['A2', 'World'],
        ['A3', 'John Doe'],
        ['A4', '  Extra Spaces  ']
      ]);
      engine.setSheetContents(sheetId, testData);
      
      const tests = [
        { address: { sheet: sheetId, col: 1, row: 0 }, formula: '=CONCATENATE(A1," ",A2)', expected: 'Hello World' },
        { address: { sheet: sheetId, col: 1, row: 1 }, formula: '=LEN(A3)', expected: 8 },
        { address: { sheet: sheetId, col: 1, row: 2 }, formula: '=UPPER(A1)', expected: 'HELLO' },
        { address: { sheet: sheetId, col: 1, row: 3 }, formula: '=TRIM(A4)', expected: 'Extra Spaces' },
        { address: { sheet: sheetId, col: 1, row: 4 }, formula: '=LEFT(A3,4)', expected: 'John' },
        { address: { sheet: sheetId, col: 1, row: 5 }, formula: '=RIGHT(A3,3)', expected: 'Doe' },
        { address: { sheet: sheetId, col: 1, row: 6 }, formula: '=MID(A3,6,3)', expected: 'Doe' }
      ];
      
      tests.forEach(({ address, formula, expected }) => {
        engine.setCellContents(address, formula);
        expect(engine.getCellValue(address)).toBe(expected);
      });
    });
  });

  describe('Error handling', () => {
    test('should propagate errors in arguments', () => {
      const testData = new Map([['A1', '#REF!']]);
      engine.setSheetContents(sheetId, testData);
      
      const address = { sheet: sheetId, col: 1, row: 0 };
      engine.setCellContents(address, '=LEN(A1)');
      
      const result = engine.getCellValue(address);
      expect(result).toBe('#REF!');
    });

    test('should handle invalid argument counts', () => {
      const tests = [
        '=LEN()',           // Too few args
        '=LEN("a","b")',    // Too many args
        '=MID("test",1)',   // Too few args
        '=MID("test",1,1,1)' // Too many args
      ];
      
      tests.forEach((formula, index) => {
        const address = { sheet: sheetId, col: 0, row: index };
        engine.setCellContents(address, formula);
        const result = engine.getCellValue(address);
        expect(typeof result === 'string' && result.startsWith('#')).toBe(true);
      });
    });
  });
});