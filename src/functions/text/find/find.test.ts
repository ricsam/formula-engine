import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("FIND function", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: string) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  describe("basic functionality", () => {
    test("should find text (case-sensitive)", () => {
      setCellContent("A1", '=FIND("World","Hello World")');
      expect(cell("A1")).toBe(7);
    });

    test("should be case-sensitive", () => {
      setCellContent("A1", '=FIND("world","Hello World")');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should support start position", () => {
      setCellContent("A1", '=FIND("l","Hello World",4)');
      expect(cell("A1")).toBe(4); // Second 'l' in "Hello"
    });

    test("should return error when not found", () => {
      setCellContent("A1", '=FIND("xyz","Hello World")');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should find at beginning of string", () => {
      setCellContent("A1", '=FIND("Hello","Hello World")');
      expect(cell("A1")).toBe(1);
    });

    test("should find single character", () => {
      setCellContent("A1", '=FIND("o","Hello")');
      expect(cell("A1")).toBe(5);
    });

    test("should find empty string", () => {
      setCellContent("A1", '=FIND("","Hello")');
      expect(cell("A1")).toBe(1);
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return error for number input", () => {
      setCellContent("A1", '=FIND("23",12345)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for boolean input", () => {
      setCellContent("A1", '=FIND("RU",TRUE)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity text input", () => {
      setCellContent("A1", '=FIND("INF",INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity text input", () => {
      setCellContent("A1", '=FIND("-",-INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity as start position", () => {
      setCellContent("A1", '=FIND("l","Hello",INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle very large start position", () => {
      setCellContent("A1", '=FIND("l","Hello",100)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("start position validation", () => {
    test("should return error for start position less than 1", () => {
      setCellContent("A1", '=FIND("l","Hello",0)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative start position", () => {
      setCellContent("A1", '=FIND("l","Hello",-1)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle decimal start position (should floor)", () => {
      setCellContent("A1", '=FIND("l","Hello",3.9)');
      expect(cell("A1")).toBe(3); // Floors to 3, finds 'l' at position 3
    });

    test("should handle start position equal to string length", () => {
      setCellContent("A1", '=FIND("o","Hello",5)');
      expect(cell("A1")).toBe(5);
    });

    test("should handle start position beyond string length", () => {
      setCellContent("A1", '=FIND("o","Hello",6)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled findText values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", ","],
          ["A2", "cat"],
          ["A3", "green"],
          ["B1", '=FIND(A1:A3,"apple,banana,cherry dog,cat,bird red,green,blue")'],
        ])
      );

      expect(cell("B1")).toBe(6); // Position of "," in the long string
      expect(cell("B2")).toBe(25); // Position of "cat"
      expect(cell("B3")).toBe(38); // Position of "green"
    });

    test("should handle spilled withinText values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", '=FIND(",",A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe(6); // Position of "," in "apple,banana"
      expect(cell("B2")).toBe(4); // Position of "," in "dog,cat,bird"
      expect(cell("B3")).toBe(4); // Position of "," in "red,green,blue"
    });

    test("should handle zipped spilled findText and withinText", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", ","],
          ["A2", "cat"],
          ["A3", "blue"],
          ["B1", "apple,banana,cherry"],
          ["B2", "dog,cat,bird"],
          ["B3", "red,green,blue"],
          ["C1", '=FIND(A1:A3,B1:B3)'],
        ])
      );

      expect(cell("C1")).toBe(6); // "," in "apple,banana,cherry"
      expect(cell("C2")).toBe(5); // "cat" in "dog,cat,bird"
      expect(cell("C3")).toBe(11); // "blue" in "red,green,blue"
    });

    test("should work with LEFT for comma extraction", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana,cherry"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", '=FIND(",",A1:A3)'],
          ["C1", '=LEFT(A1:A3,B1:B3-1)'],
        ])
      );

      expect(cell("B1")).toBe(6);
      expect(cell("B2")).toBe(4);
      expect(cell("B3")).toBe(4);
      expect(cell("C1")).toBe("apple");
      expect(cell("C2")).toBe("dog");
      expect(cell("C3")).toBe("red");
    });

    test("should handle spilled arrays with some text not found", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple"],
          ["A2", "xyz"], // This won't be found
          ["A3", "green"],
          ["B1", '=FIND(A1:A3,"apple,banana,green")'],
        ])
      );

      expect(cell("B1")).toBe(1); // "apple" found at position 1
      expect(cell("B2")).toBe("#VALUE!"); // "xyz" not found
      expect(cell("B3")).toBe(14); // "green" found at position 14
    });
  });

  describe("error handling", () => {
    test("should return error for invalid findText argument", () => {
      // Error propagation may transform #REF! to #ERROR! through the engine
      setCellContent("A1", '=FIND(#REF!,"Hello")');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for invalid withinText argument", () => {
      // Error propagation may transform #REF! to #ERROR! through the engine
      setCellContent("A1", '=FIND("l",#REF!)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for invalid startNum argument", () => {
      // Function validation rejects #REF! for startNum with #VALUE!
      setCellContent("A1", '=FIND("l","Hello",#REF!)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", '=FIND("Hello")');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for too many arguments", () => {
      setCellContent("A1", '=FIND("Hello","World",1,2)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle spilled startNum argument", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 3],
          ["A3", 7],
          ["B1", '=FIND("l","Hello World",A1:A3)'],
        ])
      );

      // Should create spilled results based on Excel behavior
      expect(cell("B1")).toBe(3); // First 'l' at position 3 (start=1)
      expect(cell("B2")).toBe(3); // First 'l' at position 3 (start=3)  
      expect(cell("B3")).toBe(10); // Second 'l' at position 10 (start=7)
    });

    test("should handle array syntax for startNum - Excel: FIND('o','Hello World',{1;6;7}) gives 5;8;8", () => {
      setCellContent("A1", '=FIND("o","Hello World",{1;6;7})');
      expect(cell("A1")).toBe(5);  // 'o' at position 5 (start=1)
      expect(cell("A2")).toBe(8);  // 'o' at position 8 (start=6) 
      expect(cell("A3")).toBe(8);  // 'o' at position 8 (start=7)
    });

    test("should handle spilled findText - Excel: FIND({'l';'o';'W'},'Hello World',{1;1;1}) gives 3;5;7", () => {
      setCellContent("A1", '=FIND({"l";"o";"W"},"Hello World",{1;1;1})');
      expect(cell("A1")).toBe(3);  // 'l' at position 3
      expect(cell("A2")).toBe(5);  // 'o' at position 5
      expect(cell("A3")).toBe(7);  // 'W' at position 7
    });

    test("should handle all spilled arguments - Excel: FIND({'l';'o';'W'},{'Hello World';'Hello World';'Hello World'},{1;6;7}) gives 3;8;7", () => {
      setCellContent("A1", '=FIND({"l";"o";"W"},{"Hello World";"Hello World";"Hello World"},{1;6;7})');
      expect(cell("A1")).toBe(3);  // 'l' in "Hello World" starting at 1 = position 3
      expect(cell("A2")).toBe(8);  // 'o' in "Hello World" starting at 6 = position 8  
      expect(cell("A3")).toBe(7);  // 'W' in "Hello World" starting at 7 = position 7
    });

    test("should handle errors in spilled arrays - Excel: FIND('l','Hello',{1;3;10;20}) gives 3;3;#VALUE!;#VALUE!", () => {
      setCellContent("A1", '=FIND("l","Hello",{1;3;10;20})');
      expect(cell("A1")).toBe(3);         // 'l' at position 3 (start=1)
      expect(cell("A2")).toBe(3);         // 'l' at position 3 (start=3)
      expect(cell("A3")).toBe("#VALUE!"); // start=10 beyond string length
      expect(cell("A4")).toBe("#VALUE!"); // start=20 beyond string length
    });

    test("should handle negative startNum in arrays - Excel: FIND('l','Hello',{1;-1;3}) gives 3;#VALUE!;3", () => {
      setCellContent("A1", '=FIND("l","Hello",{1;-1;3})');
      expect(cell("A1")).toBe(3);         // 'l' at position 3 (start=1)
      expect(cell("A2")).toBe("#VALUE!"); // start=-1 invalid
      expect(cell("A3")).toBe(3);         // 'l' at position 3 (start=3)
    });
  });

  describe("complex scenarios", () => {
    test("should find overlapping patterns", () => {
      setCellContent("A1", '=FIND("aa","aaaa")');
      expect(cell("A1")).toBe(1); // First occurrence
    });

    test("should find with start position after first occurrence", () => {
      setCellContent("A1", '=FIND("aa","aaaa",2)');
      expect(cell("A1")).toBe(2); // Second occurrence
    });

    test("should handle Unicode characters", () => {
      setCellContent("A1", '=FIND("😀","Hello 😀 World")');
      expect(cell("A1")).toBe(7);
    });

    test("should handle very long strings", () => {
      const longString = "A".repeat(500) + "FIND_ME" + "B".repeat(500);
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", longString],
          ["B1", '=FIND("FIND_ME",A1)'],
        ])
      );

      expect(cell("B1")).toBe(501);
    });
  });
});