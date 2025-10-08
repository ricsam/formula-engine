import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("MID function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, workbookName, ...parseCellReference(ref) }, content);
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook(workbookName);
    engine.addSheet({ workbookName, sheetName });
  });

  describe("basic functionality", () => {
    test("MID should return middle characters", () => {
      setCellContent("A1", '=MID("Hello World",7,5)');
      expect(cell("A1")).toBe("World");
    });

    test("MID should handle start position 1", () => {
      setCellContent("A1", '=MID("Hello",1,3)');
      expect(cell("A1")).toBe("Hel");
    });

    test("should handle requests starting beyond string length", () => {
      setCellContent("A1", '=MID("Hi",10,5)');
      expect(cell("A1")).toBe("");
    });

    test("should handle requests longer than remaining string", () => {
      setCellContent("A1", '=MID("Hello",3,10)');
      expect(cell("A1")).toBe("llo");
    });

    test("should return error for start position less than 1", () => {
      setCellContent("A1", '=MID("Hello",0,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative start position", () => {
      setCellContent("A1", '=MID("Hello",-1,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative num_chars", () => {
      setCellContent("A1", '=MID("Hello",2,-1)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should handle empty string", () => {
      setCellContent("A1", '=MID("",1,5)');
      expect(cell("A1")).toBe("");
    });

    test("should handle zero characters", () => {
      setCellContent("A1", '=MID("Hello",3,0)');
      expect(cell("A1")).toBe("");
    });

    test("should handle decimal start position (should floor)", () => {
      setCellContent("A1", '=MID("Hello",2.9,3)');
      expect(cell("A1")).toBe("ell");
    });

    test("should handle decimal num_chars (should floor)", () => {
      setCellContent("A1", '=MID("Hello",2,3.9)');
      expect(cell("A1")).toBe("ell");
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return error for number text input", () => {
      setCellContent("A1", '=MID(12345,2,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for boolean text input", () => {
      setCellContent("A1", '=MID(TRUE,1,2)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for FALSE text input", () => {
      setCellContent("A1", '=MID(FALSE,1,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for non-number start_num", () => {
      setCellContent("A1", '=MID("Hello",TRUE,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for non-number num_chars", () => {
      setCellContent("A1", '=MID("Hello",2,TRUE)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("strict number-only handling", () => {
    test("should return error for infinity start_num", () => {
      setCellContent("A1", '=MID("Hello",INFINITY,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity start_num", () => {
      setCellContent("A1", '=MID("Hello",-INFINITY,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity num_chars", () => {
      setCellContent("A1", '=MID("Hello",2,INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity num_chars", () => {
      setCellContent("A1", '=MID("Hello",2,-INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity text input", () => {
      setCellContent("A1", '=MID(INFINITY,2,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for negative infinity text input", () => {
      setCellContent("A1", '=MID(-INFINITY,2,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("dynamic arrays (spilled values)", () => {
    test("should handle spilled text values with single start_num and num_chars", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana,cherry"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", '=MID(A1:A3,1,3)'],
        ])
      );

      expect(cell("B1")).toBe("app");
      expect(cell("B2")).toBe("dog");
      expect(cell("B3")).toBe("red");
    });

    test("should handle single text with spilled start_num values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 3],
          ["A3", 5],
          ["B1", '=MID("Hello World",A1:A3,3)'],
        ])
      );

      expect(cell("B1")).toBe("Hel");
      expect(cell("B2")).toBe("llo");
      expect(cell("B3")).toBe("o W");
    });

    test("should handle single text with spilled num_chars values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 1],
          ["A2", 3],
          ["A3", 5],
          ["B1", '=MID("Hello World",2,A1:A3)'],
        ])
      );

      expect(cell("B1")).toBe("e");
      expect(cell("B2")).toBe("ell");
      expect(cell("B3")).toBe("ello ");
    });

    test("should handle zipped spilled text, start_num, and num_chars values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "apple,banana,cherry"],
          ["A2", "dog,cat,bird"],
          ["A3", "red,green,blue"],
          ["B1", 2],
          ["B2", 1],
          ["B3", 3],
          ["C1", 3],
          ["C2", 3],
          ["C3", 2],
          ["D1", '=MID(A1:A3,B1:B3,C1:C3)'],
        ])
      );

      expect(cell("D1")).toBe("ppl");
      expect(cell("D2")).toBe("dog");
      expect(cell("D3")).toBe("d,");
    });

    test("should work with mixed spilled and single values", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "World"],
          ["A3", "Test"],
          ["B1", 2],
          ["B2", 3],
          ["B3", 1],
          ["C1", '=MID(A1:A3,B1:B3,2)'],
        ])
      );

      expect(cell("C1")).toBe("el");
      expect(cell("C2")).toBe("rl");
      expect(cell("C3")).toBe("Te");
    });
  });

  describe("error handling", () => {
    test("should return error for invalid text argument", () => {
      setCellContent("A1", '=MID(#REF!,2,3)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for invalid start_num argument", () => {
      setCellContent("A1", '=MID("Hello",#REF!,3)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for invalid num_chars argument", () => {
      setCellContent("A1", '=MID("Hello",2,#REF!)');
      expect(cell("A1")).toBe("#ERROR!");
    });

    test("should return error for wrong number of arguments (too few)", () => {
      setCellContent("A1", '=MID("Hello",2)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for wrong number of arguments (too many)", () => {
      setCellContent("A1", '=MID("Hello",2,3,4)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for no arguments", () => {
      setCellContent("A1", '=MID()');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("edge cases", () => {
    test("should handle very long strings", () => {
      const longString = "A".repeat(1000);
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", longString],
          ["B1", '=MID(A1,500,100)'],
        ])
      );

      expect(cell("B1")).toBe("A".repeat(100));
    });

    test("should handle very large start position", () => {
      setCellContent("A1", '=MID("Hello",999999,3)');
      expect(cell("A1")).toBe("");
    });

    test("should handle very large num_chars", () => {
      setCellContent("A1", '=MID("Hello",2,999999)');
      expect(cell("A1")).toBe("ello");
    });

    test("should handle single character extraction", () => {
      setCellContent("A1", '=MID("Hello",3,1)');
      expect(cell("A1")).toBe("l");
    });

    test("should handle extraction from end of string", () => {
      setCellContent("A1", '=MID("Hello",5,1)');
      expect(cell("A1")).toBe("o");
    });

    test("should handle extraction exactly at string boundary", () => {
      setCellContent("A1", '=MID("Hello",6,1)');
      expect(cell("A1")).toBe("");
    });
  });
});
