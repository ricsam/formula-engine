import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("RIGHT function", () => {
  const sheetName = "TestSheet";
  let engine: FormulaEngine;

  const cell = (ref: string, debug?: boolean) =>
    engine.getCellValue({ sheetName, ...parseCellReference(ref) }, debug);

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent({ sheetName, ...parseCellReference(ref) }, content);
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addSheet(sheetName);
  });

  describe("basic functionality", () => {
    test("RIGHT should return rightmost characters", () => {
      setCellContent("A1", '=RIGHT("Hello World",5)');
      expect(cell("A1")).toBe("World");
    });

    test("RIGHT should default to 1 character", () => {
      setCellContent("A1", '=RIGHT("Hello")');
      expect(cell("A1")).toBe("o");
    });

    test("should handle requests longer than string", () => {
      setCellContent("A1", '=RIGHT("Hi",10)');
      expect(cell("A1")).toBe("Hi");
    });

    test("should return error for negative numbers", () => {
      setCellContent("A1", '=RIGHT("Hello",-1)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("strict type checking (no coercion)", () => {
    test("should return error for infinity numChars", () => {
      setCellContent("A1", '=RIGHT("Hello",INFINITY)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for number input", () => {
      setCellContent("A1", '=RIGHT(12345,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for infinity text input", () => {
      setCellContent("A1", '=RIGHT(INFINITY,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });

  describe("dynamic arrays", () => {
    test("should handle spilled text values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple"],
          ["A2", "banana"],
          ["A3", "cherry"],
          ["B1", '=RIGHT(A1:A3,2)'],
        ])
      );

      expect(cell("B1")).toBe("le");
      expect(cell("B2")).toBe("na");
      expect(cell("B3")).toBe("ry");
    });

    test("should handle zipped spilled values", () => {
      engine.setSheetContent(
        sheetName,
        new Map<string, SerializedCellValue>([
          ["A1", "apple"],
          ["A2", "banana"],
          ["A3", "cherry"],
          ["B1", 2],
          ["B2", 3],
          ["B3", 4],
          ["C1", '=RIGHT(A1:A3,B1:B3)'],
        ])
      );

      expect(cell("C1")).toBe("le");
      expect(cell("C2")).toBe("ana");
      expect(cell("C3")).toBe("erry");
    });
  });

  describe("error handling", () => {
    test("should return error for wrong number of arguments", () => {
      setCellContent("A1", '=RIGHT()');
      expect(cell("A1")).toBe("#VALUE!");
    });

    test("should return error for too many arguments", () => {
      setCellContent("A1", '=RIGHT("Hello",2,3)');
      expect(cell("A1")).toBe("#VALUE!");
    });
  });
});