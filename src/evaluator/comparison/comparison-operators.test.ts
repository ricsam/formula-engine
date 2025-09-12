import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "src/core/engine";
import { type SerializedCellValue } from "src/core/types";
import { parseCellReference } from "src/core/utils";

describe("Comparison Operators", () => {
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

  describe("equals operator (=)", () => {
    test("should compare equal numbers", () => {
      setCellContent("A1", "=5=5");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=10=5");
      expect(cell("A2")).toBe(false);
    });

    test("should compare equal strings", () => {
      setCellContent("A1", '="Hello"="Hello"');
      expect(cell("A1")).toBe(true);

      setCellContent("A2", '="Hello"="World"');
      expect(cell("A2")).toBe(false);
    });

    test("should compare equal booleans", () => {
      setCellContent("A1", "=TRUE=TRUE");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=TRUE=FALSE");
      expect(cell("A2")).toBe(false);
    });

    test("should handle infinity comparisons", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"], // Positive infinity
          ["A2", "=1/0"], // Positive infinity
          ["A3", "=-1/0"], // Negative infinity
          ["B1", "=A1=A2"], // +∞ = +∞
          ["B2", "=A1=A3"], // +∞ = -∞
        ])
      );

      expect(cell("B1")).toBe(true);  // Same infinity
      expect(cell("B2")).toBe(false); // Different infinities
    });

    test("should return false for different types", () => {
      setCellContent("A1", '=5="5"');
      expect(cell("A1")).toBe(false);

      setCellContent("A2", '="TRUE"=TRUE');
      expect(cell("A2")).toBe(false);
    });

    test("should handle cell references", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 10],
          ["B1", 10],
          ["C1", 5],
          ["D1", "=A1=B1"], // 10 = 10
          ["D2", "=A1=C1"], // 10 = 5
        ])
      );

      expect(cell("D1")).toBe(true);
      expect(cell("D2")).toBe(false);
    });
  });

  describe("not equals operator (<>)", () => {
    test("should return opposite of equals", () => {
      setCellContent("A1", "=5<>5");
      expect(cell("A1")).toBe(false);

      setCellContent("A2", "=10<>5");
      expect(cell("A2")).toBe(true);
    });

    test("should work with different types", () => {
      setCellContent("A1", '=5<>"5"');
      expect(cell("A1")).toBe(true);
    });
  });

  describe("less than operator (<)", () => {
    test("should compare numbers", () => {
      setCellContent("A1", "=5<10");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=10<5");
      expect(cell("A2")).toBe(false);

      setCellContent("A3", "=5<5");
      expect(cell("A3")).toBe(false);
    });

    test("should handle infinity", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "=1/0"],  // +∞
          ["A2", "=-1/0"], // -∞
          ["B1", "=5<A1"], // 5 < +∞
          ["B2", "=5<A2"], // 5 < -∞
          ["B3", "=A2<A1"], // -∞ < +∞
        ])
      );

      expect(cell("B1")).toBe(true);  // 5 < +∞
      expect(cell("B2")).toBe(false); // 5 < -∞
      expect(cell("B3")).toBe(true);  // -∞ < +∞
    });

    test("should return error for non-numeric types", () => {
      setCellContent("A1", '="Hello"<"World"');
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=TRUE<FALSE");
      expect(cell("A2")).toBe("#VALUE!");
    });
  });

  describe("less than or equal operator (<=)", () => {
    test("should combine less than and equals", () => {
      setCellContent("A1", "=5<=10");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=5<=5");
      expect(cell("A2")).toBe(true);

      setCellContent("A3", "=10<=5");
      expect(cell("A3")).toBe(false);
    });
  });

  describe("greater than operator (>)", () => {
    test("should compare numbers", () => {
      setCellContent("A1", "=10>5");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=5>10");
      expect(cell("A2")).toBe(false);

      setCellContent("A3", "=5>5");
      expect(cell("A3")).toBe(false);
    });
  });

  describe("greater than or equal operator (>=)", () => {
    test("should combine greater than and equals", () => {
      setCellContent("A1", "=10>=5");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=5>=5");
      expect(cell("A2")).toBe(true);

      setCellContent("A3", "=5>=10");
      expect(cell("A3")).toBe(false);
    });
  });

  describe("concatenation operator (&)", () => {
    test("should concatenate strings", () => {
      setCellContent("A1", '="Hello"&"World"');
      expect(cell("A1")).toBe("HelloWorld");

      setCellContent("A2", '="Hello"&" "&"World"');
      expect(cell("A2")).toBe("Hello World");
    });

    test("should concatenate numbers", () => {
      setCellContent("A1", "=123&456");
      expect(cell("A1")).toBe("123456");

      setCellContent("A2", "=0&1");
      expect(cell("A2")).toBe("01");
    });

    test("should concatenate mixed numbers and strings", () => {
      setCellContent("A1", '="Value: "&42');
      expect(cell("A1")).toBe("Value: 42");

      setCellContent("A2", '=100&" percent"');
      expect(cell("A2")).toBe("100 percent");
    });

    test("should handle decimal numbers", () => {
      setCellContent("A1", '="Pi is "&3.14159');
      expect(cell("A1")).toBe("Pi is 3.14159");
    });

    test("should return error for unsupported types", () => {
      setCellContent("A1", '="Hello"&TRUE');
      expect(cell("A1")).toBe("#VALUE!");

      setCellContent("A2", "=FALSE&5");
      expect(cell("A2")).toBe("#VALUE!");
    });

    test("should work with cell references", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["B1", " "],
          ["C1", "World"],
          ["D1", "=A1&B1&C1"],
        ])
      );

      expect(cell("D1")).toBe("Hello World");
    });
  });

  describe("dynamic arrays with operators", () => {
    test("should handle spilled values in comparisons", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 10],
          ["A3", 15],
          ["B1", "=A1:A3>10"],
        ])
      );

      expect(cell("B1")).toBe(false); // 5 > 10
      expect(cell("B2")).toBe(false); // 10 > 10
      expect(cell("B3")).toBe(true);  // 15 > 10
    });

    test("should handle spilled values in concatenation", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", "Hello"],
          ["A2", "Good"],
          ["A3", "Nice"],
          ["B1", "=A1:A3&\" World\""],
        ])
      );

      expect(cell("B1")).toBe("Hello World");
      expect(cell("B2")).toBe("Good World");
      expect(cell("B3")).toBe("Nice World");
    });

    test("should handle multiple spilled arrays", () => {
      engine.setSheetContent(
        sheetAddress,
        new Map<string, SerializedCellValue>([
          ["A1", 5],
          ["A2", 10],
          ["B1", 10],
          ["B2", 5],
          ["C1", "=A1:A2=B1:B2"],
        ])
      );

      expect(cell("C1")).toBe(false); // 5 = 10
      expect(cell("C2")).toBe(false); // 10 = 5
    });
  });

  describe("edge cases", () => {
    test("should handle zero comparisons", () => {
      setCellContent("A1", "=0=0");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=0<1");
      expect(cell("A2")).toBe(true);

      setCellContent("A3", "=-1<0");
      expect(cell("A3")).toBe(true);
    });

    test("should handle negative numbers", () => {
      setCellContent("A1", "=-5<-3");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=-10>-20");
      expect(cell("A2")).toBe(true);
    });

    test("should handle decimal comparisons", () => {
      setCellContent("A1", "=3.14<3.15");
      expect(cell("A1")).toBe(true);

      setCellContent("A2", "=2.5>=2.5");
      expect(cell("A2")).toBe(true);
    });

    test("should handle empty string concatenation", () => {
      setCellContent("A1", '=""&"Hello"');
      expect(cell("A1")).toBe("Hello");

      setCellContent("A2", '="Hello"&""');
      expect(cell("A2")).toBe("Hello");
    });
  });
});
