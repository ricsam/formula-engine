import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("ISERROR function", () => {
  const sheetName = "TestSheet";
  const workbookName = "TestWorkbook";
  const sheetAddress = { workbookName, sheetName };
  let engine: FormulaEngine;

  const cell = (ref: string) =>
    engine.getCellValue({ sheetName, workbookName, ...parseCellReference(ref) });

  const setCellContent = (ref: string, content: SerializedCellValue) => {
    engine.setCellContent(
      { sheetName, workbookName, ...parseCellReference(ref) },
      content
    );
  };

  const setContent = (content: Array<[string, SerializedCellValue]>) => {
    engine.setSheetContent(
      sheetAddress,
      new Map<string, SerializedCellValue>(content)
    );
  };

  beforeEach(() => {
    engine = FormulaEngine.buildEmpty();
    engine.addWorkbook({ workbookName });
    engine.addSheet({ workbookName, sheetName });
  });

  describe("single values", () => {
    test("is TRUE for every kind of error", () => {
      setContent([
        ["A1", "=ISERROR(NOSUCHFN())"], // #NAME?
        ["A2", "=ISERROR(UnknownName)"], // #NAME?
        ["A3", '=ISERROR(1+"x")'], // #VALUE!
        ["A4", '=ISERROR(MATCH("z", {"x","y"}, 0))'], // #N/A
        ["A5", "=ISERROR(CEILING(-7.3, 1))"], // #NUM!
        ["A6", "=ISERROR(0/0)"], // #NUM!
        ["A7", "=ISERROR(MOD(5, 0))"], // #DIV/0!
        ["A8", "=ISERROR(#REF!)"], // error literal
        ["A9", "=ISERROR(#N/A)"],
        ["A10", '=ISERROR(FILTER({1;2}, {1;2}=3))'], // #CALC!
      ]);

      for (let row = 1; row <= 10; row++) {
        expect(cell(`A${row}`)).toBe(true);
      }
    });

    test("is TRUE for cycles, blocked spills and missing sheets", () => {
      setContent([
        ["B1", "=C1"],
        ["C1", "=B1"],
        ["B3", "=SEQUENCE(2)"],
        ["B4", "blocks the spill"],
        ["A1", "=ISERROR(B1)"],
        ["A3", "=ISERROR(B3)"],
        ["A5", "=ISERROR(Missing!A1)"],
      ]);

      expect(cell("B1")).toBe(FormulaError.CYCLE);
      expect(cell("A1")).toBe(true);
      expect(cell("B3")).toBe(FormulaError.SPILL);
      expect(cell("A3")).toBe(true);
      expect(cell("A5")).toBe(true);

      setCellContent("B4", undefined);
      expect(cell("A3")).toBe(false);
    });

    test("is FALSE for values", () => {
      setContent([
        // B1 is empty
        ["B2", 0],
        ["A1", "=ISERROR(B1)"],
        ["A2", "=ISERROR(B2)"],
        ["A3", "=ISERROR(42)"],
        ["A4", '=ISERROR("text")'],
        ["A5", '=ISERROR("")'],
        ["A6", "=ISERROR(TRUE)"],
        ["A7", "=ISERROR(FALSE)"],
        ["A8", '=ISERROR(IFERROR(1+"x", 0))'],
      ]);

      for (let row = 1; row <= 8; row++) {
        expect(cell(`A${row}`)).toBe(false);
      }
    });

    test("is FALSE for INFINITY, this engine's result of a division by zero", () => {
      setContent([
        ["A1", "=ISERROR(1/0)"],
        ["A2", "=ISERROR(INFINITY)"],
        ["A3", "=ISERROR(-1/0)"],
      ]);

      expect(cell("A1")).toBe(false);
      expect(cell("A2")).toBe(false);
      expect(cell("A3")).toBe(false);
    });

    test("tells text that reads like an error from the error", () => {
      setContent([
        ["B1", "#N/A"],
        ["B2", "=#N/A"],
        ["A1", "=ISERROR(B1)"],
        ["A2", "=ISERROR(B2)"],
      ]);

      // both read back as "#N/A"; only the formula is an error
      expect(cell("B1")).toBe("#N/A");
      expect(cell("B2")).toBe(FormulaError.NA);
      expect(cell("A1")).toBe(false);
      expect(cell("A2")).toBe(true);
    });

    test("tests referenced cells and recalculates", () => {
      setContent([
        ["B1", 4],
        ["B2", "=10/B1"],
        ["A1", "=ISERROR(B2)"],
      ]);

      expect(cell("A1")).toBe(false);

      setCellContent("B1", "four");
      expect(cell("B2")).toBe(FormulaError.VALUE);
      expect(cell("A1")).toBe(true);

      setCellContent("B1", 5);
      expect(cell("A1")).toBe(false);
    });

    test("is TRUE exactly where IFERROR returns its fallback", () => {
      const expressions = [
        "NOSUCHFN()",
        '1+"x"',
        "1/0",
        "0/0",
        '"text"',
        "B1",
        "MOD(7, 0)",
        "42",
      ];
      setContent(
        expressions.flatMap((expression, index): Array<[string, string]> => [
          [`C${index + 1}`, `=IF(ISERROR(${expression}), "fallback", ${expression})`],
          [`D${index + 1}`, `=IFERROR(${expression}, "fallback")`],
        ])
      );

      const column = (letter: string) =>
        expressions.map((_, index) => cell(`${letter}${index + 1}`));
      expect(column("C")).toEqual(column("D"));
      expect(cell("C1")).toBe("fallback");
      expect(cell("C3")).toBe("INFINITY");
    });

    test("guards a lookup", () => {
      setContent([
        ["B1", "apple"],
        ["B2", "pear"],
        ["A1", '=IF(ISERROR(MATCH("pear", B1:B2, 0)), "missing", "found")'],
        ["A2", '=IF(ISERROR(MATCH("plum", B1:B2, 0)), "missing", "found")'],
      ]);

      expect(cell("A1")).toBe("found");
      expect(cell("A2")).toBe("missing");
    });

    test("requires exactly one argument", () => {
      setContent([
        ["A1", "=ISERROR()"],
        ["A2", "=ISERROR(1, 2)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
    });
  });

  describe("arrays", () => {
    test("tests a range element by element and spills", () => {
      setContent([
        ["A1", 1],
        ["A2", "two"],
        // A3 is empty
        ["A4", '=1+"x"'],
        ["A5", "=NOSUCHFN()"],
        ["A6", "=1/0"],
        ["C1", "=ISERROR(A1:A6)"],
      ]);

      expect(["C1", "C2", "C3", "C4", "C5", "C6"].map(cell)).toEqual([
        false,
        false,
        false,
        true,
        true,
        false,
      ]);

      setCellContent("A3", "=A2*2");
      expect(cell("C3")).toBe(true);
      setCellContent("A4", 4);
      expect(cell("C4")).toBe(false);
    });

    test("tests computed arrays and array literals", () => {
      setContent([
        ["A1", '=ISERROR({1,"a";#N/A,2})'],
        ["A3", '=ISERROR(SEQUENCE(3)/{1;0;"x"})'],
      ]);

      expect(["A1", "B1", "A2", "B2"].map(cell)).toEqual([
        false,
        false,
        true,
        false,
      ]);
      // 1, INFINITY and #VALUE!
      expect(["A3", "A4", "A5"].map(cell)).toEqual([false, false, true]);
    });

    test("counts the errors of a range in an aggregate", () => {
      setContent([
        ["A1", 1],
        ["A2", "=NOSUCHFN()"],
        ["A4", '=1+"x"'],
        ["A5", "x"],
        ["B1", "=SUM(IF(ISERROR(A1:A6), 1, 0))"],
        // every cell of the bounded result, the empty ones included
        ["B2", "=COUNTA(ISERROR(A1:A6))"],
        ["B3", "=SUM(IF(ISERROR(A1:A6), 0, 1))"],
      ]);

      expect(cell("B1")).toBe(2);
      expect(cell("B2")).toBe(6);
      expect(cell("B3")).toBe(4);
    });

    test("reads an open-ended range by its populated cells", () => {
      setContent([
        ["A1", 1],
        ["A3", "=NOSUCHFN()"],
        ["A4", "=MOD(1, 0)"],
        ["B1", "=SUM(IF(ISERROR(A:A), 1, 0))"],
      ]);

      expect(cell("B1")).toBe(2);
    });

    test("filters out the rows whose value is an error", () => {
      setContent([
        ["A1", 4],
        ["A2", "=NOSUCHFN()"],
        ["A3", 9],
        ["C1", "=FILTER(A1:A3, IF(ISERROR(A1:A3), FALSE, TRUE))"],
      ]);

      expect(["C1", "C2"].map(cell)).toEqual([4, 9]);
    });
  });
});
