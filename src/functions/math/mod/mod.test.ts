import { beforeEach, describe, expect, test } from "bun:test";
import { FormulaEngine } from "../../../core/engine";
import { FormulaError, type SerializedCellValue } from "../../../core/types";
import { parseCellReference } from "../../../core/utils";

describe("MOD function", () => {
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
    test("returns the remainder with the sign of the divisor", () => {
      setContent([
        // Excel's documented examples
        ["A1", "=MOD(3, 2)"],
        ["A2", "=MOD(-3, 2)"],
        ["A3", "=MOD(3, -2)"],
        ["A4", "=MOD(-3, -2)"],
        ["A5", "=MOD(10, 3)"],
        ["A6", "=MOD(12, 3)"],
        ["A7", "=MOD(100, 33)"],
        ["A8", "=MOD(7, -3)"],
        ["A9", "=MOD(0, 5)"],
        ["A10", "=MOD(6.25, 1)"],
        ["A11", "=MOD(5.5, 2)"],
        ["A12", "=MOD(-5.5, 2)"],
        ["A13", "=MOD(5, 7)"],
        ["A14", "=MOD(-5, 7)"],
      ]);

      expect(cell("A1")).toBe(1);
      expect(cell("A2")).toBe(1);
      expect(cell("A3")).toBe(-1);
      expect(cell("A4")).toBe(-1);
      expect(cell("A5")).toBe(1);
      expect(cell("A6")).toBe(0);
      expect(cell("A7")).toBe(1);
      expect(cell("A8")).toBe(-2);
      expect(cell("A9")).toBe(0);
      expect(cell("A10")).toBe(0.25);
      expect(cell("A11")).toBe(1.5);
      expect(cell("A12")).toBe(0.5);
      expect(cell("A13")).toBe(5);
      expect(cell("A14")).toBe(2);
    });

    test("never returns negative zero", () => {
      setContent([
        ["A1", "=MOD(-6, 3)"],
        ["A2", "=MOD(6, -3)"],
        ["A3", "=MOD(-0, 3)"],
        ["A4", "=MOD(0, -3)"],
        ["A5", "=MOD(-1.5, 0.5)"],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5"]) {
        expect(Object.is(cell(ref), 0)).toBe(true);
      }
    });

    test("is 0 for a multiple of the divisor in the engine's arithmetic", () => {
      setContent([
        // the binary 0.5 is slightly below five times the binary 0.1, but
        // 0.1*5 is 0.5
        ["A1", "=MOD(0.5, 0.1)"],
        ["A2", "=MOD(1, 0.1)"],
        ["A3", "=MOD(1.1, 0.1)"],
        ["A4", "=MOD(100, 0.01)"],
        ["A5", "=MOD(1, 0.2)"],
        ["A6", "=MOD(-0.5, 0.1)"],
        ["A7", "=MOD(0.5, -0.1)"],
        ["A8", "=MOD(12.3, 0.1)"],
        ["B1", "=0.1*5=0.5"],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8"]) {
        expect(cell(ref)).toBe(0);
      }
      expect(cell("B1")).toBe(true);
    });

    test("otherwise returns the remainder of the stored binary values", () => {
      setContent([
        // 0.1*3 is 0.30000000000000004, so 0.3 is not a multiple of 0.1
        ["A1", "=MOD(0.3, 0.1)"],
        ["B1", "=0.1*3=0.3"],
        // the fractional part of 5.1 is exactly 5.1-5
        ["A2", "=MOD(5.1, 1)"],
        ["B2", "=5.1-5"],
        ["A3", "=MOD(10, 3.3)"],
        // Excel returns 0.999999999992724 for this too: the product is one
        // unit of the last digit below 62142
        ["A4", "=MOD(621.42*100, 1)"],
        ["A5", "=MOD(-0.8, 1)"],
        ["B5", "=1-0.8"],
      ]);

      expect(cell("A1")).toBe(0.09999999999999998);
      expect(cell("B1")).toBe(false);
      expect([cell("A2"), cell("B2")]).toEqual([
        0.09999999999999964, 0.09999999999999964,
      ]);
      expect(cell("A3")).toBe(0.10000000000000053);
      expect(cell("A4")).toBe(0.999999999992724);
      expect([cell("A5"), cell("B5")]).toEqual([
        0.19999999999999996, 0.19999999999999996,
      ]);
    });

    test("extracts the time of a date-time and spans midnight", () => {
      setContent([
        // 2025-01-21 18:00
        ["B1", 45678.75],
        // 21:00 to 03:00 the next day
        ["B2", 0.875],
        ["B3", 0.125],
        ["A1", "=MOD(B1, 1)"],
        ["A2", "=MOD(B3-B2, 1)"],
        ["A3", "=MOD(B3-B2, 1)*24"],
      ]);

      expect(cell("A1")).toBe(0.75);
      expect(cell("A2")).toBe(0.25);
      expect(cell("A3")).toBe(6);
    });

    test("tests divisibility and alternates with ROW", () => {
      setContent([
        ["B1", 2024],
        ["A1", '=IF(MOD(B1, 4)=0, "leap", "common")'],
        ["C1", "=MOD(ROW(), 2)"],
        ["C2", "=MOD(ROW(), 2)"],
        ["C3", "=MOD(ROW(), 2)"],
      ]);

      expect(cell("A1")).toBe("leap");
      expect(cell("C1")).toBe(1);
      expect(cell("C2")).toBe(0);
      expect(cell("C3")).toBe(1);
    });

    test("recalculates when a referenced cell changes", () => {
      setContent([
        ["B1", 17],
        ["B2", 5],
        ["A1", "=MOD(B1, B2)"],
      ]);

      expect(cell("A1")).toBe(2);

      setCellContent("B1", 18);
      expect(cell("A1")).toBe(3);

      setCellContent("B2", -4);
      expect(cell("A1")).toBe(-2);

      setCellContent("B2", 0);
      expect(cell("A1")).toBe(FormulaError.DIV0);
    });

    test("is #DIV/0! for a divisor of 0", () => {
      setContent([
        ["B1", 0],
        ["A1", "=MOD(5, 0)"],
        ["A2", "=MOD(0, 0)"],
        ["A3", "=MOD(-5, B1)"],
        ["A4", "=MOD(INFINITY, 0)"],
        ["A5", '=IFERROR(MOD(5, 0), "none")'],
      ]);

      expect(cell("A1")).toBe(FormulaError.DIV0);
      expect(cell("A2")).toBe(FormulaError.DIV0);
      expect(cell("A3")).toBe(FormulaError.DIV0);
      expect(cell("A4")).toBe(FormulaError.DIV0);
      expect(cell("A5")).toBe("none");
    });

    test("requires numbers", () => {
      setContent([
        ["B1", "7"],
        // B2 is empty
        ["A1", '=MOD("7", 2)'],
        ["A2", '=MOD(7, "2")'],
        ["A3", "=MOD(TRUE, 2)"],
        ["A4", "=MOD(7, FALSE)"],
        ["A5", "=MOD(B1, 2)"],
        ["A6", "=MOD(B2, 2)"],
        // Excel: #DIV/0!, an empty divisor is 0 there
        ["A7", "=MOD(7, B2)"],
        ['A8', '=MOD("", 2)'],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8"]) {
        expect(cell(ref)).toBe(FormulaError.VALUE);
      }
    });

    test("is #NUM! for an infinite argument", () => {
      setContent([
        ["A1", "=MOD(INFINITY, 2)"],
        ["A2", "=MOD(-INFINITY, 2)"],
        ["A3", "=MOD(5, INFINITY)"],
        ["A4", "=MOD(5, -INFINITY)"],
        ["A5", "=MOD(INFINITY, INFINITY)"],
        // division by zero is INFINITY in this engine
        ["A6", "=MOD(1/0, 2)"],
        // what the engine computes for Excel's definition n - d*INT(n/d)
        ["B1", "=INFINITY-2*INT(INFINITY/2)"],
        ["B3", "=5-INFINITY*INT(5/INFINITY)"],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5", "A6", "B1", "B3"]) {
        expect(cell(ref)).toBe(FormulaError.NUM);
      }
    });

    test("is #NUM! for a quotient of 2^53 or more", () => {
      setContent([
        ["A1", "=MOD(2^53, 1)"],
        ["A2", "=MOD(-(2^53), 1)"],
        ["A3", "=MOD(1E300, 7)"],
        ["A4", "=MOD(1, 2^-60)"],
        ["A5", "=MOD(1E300, 1E-300)"],
        // still exact below the limit
        ["B1", "=MOD(2^53-1, 2)"],
        ["B2", "=MOD(2^52+1, 2)"],
        ["B3", "=MOD(1E15+7, 10)"],
      ]);

      for (const ref of ["A1", "A2", "A3", "A4", "A5"]) {
        expect(cell(ref)).toBe(FormulaError.NUM);
      }
      expect(cell("B1")).toBe(1);
      expect(cell("B2")).toBe(1);
      expect(cell("B3")).toBe(7);
    });

    test("propagates errors, the number's first", () => {
      setContent([
        ["B1", '=1+"x"'],
        ["A1", "=MOD(B1, 2)"],
        ["A2", "=MOD(NOSUCHFN(), 2)"],
        ["A3", "=MOD(5, NOSUCHFN())"],
        ["A4", "=MOD(NOSUCHFN(), B1)"],
        ["A5", "=MOD(#REF!, 2)"],
        ["A6", '=MOD(MATCH("z", {"x","y"}, 0), 2)'],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.NAME);
      expect(cell("A3")).toBe(FormulaError.NAME);
      expect(cell("A4")).toBe(FormulaError.NAME);
      expect(cell("A5")).toBe(FormulaError.REF);
      expect(cell("A6")).toBe(FormulaError.NA);
    });

    test("reads an omitted argument as 0", () => {
      setContent([
        ["A1", "=MOD(, 3)"],
        ["A2", "=MOD(5, )"],
        ["A3", "=MOD(, )"],
      ]);

      expect(cell("A1")).toBe(0);
      expect(cell("A2")).toBe(FormulaError.DIV0);
      expect(cell("A3")).toBe(FormulaError.DIV0);
    });

    test("requires exactly two arguments", () => {
      setContent([
        ["A1", "=MOD()"],
        ["A2", "=MOD(5)"],
        ["A3", "=MOD(5, 2, 1)"],
      ]);

      expect(cell("A1")).toBe(FormulaError.VALUE);
      expect(cell("A2")).toBe(FormulaError.VALUE);
      expect(cell("A3")).toBe(FormulaError.VALUE);
    });

    test("is case-insensitive", () => {
      setContent([["A1", "=mod(7, 4)"]]);

      expect(cell("A1")).toBe(3);
    });

    test("keeps the remainder between 0 and the divisor, on the divisor's side", () => {
      const divisors = [0.1, -0.3, 0.07, 2.5, -7];
      const numbers: number[] = [];
      for (let i = -1000; i <= 1000; i += 7) {
        numbers.push(i / 100);
      }
      const cases = divisors.flatMap((divisor, column) =>
        numbers.map((number, row) => ({
          ref: `${String.fromCharCode(65 + column)}${row + 1}`,
          number,
          divisor,
        }))
      );
      setContent(
        cases.map(({ ref, number, divisor }) => [
          ref,
          `=MOD(${number}, ${divisor})`,
        ])
      );

      for (const { ref, number, divisor } of cases) {
        const remainder = cell(ref) as number;
        expect(typeof remainder).toBe("number");
        if (divisor > 0) {
          expect(remainder >= 0 && remainder < divisor).toBe(true);
        } else {
          expect(remainder <= 0 && remainder > divisor).toBe(true);
        }
        // number - remainder is a whole multiple of the divisor
        const multiple = (number - remainder) / divisor;
        expect(Math.abs(multiple - Math.round(multiple))).toBeLessThan(1e-9);
      }
    });
  });

  describe("arrays", () => {
    test("repeats a cycle over a sequence", () => {
      setContent([
        ["A1", "=MOD(SEQUENCE(6), 3)"],
        ["B1", "=MOD(SEQUENCE(6, 1, 0), 3)+1"],
      ]);

      expect(["A1", "A2", "A3", "A4", "A5", "A6"].map(cell)).toEqual([
        1, 2, 0, 1, 2, 0,
      ]);
      expect(["B1", "B2", "B3", "B4", "B5", "B6"].map(cell)).toEqual([
        1, 2, 3, 1, 2, 3,
      ]);
    });

    test("divides a range element by element and follows edits", () => {
      setContent([
        ["A1", 5],
        ["A2", -5],
        ["A3", 6],
        ["A4", 7.5],
        ["C1", "=MOD(A1:A4, 3)"],
      ]);

      expect(["C1", "C2", "C3", "C4"].map(cell)).toEqual([2, 1, 0, 1.5]);

      setCellContent("A3", 8);
      expect(cell("C3")).toBe(2);
    });

    test("broadcasts the number and the divisor against each other", () => {
      setContent([
        ["A1", "=MOD(10, {3,4,6})"],
        ["A3", "=MOD({7;8}, {2,3})"],
        ["A6", "=MOD({1;2;3}, {2;3})"],
      ]);

      expect(["A1", "B1", "C1"].map(cell)).toEqual([1, 2, 4]);
      expect(["A3", "B3", "A4", "B4"].map(cell)).toEqual([1, 1, 0, 2]);
      // cells beyond the shorter array are #N/A
      expect(["A6", "A7", "A8"].map(cell)).toEqual([1, 2, FormulaError.NA]);
    });

    test("keeps an error in the element it belongs to", () => {
      setContent([
        ["A1", '=MOD({4;"x";9}, 2)'],
        ["B1", "=MOD(9, {2;0;4})"],
      ]);

      expect(["A1", "A2", "A3"].map(cell)).toEqual([0, FormulaError.VALUE, 1]);
      expect(["B1", "B2", "B3"].map(cell)).toEqual([1, FormulaError.DIV0, 1]);
    });

    test("selects every other row with IF and FILTER", () => {
      setContent([
        ["A1", 10],
        ["A2", 20],
        ["A3", 30],
        ["A4", 40],
        ["C1", '=IF(MOD(SEQUENCE(4), 2)=0, "even", "odd")'],
        ["D1", "=FILTER(A1:A4, MOD(SEQUENCE(4), 2)=0)"],
        ["E1", "=SUM(FILTER(A1:A4, MOD(SEQUENCE(4), 2)=0))"],
      ]);

      expect(["C1", "C2", "C3", "C4"].map(cell)).toEqual([
        "odd",
        "even",
        "odd",
        "even",
      ]);
      expect(["D1", "D2"].map(cell)).toEqual([20, 40]);
      expect(cell("E1")).toBe(60);
    });

    test("is read by aggregates", () => {
      setContent([
        ["A1", 5],
        ["A2", 7],
        // A3 is empty
        ["A4", 9],
        ["C1", "=SUM(MOD(SEQUENCE(10), 2))"],
        ["C2", "=SUM(MOD(A1:A4, 4))"],
        ["C3", "=MAX(MOD(A1:A4, 4))"],
        ["C4", "=SUM(MOD(A:A, 4))"],
        ["E1", "=MOD(A1:A4, 4)"],
      ]);

      expect(cell("C1")).toBe(5);
      // aggregates read the populated cells: 1 + 3 + 1
      expect(cell("C2")).toBe(5);
      expect(cell("C3")).toBe(3);
      expect(cell("C4")).toBe(5);
      // the spill shows the empty cell's #VALUE!, like INT and ROUND
      expect(["E1", "E2", "E3", "E4"].map(cell)).toEqual([
        1,
        3,
        FormulaError.VALUE,
        1,
      ]);
    });

    test("returns a scalar error argument instead of spilling", () => {
      setContent([["A1", "=MOD(SEQUENCE(3), NOSUCHFN())"]]);

      expect(cell("A1")).toBe(FormulaError.NAME);
      // nothing spilled into the cell below
      expect(cell("A2")).toBe("");
    });
  });
});
